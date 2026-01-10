# 03_teacher_only_watch_clips_BYTAG.R
# ============================================================
# Goal: pick review clips that are "teacher-on-screen" as much as possible
# Input:
#   - exports/wrist_tracks/wrist_tracks_<TAG>.csv  (preferred)
#   - exports/<TAG>/wrist_tracks_<TAG>.csv         (fallback)
# Output (ALL under exports/<TAG>/ and review_clips/<TAG>/):
#   - exports/<TAG>/watch_points_<TAG>_teacher.csv
#   - exports/<TAG>/teacher_qc_<TAG>.csv
#   - review_clips/<TAG>/cut_clips_teacher_focus.bat
#
# Key idea:
#   Use PoseLandmarker shoulder visibility (ls_vis / rs_vis) to define teacher_present.
#   This filters out student close-ups / student hand close-ups where shoulders are missing.
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
})

# ----------------------------
# (0) USER SETTINGS (PUBLIC TEMPLATE)
# ----------------------------

get_env_or <- function(key, default) {
  v <- Sys.getenv(key, unset = "")
  if (nzchar(v)) v else default
}

# ---- identifiers (safe to commit) ----
TAG <- get_env_or("GESTURE_TAG", "mXX")

# ---- basic parameters ----
fps <- 30

# ---- project roots ----
# Expected structure under project_root:
#   exports/
#   segments/
#   review_clips/
project_root <- get_env_or(
  "GESTURE_ROOT",
  "<PATH_TO_PROJECT_ROOT>"
)

exports_root <- file.path(project_root, "exports")
segments_dir <- file.path(project_root, "segments")
clips_root   <- file.path(project_root, "review_clips")

# ----------------------------
# (1) PARAMETERS YOU MAY TUNE
# ----------------------------

# A) Teacher-present definition (most important)
conf_min <- 0.35          # YOLO conf too low -> probably wrong person/garbage box
shoulder_vis_min <- 0.45  # ls_vis/rs_vis threshold (0-1). Higher = stricter teacher
shoulder_ok_mode <- "both"
# "both" = require BOTH shoulders visible (stricter, fewer false positives)
# "either" = require at least one shoulder visible (looser, catches more teacher)

# B) Motion signal (teacher-only)
speed_q   <- 0.97         # quantile for peak candidates (higher => fewer, more "big gesture" only)
min_gap_s <- 2.0          # cooldown: within 2s keep only the strongest peak
MAX_KEEP  <- 20           # final number of clips you want

# C) Clip window
WINDOW_BEFORE <- 2
WINDOW_AFTER  <- 2

# D) Teacher ratio filter inside the whole clip window
TEACHER_RATIO_MIN <- 0.80  # raise to 0.85 if still too many students; fallback exists

# E) ffmpeg cutting
ffmpeg_crf    <- 23
ffmpeg_preset <- "veryfast"

# ----------------------------
# (2) READ + ENSURE frame / t_sec
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE)

if (!("frame" %in% names(dat))) dat <- dat %>% mutate(frame = row_number() - 1)
if (!("t_sec" %in% names(dat))) dat <- dat %>% mutate(t_sec = frame / fps)

# ----------------------------
# (3) COMPUTE HAND SPEED (dominant-hand agnostic)
# ----------------------------
dat2 <- dat %>%
  arrange(frame) %>%
  mutate(
    lw_dx = lw_x - lag(lw_x),
    lw_dy = lw_y - lag(lw_y),
    lw_disp = sqrt(lw_dx^2 + lw_dy^2),
    lw_speed = lw_disp * fps,
    
    rw_dx = rw_x - lag(rw_x),
    rw_dy = rw_y - lag(rw_y),
    rw_disp = sqrt(rw_dx^2 + rw_dy^2),
    rw_speed = rw_disp * fps,
    
    hand_speed = pmax(lw_speed, rw_speed, na.rm = TRUE),
    hand_speed = ifelse(is.infinite(hand_speed), NA_real_, hand_speed)
  )

# ----------------------------
# (4) DEFINE teacher_present USING SHOULDER VISIBILITY
# ----------------------------
dat2 <- dat2 %>%
  mutate(
    conf_ok = !is.na(conf_det) & conf_det >= conf_min,
    
    ls_ok = !is.na(ls_vis) & (ls_vis >= shoulder_vis_min),
    rs_ok = !is.na(rs_vis) & (rs_vis >= shoulder_vis_min),
    
    shoulders_ok = case_when(
      shoulder_ok_mode == "both"   ~ (ls_ok & rs_ok),
      shoulder_ok_mode == "either" ~ (ls_ok | rs_ok),
      TRUE                         ~ (ls_ok & rs_ok)
    ),
    
    teacher_present = conf_ok & shoulders_ok
  )

# ----------------------------
# (5) TEACHER-ONLY SPEED SERIES (mask out non-teacher frames)
# ----------------------------
dat2 <- dat2 %>%
  mutate(speed_teacher = ifelse(teacher_present, hand_speed, NA_real_))

# ----------------------------
# (6) PEAK CANDIDATES (high-speed moments among teacher frames)
# ----------------------------
thr <- quantile(dat2$speed_teacher, speed_q, na.rm = TRUE)

cand <- dat2 %>%
  filter(!is.na(speed_teacher) & speed_teacher >= thr) %>%
  select(frame, t_sec, speed_teacher) %>%
  arrange(desc(speed_teacher))

if (nrow(cand) == 0) {
  message("No candidates at q=", speed_q, ". Relaxing to q=0.95 ...")
  thr <- quantile(dat2$speed_teacher, 0.95, na.rm = TRUE)
  cand <- dat2 %>%
    filter(!is.na(speed_teacher) & speed_teacher >= thr) %>%
    select(frame, t_sec, speed_teacher) %>%
    arrange(desc(speed_teacher))
}

# ----------------------------
# (7) COOLDOWN SELECTION (event-level dedupe)
# ----------------------------
kept <- cand[0, ]
kept_times <- c()

for (i in seq_len(nrow(cand))) {
  t <- cand$t_sec[i]
  if (length(kept_times) == 0 || all(abs(t - kept_times) >= min_gap_s)) {
    kept <- bind_rows(kept, cand[i, ])
    kept_times <- c(kept_times, t)
  }
  if (nrow(kept) >= 300) break
}

watch_tbl <- kept %>%
  mutate(
    time_s = t_sec,
    time_start = pmax(time_s - WINDOW_BEFORE, 0),
    time_end   = time_s + WINDOW_AFTER
  ) %>%
  arrange(time_s) %>%
  rename(speed = speed_teacher) %>%
  select(frame, time_s, time_start, time_end, speed)

# ----------------------------
# (8) TEACHER RATIO FILTER (inside full clip window)
# ----------------------------
clip_teacher_ratio <- function(t0, t1) {
  seg <- dat2 %>% filter(t_sec >= t0, t_sec <= t1)
  if (nrow(seg) == 0) return(0)
  mean(seg$teacher_present, na.rm = TRUE)
}

watch_tbl2 <- watch_tbl %>%
  rowwise() %>%
  mutate(teacher_ratio = clip_teacher_ratio(time_start, time_end)) %>%
  ungroup() %>%
  arrange(desc(speed))

watch_keep <- watch_tbl2 %>%
  filter(teacher_ratio >= TEACHER_RATIO_MIN) %>%
  slice(1:MAX_KEEP) %>%
  arrange(time_s)

# Fallback: relax teacher ratio if not enough clips
if (nrow(watch_keep) < MAX_KEEP) {
  message("Not enough clips at teacher_ratio >= ", TEACHER_RATIO_MIN, ". Relax to 0.60 ...")
  watch_keep <- watch_tbl2 %>%
    filter(teacher_ratio >= 0.60) %>%
    slice(1:MAX_KEEP) %>%
    arrange(time_s)
}

if (nrow(watch_keep) == 0) {
  stop("0 clips kept. You likely set shoulder_vis_min too high or pose landmarks miss shoulders often.")
}

write_csv(watch_keep, out_watch_csv)
message("Saved teacher-focused watch list: ", out_watch_csv)

# ----------------------------
# (9) GENERATE FFMPEG .BAT (RE-ENCODE, accurate cutting)
# ----------------------------
cmd_tbl <- watch_keep %>%
  mutate(
    out_mp4 = file.path(
      clips_dir,
      sprintf("clip_%03d_%06.2fs_tr%.2f.mp4", row_number(), time_s, teacher_ratio)
    ),
    ffmpeg_cmd = sprintf(
      'ffmpeg -y -ss %.3f -to %.3f -i "%s" -c:v libx264 -preset %s -crf %d -c:a aac -b:a 128k "%s"',
      time_start, time_end, video_path, ffmpeg_preset, ffmpeg_crf, out_mp4
    )
  )

writeLines(cmd_tbl$ffmpeg_cmd, bat_path)
message("Saved ffmpeg batch: ", bat_path)
message("Run the .bat to generate teacher-focused clips.")

# ----------------------------
# (10) QC SUMMARY (quick diagnosis)
# ----------------------------
qc <- dat2 %>%
  summarise(
    TAG = TAG,
    total_frames = n(),
    teacher_present_ratio = mean(teacher_present, na.rm = TRUE),
    conf_ok_ratio = mean(conf_ok, na.rm = TRUE),
    shoulders_ok_ratio = mean(shoulders_ok, na.rm = TRUE),
    speed_thr = thr,
    n_candidates = nrow(cand),
    n_watch_final = nrow(watch_keep),
    mean_teacher_ratio_in_clips = mean(watch_keep$teacher_ratio, na.rm = TRUE)
  )

write_csv(qc, out_qc_csv)
print(qc)

cat("\nSaved:\n",
    "watch :", out_watch_csv, "\n",
    "bat   :", bat_path, "\n",
    "qc    :", out_qc_csv, "\n",
    "tag_dir :", tag_dir, "\n",
    "clips_dir :", clips_dir, "\n")

