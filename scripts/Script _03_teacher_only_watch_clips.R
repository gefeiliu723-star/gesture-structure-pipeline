## ============================================================
# Script 03 (GITHUB-READY): TEACHER-FIRST + EVENT-ANCHORED (+ STRONG MODE)
#
# Goal:
#   Pick review clips centered on gesture_events (Script 02), then filter by teacher_ratio.
#   Supports 3 selection modes:
#     - "best"   : most teacher-like (highest teacher_ratio), then cooldown
#     - "uniform": time-stratified; per-bin best teacher_ratio; optional strict + backfill
#     - "strong" : strongest events by kinematic peak strength near event time (teacher-only), then cooldown
#
# INPUTS (repo-relative):
#   - exports/<TAG>/gesture_events_<TAG>.csv
#   - exports/wrist_tracks/wrist_tracks_<TAG>.csv   (preferred)
#       OR exports/<TAG>/wrist_tracks_<TAG>.csv     (fallback)
#   - segments/<TAG>*.mp4   (a 30fps segment file)
#
# OUTPUTS (repo-relative; created automatically):
#   - exports/<TAG>/watch_points_<TAG>_teacher.csv
#   - exports/<TAG>/teacher_qc_<TAG>.csv
#   - review_clips/<TAG>/cut_clips_teacher_focus.bat
#
# NOTE:
#   - This script does NOT run ffmpeg. It writes a .bat with commands.
#   - Requires ffmpeg installed and available in PATH when running the .bat.
## ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# (0) USER SETTINGS
# ----------------------------
TAG <- "m05"
fps <- 30

# Selection behavior: "uniform" / "best" / "strong"
SELECT_MODE <- "strong"

# Teacher-first switches
TEACHER_FIRST  <- TRUE
UNIFORM_STRICT <- TRUE
BACKFILL_MODE  <- "best"  # "best" or "none"

# Teacher thresholds ladder (auto relax when TEACHER_FIRST=TRUE)
TEACHER_THRESHOLDS <- c(0.90, 0.85, 0.80, 0.70, 0.60)

# Strength (used in SELECT_MODE="strong"; computed anyway for QC)
STRENGTH_WINDOW <- 0.25   # seconds around event time for peak strength

# Teacher presence definition
conf_min <- 0.35
shoulder_vis_min <- 0.45
shoulder_ok_mode <- "both"   # "both" or "either"

# Selection constraints
min_gap_s <- 2.0
MAX_KEEP  <- 20

WINDOW_BEFORE <- 2
WINDOW_AFTER  <- 2

# Only used when TEACHER_FIRST = FALSE (simple mode)
TEACHER_RATIO_MIN <- 0.80

# ffmpeg command parameters (written into bat)
ffmpeg_crf    <- 23
ffmpeg_preset <- "veryfast"

# ----------------------------
# (0.1) PROJECT ROOT (GITHUB-READY)
# ----------------------------
# CN: 优先用环境变量 MKAP_ROOT；没有就用当前工作目录（建议 setwd 到 repo 根）
# EN: Use MKAP_ROOT if set; otherwise use getwd() (run from repo root).
project_root <- Sys.getenv("MKAP_ROOT", unset = getwd())

exports_root <- file.path(project_root, "exports")
segments_dir <- file.path(project_root, "segments")
clips_root   <- file.path(project_root, "review_clips")

tag_dir <- file.path(exports_root, TAG)
dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)

clips_dir <- file.path(clips_root, TAG)
dir.create(clips_dir, showWarnings = FALSE, recursive = TRUE)

events_path <- file.path(tag_dir, paste0("gesture_events_", TAG, ".csv"))

wrist_candidates <- c(
  file.path(exports_root, "wrist_tracks", paste0("wrist_tracks_", TAG, ".csv")),
  file.path(tag_dir, paste0("wrist_tracks_", TAG, ".csv"))
)
csv_path <- wrist_candidates[file.exists(wrist_candidates)][1]
if (is.na(csv_path) || !nzchar(csv_path)) csv_path <- wrist_candidates[1]

out_watch_csv <- file.path(tag_dir, paste0("watch_points_", TAG, "_teacher.csv"))
out_qc_csv    <- file.path(tag_dir, paste0("teacher_qc_", TAG, ".csv"))
bat_path      <- file.path(clips_dir, "cut_clips_teacher_focus.bat")

# Find the newest segment mp4 matching TAG prefix
cand_mp4 <- list.files(
  segments_dir,
  pattern = paste0("^", TAG, ".*\\.mp4$"),
  full.names = TRUE,
  ignore.case = TRUE
)
if (length(cand_mp4) == 0) stop("No segment mp4 found under: ", segments_dir)
video_path <- cand_mp4[which.max(file.info(cand_mp4)$mtime)]

# validate inputs
if (!file.exists(events_path)) stop("Missing gesture events CSV: ", events_path)
if (!file.exists(csv_path))     stop("Missing wrist tracks CSV: tried ", paste(wrist_candidates, collapse = " | "))
if (!file.exists(video_path))   stop("Missing segment mp4: ", video_path)

message("Project root      : ", normalizePath(project_root, winslash = "/", mustWork = FALSE))
message("Using video_path  : ", normalizePath(video_path,   winslash = "/", mustWork = FALSE))
message("Using events_path : ", normalizePath(events_path,  winslash = "/", mustWork = FALSE))
message("Using wrist_csv   : ", normalizePath(csv_path,     winslash = "/", mustWork = FALSE))

# ----------------------------
# (1) READ wrist_tracks + teacher_present
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE)
if (!("frame" %in% names(dat))) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!("t_sec" %in% names(dat))) dat <- dat %>% mutate(t_sec = frame / fps)

# Required columns for teacher_present
req_cols <- c("conf_det", "ls_vis", "rs_vis")
miss_cols <- setdiff(req_cols, names(dat))
if (length(miss_cols) > 0) {
  stop("wrist_tracks is missing required columns: ", paste(miss_cols, collapse = ", "),
       "\nExpected: conf_det, ls_vis, rs_vis (from Script 01).")
}

dat2 <- dat %>%
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
# (1.1) hand_speed + speed_teacher (for strength)
# ----------------------------
if (all(c("lw_x","lw_y","rw_x","rw_y") %in% names(dat2))) {
  dat2 <- dat2 %>%
    arrange(frame) %>%
    mutate(
      lw_dx = lw_x - lag(lw_x),
      lw_dy = lw_y - lag(lw_y),
      lw_speed = sqrt(lw_dx^2 + lw_dy^2) * fps,
      
      rw_dx = rw_x - lag(rw_x),
      rw_dy = rw_y - lag(rw_y),
      rw_speed = sqrt(rw_dx^2 + rw_dy^2) * fps,
      
      hand_speed = pmax(lw_speed, rw_speed, na.rm = TRUE),
      hand_speed = ifelse(is.infinite(hand_speed), NA_real_, hand_speed),
      speed_teacher = ifelse(teacher_present, hand_speed, NA_real_)
    )
} else {
  # fallback if Script 01 already produced hand_speed
  if (!("hand_speed" %in% names(dat2))) dat2 <- dat2 %>% mutate(hand_speed = NA_real_)
  if (!("speed_teacher" %in% names(dat2))) {
    dat2 <- dat2 %>% mutate(speed_teacher = ifelse(teacher_present, hand_speed, NA_real_))
  }
}

clip_teacher_ratio <- function(t0, t1) {
  seg <- dat2 %>% filter(t_sec >= t0, t_sec <= t1)
  if (nrow(seg) == 0) return(0)
  mean(seg$teacher_present, na.rm = TRUE)
}

event_strength <- function(t_center, window_s = STRENGTH_WINDOW) {
  t0 <- t_center - window_s
  t1 <- t_center + window_s
  seg <- dat2 %>% filter(t_sec >= t0, t_sec <= t1)
  if (nrow(seg) == 0) return(NA_real_)
  s <- suppressWarnings(max(seg$speed_teacher, na.rm = TRUE))
  if (is.infinite(s)) s <- NA_real_
  s
}

# ----------------------------
# (2) READ events -> candidates (anchored)
# ----------------------------
ev <- read_csv(events_path, show_col_types = FALSE)
names(ev) <- tolower(names(ev))

time_col <- dplyr::case_when(
  "t_peak" %in% names(ev)     ~ "t_peak",
  "tpeak" %in% names(ev)      ~ "tpeak",
  "peak_sec" %in% names(ev)   ~ "peak_sec",
  "start_sec" %in% names(ev)  ~ "start_sec",
  "start_s" %in% names(ev)    ~ "start_s",
  TRUE                        ~ ""
)
if (time_col == "") stop("gesture_events has no time column (need t_peak/peak_sec/start_sec/etc).")

cand <- tibble(
  EventID = if ("eventid" %in% names(ev)) as.character(ev$eventid) else NA_character_,
  time_s  = suppressWarnings(as.numeric(ev[[time_col]]))
) %>%
  filter(is.finite(time_s)) %>%
  arrange(time_s)

if (nrow(cand) == 0) stop("No valid event times in gesture_events.")

cand2 <- cand %>%
  mutate(
    time_start = pmax(time_s - WINDOW_BEFORE, 0),
    time_end   = time_s + WINDOW_AFTER,
    frame      = as.integer(round(time_s * fps))
  ) %>%
  rowwise() %>%
  mutate(
    teacher_ratio = clip_teacher_ratio(time_start, time_end),
    strength = event_strength(time_s)
  ) %>%
  ungroup()

# ----------------------------
# (2.5) TEACHER-FIRST FILTER (ladder)
# ----------------------------
cand3 <- cand2

if (TEACHER_FIRST) {
  picked_thr <- NA_real_
  for (thr in TEACHER_THRESHOLDS) {
    tmp <- cand2 %>% filter(teacher_ratio >= thr)
    message("Teacher-first try thr=", thr, " -> n=", nrow(tmp))
    if (nrow(tmp) >= max(5, ceiling(MAX_KEEP * 0.6))) {
      cand3 <- tmp
      picked_thr <- thr
      break
    }
  }
  if (is.na(picked_thr)) {
    message("Teacher-first: too few even at lowest thr. Using best available (no hard filter).")
    cand3 <- cand2 %>% arrange(desc(teacher_ratio), time_s)
  } else {
    message("Teacher-first: using teacher_ratio >= ", picked_thr)
  }
} else {
  cand3 <- cand2 %>% filter(teacher_ratio >= TEACHER_RATIO_MIN)
  if (nrow(cand3) < MAX_KEEP) {
    message("Not enough at teacher_ratio >= ", TEACHER_RATIO_MIN, ". Relax to 0.60 ...")
    cand3 <- cand2 %>% filter(teacher_ratio >= 0.60)
  }
}

if (nrow(cand3) == 0) stop("0 candidates after teacher filtering.")

# ----------------------------
# (3) SELECT helpers
# ----------------------------
cooldown_pick <- function(df, max_keep, min_gap_s) {
  kept <- df[0, ]
  kept_times <- c()
  for (i in seq_len(nrow(df))) {
    t <- df$time_s[i]
    if (length(kept_times) == 0 || all(abs(t - kept_times) >= min_gap_s)) {
      kept <- bind_rows(kept, df[i, ])
      kept_times <- c(kept_times, t)
    }
    if (nrow(kept) >= max_keep) break
  }
  kept %>% arrange(time_s)
}

# ----------------------------
# (3.1) SELECT_MODE logic
# ----------------------------
if (SELECT_MODE == "best") {
  
  pool <- cand3 %>% arrange(desc(teacher_ratio), time_s)
  watch_keep <- cooldown_pick(pool, MAX_KEEP, min_gap_s)
  
} else if (SELECT_MODE == "strong") {
  
  pool <- cand3 %>% arrange(desc(strength), desc(teacher_ratio), time_s)
  watch_keep <- cooldown_pick(pool, MAX_KEEP, min_gap_s)
  
} else if (SELECT_MODE == "uniform") {
  
  tmin <- min(cand3$time_s, na.rm = TRUE)
  tmax <- max(cand3$time_s, na.rm = TRUE)
  
  breaks <- seq(tmin, tmax, length.out = MAX_KEEP + 1)
  breaks <- unique(as.numeric(breaks))
  
  if (length(breaks) < 3) {
    
    pool <- cand3 %>% arrange(desc(teacher_ratio), time_s)
    watch_keep <- cooldown_pick(pool, MAX_KEEP, min_gap_s)
    
  } else {
    
    cand3b <- cand3 %>%
      mutate(bin = cut(time_s, breaks = breaks, include.lowest = TRUE, right = TRUE))
    
    picked <- cand3b %>%
      filter(!is.na(bin)) %>%
      group_by(bin) %>%
      arrange(desc(teacher_ratio), time_s) %>%
      slice(1) %>%
      ungroup() %>%
      select(-bin)
    
    if (UNIFORM_STRICT) {
      strict_thr <- if (TEACHER_FIRST) max(min(TEACHER_THRESHOLDS), 0.70) else TEACHER_RATIO_MIN
      picked2 <- picked %>% filter(teacher_ratio >= strict_thr)
      message("Uniform strict: kept ", nrow(picked2), "/", nrow(picked), " bins at thr >= ", strict_thr)
      picked <- picked2
    }
    
    pool <- picked %>% arrange(desc(teacher_ratio), time_s)
    watch_keep <- cooldown_pick(pool, MAX_KEEP, min_gap_s)
    
    if (nrow(watch_keep) < MAX_KEEP && BACKFILL_MODE == "best") {
      need <- MAX_KEEP - nrow(watch_keep)
      
      # safer anti-join (only by EventID + time_s)
      pool2 <- cand3 %>%
        anti_join(watch_keep, by = c("EventID", "time_s")) %>%
        arrange(desc(teacher_ratio), time_s)
      
      add <- cooldown_pick(pool2, need, min_gap_s)
      watch_keep <- bind_rows(watch_keep, add) %>% arrange(time_s)
      message("Backfill added: ", nrow(add), " -> final ", nrow(watch_keep))
    }
  }
  
} else {
  stop("SELECT_MODE must be 'uniform', 'best', or 'strong'.")
}

if (nrow(watch_keep) == 0) stop("0 clips kept after selection.")

# ----------------------------
# (4) OUTPUT watch points
# speed column carries event strength (peak magnitude)
# ----------------------------
watch_keep_out <- watch_keep %>%
  mutate(speed = strength) %>%
  select(frame, time_s, time_start, time_end, speed, teacher_ratio, EventID)

write_csv(watch_keep_out, out_watch_csv)
message("Saved watch list: ", out_watch_csv)

# ----------------------------
# (5) FFMPEG BAT
# ----------------------------
cmd_tbl <- watch_keep_out %>%
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

# ----------------------------
# (6) QC summary
# ----------------------------
qc <- dat2 %>%
  summarise(
    TAG = TAG,
    total_frames = n(),
    teacher_present_ratio = mean(teacher_present, na.rm = TRUE),
    conf_ok_ratio = mean(conf_ok, na.rm = TRUE),
    shoulders_ok_ratio = mean(shoulders_ok, na.rm = TRUE),
    hand_speed_available = mean(!is.na(hand_speed)),
    speed_teacher_available = mean(!is.na(speed_teacher))
  ) %>%
  mutate(
    select_mode = SELECT_MODE,
    teacher_first = TEACHER_FIRST,
    uniform_strict = UNIFORM_STRICT,
    backfill_mode = BACKFILL_MODE,
    teacher_thresholds = paste(TEACHER_THRESHOLDS, collapse = ","),
    strength_window_s = STRENGTH_WINDOW,
    gesture_events_path = events_path,
    wrist_tracks_path   = csv_path,
    time_col_used       = time_col,
    n_events_total      = nrow(cand2),
    n_events_candidate  = nrow(cand3),
    n_watch_final       = nrow(watch_keep_out),
    mean_teacher_ratio_in_clips = mean(watch_keep_out$teacher_ratio, na.rm = TRUE),
    mean_strength_in_clips = mean(watch_keep_out$speed, na.rm = TRUE),
    max_strength_in_clips  = suppressWarnings(max(watch_keep_out$speed, na.rm = TRUE)),
    min_time_s = min(watch_keep_out$time_s),
    max_time_s = max(watch_keep_out$time_s)
  )

write_csv(qc, out_qc_csv)
print(qc)

cat("\nSaved:\n",
    "watch :", out_watch_csv, "\n",
    "bat   :", bat_path, "\n",
    "qc    :", out_qc_csv, "\n",
    "tag_dir :", tag_dir, "\n",
    "clips_dir :", clips_dir, "\n")

