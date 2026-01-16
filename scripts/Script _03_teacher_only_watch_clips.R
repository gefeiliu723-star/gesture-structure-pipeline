## ============================================================
# Script 03 (CLEAN / GITHUB-READY):
# TEACHER-FIRST + EVENT-ANCHORED (+ STRONG MODE)
#
# INPUTS (repo-relative):
#   - exports/<TAG>/gesture_events_<TAG>.csv
#   - exports/wrist_tracks/wrist_tracks_<TAG>.csv   (preferred)
#       OR exports/<TAG>/wrist_tracks_<TAG>.csv     (fallback)
#   - segments/<TAG>*.mp4   (a 30fps segment file)
#
# OUTPUTS (repo-relative):
#   - exports/<TAG>/watch_points_<TAG>_teacher.csv
#   - exports/<TAG>/teacher_qc_<TAG>.csv
#   - review_clips/<TAG>/cut_clips_teacher_focus.bat
#
# CLI:
#   Rscript scripts/03_teacher_only_watchpoints_EVENT_ANCHORED.R <TAG> [fps] [SELECT_MODE]
# Example:
#   Rscript scripts/03_teacher_only_watchpoints_EVENT_ANCHORED.R m05 30 strong
#
# ROOT:
#   MKAP_ROOT env var; fallback getwd() (run from repo root)
## ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# (0) SETTINGS / CLI (TAG REQUIRED)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 1 || !nzchar(args[[1]])) {
  stop(
    "Missing TAG.\n",
    "Usage:\n  Rscript scripts/03_teacher_only_watchpoints_EVENT_ANCHORED.R <TAG> [fps] [SELECT_MODE]\n",
    call. = FALSE
  )
}

TAG <- args[[1]]
fps <- if (length(args) >= 2 && nzchar(args[[2]])) as.integer(args[[2]]) else 30L
SELECT_MODE <- if (length(args) >= 3 && nzchar(args[[3]])) tolower(args[[3]]) else "strong"
stopifnot(is.finite(fps), fps > 0)
if (!(SELECT_MODE %in% c("uniform","best","strong"))) {
  stop("SELECT_MODE must be one of: uniform, best, strong", call. = FALSE)
}

# Teacher-first switches
TEACHER_FIRST  <- TRUE
UNIFORM_STRICT <- TRUE
BACKFILL_MODE  <- "best"  # "best" or "none"

# Teacher thresholds ladder (auto relax when TEACHER_FIRST=TRUE)
TEACHER_THRESHOLDS <- c(0.90, 0.85, 0.80, 0.70, 0.60)

# Strength (used in SELECT_MODE="strong")
STRENGTH_WINDOW <- 0.25   # seconds around event time

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

# ffmpeg params (written into bat)
ffmpeg_crf    <- 23
ffmpeg_preset <- "veryfast"

# ----------------------------
# (0.1) PROJECT ROOT (portable)
# ----------------------------
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
wrist_csv <- wrist_candidates[file.exists(wrist_candidates)][1]
if (is.na(wrist_csv) || !nzchar(wrist_csv)) {
  stop("Missing wrist tracks CSV. Tried:\n  - ",
       paste(wrist_candidates, collapse = "\n  - "),
       call. = FALSE)
}

out_watch_csv <- file.path(tag_dir, paste0("watch_points_", TAG, "_teacher.csv"))
out_qc_csv    <- file.path(tag_dir, paste0("teacher_qc_", TAG, ".csv"))
bat_path      <- file.path(clips_dir, "cut_clips_teacher_focus.bat")

cand_mp4 <- list.files(
  segments_dir,
  pattern = paste0("^", TAG, ".*\\.mp4$"),
  full.names = TRUE,
  ignore.case = TRUE
)
if (length(cand_mp4) == 0) stop("No segment mp4 found under: ", segments_dir, call. = FALSE)
video_path <- cand_mp4[which.max(file.info(cand_mp4)$mtime)]

if (!file.exists(events_path)) stop("Missing gesture events CSV: ", events_path, call. = FALSE)
if (!file.exists(video_path))  stop("Missing segment mp4: ", video_path, call. = FALSE)

message("Root       : ", normalizePath(project_root, winslash = "/", mustWork = FALSE))
message("Video      : ", normalizePath(video_path,   winslash = "/", mustWork = FALSE))
message("Events CSV : ", normalizePath(events_path,  winslash = "/", mustWork = FALSE))
message("Wrist CSV  : ", normalizePath(wrist_csv,    winslash = "/", mustWork = FALSE))

# ----------------------------
# (1) HELPERS
# ----------------------------
std_names <- function(nms) {
  nms %>%
    str_replace_all("\\.+", "_") %>%
    str_replace_all("[^A-Za-z0-9_]", "_") %>%
    str_replace_all("_+", "_") %>%
    str_replace_all("^_|_$", "") %>%
    tolower()
}

pick_col <- function(df, candidates) {
  nms <- names(df)
  low <- tolower(nms)
  for (cand in candidates) {
    j <- which(low == tolower(cand))
    if (length(j) >= 1) return(nms[j[[1]]])
  }
  NA_character_
}

as_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

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
# (2) READ wrist_tracks + teacher_present
# ----------------------------
dat <- readr::read_csv(wrist_csv, show_col_types = FALSE)
names(dat) <- std_names(names(dat))

if (!("frame" %in% names(dat))) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!("t_sec" %in% names(dat))) dat <- dat %>% mutate(t_sec = frame / fps)

col_conf <- pick_col(dat, c("conf_det","conf","det_conf"))
col_ls   <- pick_col(dat, c("ls_vis","left_shoulder_vis","ls_visible"))
col_rs   <- pick_col(dat, c("rs_vis","right_shoulder_vis","rs_visible"))

miss <- c()
if (is.na(col_conf)) miss <- c(miss, "conf_det/conf/det_conf")
if (is.na(col_ls))   miss <- c(miss, "ls_vis/left_shoulder_vis")
if (is.na(col_rs))   miss <- c(miss, "rs_vis/right_shoulder_vis")
if (length(miss) > 0) {
  stop("wrist_tracks missing required columns: ", paste(miss, collapse = ", "),
       "\nExpected these to exist (from Script 01).",
       call. = FALSE)
}

dat2 <- dat %>%
  mutate(
    conf_det = as_num(.data[[col_conf]]),
    ls_vis   = as_num(.data[[col_ls]]),
    rs_vis   = as_num(.data[[col_rs]]),
    
    conf_ok = is.finite(conf_det) & conf_det >= conf_min,
    ls_ok   = is.finite(ls_vis) & (ls_vis >= shoulder_vis_min),
    rs_ok   = is.finite(rs_vis) & (rs_vis >= shoulder_vis_min),
    shoulders_ok = case_when(
      shoulder_ok_mode == "both"   ~ (ls_ok & rs_ok),
      shoulder_ok_mode == "either" ~ (ls_ok | rs_ok),
      TRUE                         ~ (ls_ok & rs_ok)
    ),
    teacher_present = conf_ok & shoulders_ok
  )

# ----------------------------
# (2.1) hand_speed + speed_teacher (for strength)
# ----------------------------
# If lw/rw coords exist, compute; else rely on precomputed hand_speed if present.
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
  seg <- dat2 %>% filter(t_sec >= (t_center - window_s), t_sec <= (t_center + window_s))
  if (nrow(seg) == 0) return(NA_real_)
  s <- suppressWarnings(max(seg$speed_teacher, na.rm = TRUE))
  if (is.infinite(s)) s <- NA_real_
  s
}

# ----------------------------
# (3) READ events -> anchored candidates
# ----------------------------
ev <- readr::read_csv(events_path, show_col_types = FALSE)
names(ev) <- std_names(names(ev))

col_time <- pick_col(ev, c("t_peak","peak_sec","t_peak_sec","peak","start_sec","time_s","t_sec"))
col_eid  <- pick_col(ev, c("eventid","event_id","id"))
if (is.na(col_time)) stop("gesture_events has no usable time column (t_peak/peak_sec/start_sec/...).", call. = FALSE)

cand <- tibble(
  EventID = if (!is.na(col_eid)) as.character(ev[[col_eid]]) else NA_character_,
  time_s  = as_num(ev[[col_time]])
) %>%
  filter(is.finite(time_s)) %>%
  arrange(time_s)

if (nrow(cand) == 0) stop("No valid event times in gesture_events.", call. = FALSE)

cand2 <- cand %>%
  mutate(
    time_start = pmax(time_s - WINDOW_BEFORE, 0),
    time_end   = time_s + WINDOW_AFTER,
    frame      = as.integer(round(time_s * fps))
  ) %>%
  rowwise() %>%
  mutate(
    teacher_ratio = clip_teacher_ratio(time_start, time_end),
    strength      = event_strength(time_s)
  ) %>%
  ungroup()

# ----------------------------
# (4) TEACHER-FIRST FILTER (ladder)
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

if (nrow(cand3) == 0) stop("0 candidates after teacher filtering.", call. = FALSE)

# ----------------------------
# (5) SELECT_MODE logic
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
  
  breaks <- unique(as.numeric(seq(tmin, tmax, length.out = MAX_KEEP + 1)))
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
      
      pool2 <- cand3 %>%
        anti_join(watch_keep, by = c("EventID", "time_s")) %>%
        arrange(desc(teacher_ratio), time_s)
      
      add <- cooldown_pick(pool2, need, min_gap_s)
      watch_keep <- bind_rows(watch_keep, add) %>% arrange(time_s)
      message("Backfill added: ", nrow(add), " -> final ", nrow(watch_keep))
    }
  }
  
} else {
  stop("SELECT_MODE must be 'uniform', 'best', or 'strong'.", call. = FALSE)
}

if (nrow(watch_keep) == 0) stop("0 clips kept after selection.", call. = FALSE)

# ----------------------------
# (6) OUTPUT watch points
# speed column carries event strength (peak magnitude)
# ----------------------------
watch_keep_out <- watch_keep %>%
  mutate(speed = strength) %>%
  select(frame, time_s, time_start, time_end, speed, teacher_ratio, EventID)

readr::write_csv(watch_keep_out, out_watch_csv)
message("Saved watch list: ", normalizePath(out_watch_csv, winslash = "/", mustWork = FALSE))

# ----------------------------
# (7) FFMPEG BAT (does NOT run ffmpeg)
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
message("Saved ffmpeg batch: ", normalizePath(bat_path, winslash = "/", mustWork = FALSE))

# ----------------------------
# (8) QC summary
# ----------------------------
qc <- dat2 %>%
  summarise(
    tag = TAG,
    total_frames = n(),
    teacher_present_ratio = mean(teacher_present, na.rm = TRUE),
    conf_ok_ratio = mean(conf_ok, na.rm = TRUE),
    shoulders_ok_ratio = mean(shoulders_ok, na.rm = TRUE),
    hand_speed_available = mean(!is.na(hand_speed)),
    speed_teacher_available = mean(!is.na(speed_teacher)),
    .groups = "drop"
  ) %>%
  mutate(
    select_mode = SELECT_MODE,
    teacher_first = TEACHER_FIRST,
    uniform_strict = UNIFORM_STRICT,
    backfill_mode = BACKFILL_MODE,
    teacher_thresholds = paste(TEACHER_THRESHOLDS, collapse = ","),
    strength_window_s = STRENGTH_WINDOW,
    time_col_used = col_time,
    n_events_total = nrow(cand2),
    n_events_candidate = nrow(cand3),
    n_watch_final = nrow(watch_keep_out),
    mean_teacher_ratio_in_clips = mean(watch_keep_out$teacher_ratio, na.rm = TRUE),
    mean_strength_in_clips = mean(watch_keep_out$speed, na.rm = TRUE),
    max_strength_in_clips  = suppressWarnings(max(watch_keep_out$speed, na.rm = TRUE)),
    min_time_s = min(watch_keep_out$time_s),
    max_time_s = max(watch_keep_out$time_s),
    # keep paths repo-relative in QC csv (cleaner)
    events_path = file.path("exports", TAG, paste0("gesture_events_", TAG, ".csv")),
    wrist_tracks_path = if (grepl(paste0("exports[/\\\\]wrist_tracks"), wrist_csv)) {
      file.path("exports", "wrist_tracks", paste0("wrist_tracks_", TAG, ".csv"))
    } else {
      file.path("exports", TAG, paste0("wrist_tracks_", TAG, ".csv"))
    },
    segment_video_path = file.path("segments", basename(video_path))
  )

readr::write_csv(qc, out_qc_csv)
print(qc)

cat("\nSaved:\n",
    "watch :", out_watch_csv, "\n",
    "bat   :", bat_path, "\n",
    "qc    :", out_qc_csv, "\n")
