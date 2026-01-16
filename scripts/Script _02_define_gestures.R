# ============================================================
# 02_define_gestures.R  (GitHub-ready / NO LOCAL PATH / NO TAG EDIT)
#
# RUN:
#   Rscript scripts/02_define_gestures.R path/to/wrist_tracks_m02.csv
#
# Optional:
#   Rscript scripts/02_define_gestures.R path/to/wrist_tracks_m02.csv 30 0.94 3
#
# Output (fixed relative to repo):
#   exports/<TAG>/
#     - gesture_events_<TAG>.csv
#     - summary_<TAG>.csv
#     - robust_<TAG>.csv
#     - debug_missing_reason_<TAG>_q<q>.csv
#     - gesture_qc_<TAG>.csv
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tibble)
  library(stringr)
})

# ----------------------------
# (0) CLI (NO TAG EDIT)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 1) {
  stop(
    "Usage:\n",
    "  Rscript scripts/02_define_gestures.R <wrist_tracks_csv> [fps] [q_main] [min_len]\n\n",
    "Example:\n",
    "  Rscript scripts/02_define_gestures.R exports/wrist_tracks/wrist_tracks_m02.csv 30 0.94 3\n",
    call. = FALSE
  )
}

csv_path <- args[[1]]
fps      <- if (length(args) >= 2) as.numeric(args[[2]]) else 30
q_main   <- if (length(args) >= 3) as.numeric(args[[3]]) else 0.94
min_len  <- if (length(args) >= 4) as.integer(args[[4]]) else 3L

q_robust <- c(0.94, 0.95, 0.96)

stopifnot(is.finite(fps), fps > 0)
stopifnot(is.finite(q_main), q_main > 0, q_main < 1)
stopifnot(is.finite(min_len), min_len >= 1)

if (!file.exists(csv_path)) {
  stop("Cannot find input CSV:\n  ", csv_path, call. = FALSE)
}

# ----------------------------
# (0.1) PROJECT ROOT (repo root)
# ----------------------------
# Prefer env var; fallback getwd().
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = "")
if (!nzchar(project_root)) project_root <- getwd()
exports_root <- file.path(project_root, "exports")

# ----------------------------
# (0.2) AUTO TAG (from filename)
# ----------------------------
# Try to parse something like wrist_tracks_m02.csv -> m02
base <- basename(csv_path)
tag_guess <- str_match(tolower(base), "wrist[_-]?tracks[_-]?([a-z0-9]+)\\.(csv|txt)$")[,2]
TAG <- ifelse(!is.na(tag_guess) && nzchar(tag_guess), tag_guess, "tag")

exports_dir <- file.path(exports_root, TAG)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

# ----------------------------
# OUTPUT PATHS
# ----------------------------
out_events_csv  <- file.path(exports_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(exports_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(exports_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(exports_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))
out_gesture_qc  <- file.path(exports_dir, paste0("gesture_qc_", TAG, ".csv"))

cat("\n[INFO] project_root:", normalizePath(project_root, winslash = "/", mustWork = FALSE), "\n")
cat("[INFO] input_csv   :", normalizePath(csv_path, winslash = "/", mustWork = FALSE), "\n")
cat("[INFO] TAG         :", TAG, "\n")
cat("[INFO] out_dir     :", normalizePath(exports_dir, winslash = "/", mustWork = FALSE), "\n\n")

# ----------------------------
# (1) READ + BASIC STRUCTURE
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE, progress = FALSE)

if (!"frame" %in% names(dat)) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!"t_sec" %in% names(dat))  dat <- dat %>% mutate(t_sec = frame / fps)

needed <- c("lw_x","lw_y","rw_x","rw_y")
miss <- setdiff(needed, names(dat))
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "), call. = FALSE)

# ----------------------------
# (1.1) GESTURE SIGNAL QC
# ----------------------------
gesture_qc_tbl <- tibble(
  TAG = TAG,
  total_frames = nrow(dat),
  lw_detect_ratio = mean(!is.na(dat$lw_x)),
  rw_detect_ratio = mean(!is.na(dat$rw_x)),
  any_wrist_detect_ratio = mean(!is.na(dat$lw_x) | !is.na(dat$rw_x)),
  has_teacher_present_col = ("teacher_present" %in% names(dat)),
  has_conf_det_col        = ("conf_det" %in% names(dat)),
  has_ls_vis_col          = ("ls_vis" %in% names(dat)),
  has_rs_vis_col          = ("rs_vis" %in% names(dat))
)

write_csv(gesture_qc_tbl, out_gesture_qc)
cat("[QC] Saved:", out_gesture_qc, "\n")
print(gesture_qc_tbl)

# ----------------------------
# (2) SPEED SIGNAL
# ----------------------------
dat2 <- dat %>%
  arrange(frame) %>%
  mutate(
    lw_dx = lw_x - lag(lw_x),
    lw_dy = lw_y - lag(lw_y),
    rw_dx = rw_x - lag(rw_x),
    rw_dy = rw_y - lag(rw_y),
    
    lw_speed = sqrt(lw_dx^2 + lw_dy^2) * fps,
    rw_speed = sqrt(rw_dx^2 + rw_dy^2) * fps,
    
    hand_speed = pmax(lw_speed, rw_speed, na.rm = TRUE),
    hand_speed = ifelse(is.finite(hand_speed), hand_speed, NA_real_)
  )

video_minutes <- max(dat2$t_sec, na.rm = TRUE) / 60

# ----------------------------
# (3) EVENT EXTRACTION
# ----------------------------
extract_events <- function(data, q, min_len, fps) {
  hs <- data$hand_speed
  if (length(hs) == 0 || all(is.na(hs))) return(tibble())
  
  thr <- suppressWarnings(as.numeric(quantile(hs, q, na.rm = TRUE, type = 7)))
  if (!is.finite(thr)) return(tibble())
  
  tmp <- data %>%
    arrange(frame) %>%
    mutate(
      above = !is.na(hand_speed) & hand_speed > thr,
      grp   = cumsum(above != lag(above, default = FALSE))
    )
  
  tmp %>%
    filter(above) %>%
    group_by(grp) %>%
    reframe(
      start_frame     = min(frame),
      end_frame       = max(frame),
      duration_frames = dplyr::n(),
      start_sec       = min(t_sec),
      end_sec         = max(t_sec),
      duration_sec    = duration_frames / fps,
      
      mean_speed      = mean(hand_speed, na.rm = TRUE),
      max_speed       = max(hand_speed, na.rm = TRUE),
      
      peak_idx        = which.max(replace(hand_speed, is.na(hand_speed), -Inf))[1],
      peak_frame      = frame[peak_idx][1],
      peak_sec        = t_sec[peak_idx][1]
    ) %>%
    filter(duration_frames >= min_len) %>%
    arrange(start_sec) %>%
    mutate(
      EventID = sprintf("E%04d", row_number()),
      peak_sec_proxy = peak_sec,
      q = q,
      threshold = thr
    ) %>%
    select(
      EventID, grp,
      start_frame, end_frame, duration_frames,
      start_sec, end_sec, duration_sec,
      mean_speed, max_speed,
      peak_frame, peak_sec, peak_sec_proxy,
      q, threshold
    )
}

# ----------------------------
# (4) MAIN EVENTS + SUMMARY
# ----------------------------
gesture_events <- extract_events(dat2, q_main, min_len, fps)
write_csv(gesture_events, out_events_csv)

summary_tbl <- tibble(
  TAG = TAG,
  q = q_main,
  n_events = nrow(gesture_events),
  events_per_min = ifelse(video_minutes > 0, nrow(gesture_events) / video_minutes, NA_real_),
  mean_amplitude = if (nrow(gesture_events) > 0) mean(gesture_events$max_speed, na.rm = TRUE) else NA_real_,
  sd_amplitude   = if (nrow(gesture_events) > 1)  sd(gesture_events$max_speed, na.rm = TRUE) else NA_real_,
  mean_duration  = if (nrow(gesture_events) > 0) mean(gesture_events$duration_sec, na.rm = TRUE) else NA_real_,
  sd_duration    = if (nrow(gesture_events) > 1)  sd(gesture_events$duration_sec, na.rm = TRUE) else NA_real_
)

write_csv(summary_tbl, out_summary_csv)

cat("\n[EVENTS] Head:\n")
print(head(gesture_events, 10))
cat("\n[SUMMARY]\n")
print(summary_tbl)

# ----------------------------
# (4.5) DEBUG: Missing-Event Diagnostics (1s windows)
# ----------------------------
win_sec <- 1.0
thr_main <- suppressWarnings(as.numeric(quantile(dat2$hand_speed, q_main, na.rm = TRUE, type = 7)))

debug_windows <- dat2 %>%
  mutate(
    above_thr = is.finite(thr_main) & !is.na(hand_speed) & (hand_speed > thr_main),
    hand_na   = is.na(hand_speed),
    win_id    = floor(t_sec / win_sec)
  ) %>%
  group_by(win_id) %>%
  reframe(
    TAG = TAG,
    q = q_main,
    threshold = thr_main,
    t_start = min(t_sec, na.rm = TRUE),
    t_end   = max(t_sec, na.rm = TRUE),
    n_frames = dplyr::n(),
    na_rate = mean(hand_na),
    above_thr_rate = mean(above_thr),
    max_run_above = {
      r <- rle(above_thr)
      if (any(r$values)) max(r$lengths[r$values]) else 0L
    }
  ) %>%
  mutate(
    missing_reason = case_when(
      na_rate > 0.30 ~ "NA_dropout",
      above_thr_rate < 0.10 ~ "below_threshold",
      max_run_above < min_len ~ "too_short",
      TRUE ~ "none"
    )
  )

event_win_ids <- integer()
if (nrow(gesture_events) > 0) {
  event_win_ids <- unique(unlist(mapply(
    FUN = function(s, e) seq.int(floor(s / win_sec), floor(e / win_sec)),
    s = gesture_events$start_sec,
    e = gesture_events$end_sec,
    SIMPLIFY = FALSE
  )))
}

debug_missing <- debug_windows %>%
  mutate(has_event = win_id %in% event_win_ids) %>%
  filter(!has_event) %>%
  arrange(t_start)

write_csv(debug_missing, out_debug_csv)
cat("\n[DEBUG] Saved:", out_debug_csv, "\n")

# ----------------------------
# (5) ROBUSTNESS GRID
# ----------------------------
robust_tbl <- bind_rows(lapply(q_robust, function(qq) {
  ev <- extract_events(dat2, qq, min_len, fps)
  tibble(
    TAG = TAG,
    q = qq,
    n_events = nrow(ev),
    events_per_min = ifelse(video_minutes > 0, nrow(ev) / video_minutes, NA_real_),
    mean_amplitude = if (nrow(ev) > 0) mean(ev$max_speed, na.rm = TRUE) else NA_real_,
    mean_duration  = if (nrow(ev) > 0) mean(ev$duration_sec, na.rm = TRUE) else NA_real_
  )
}))

write_csv(robust_tbl, out_robust_csv)
cat("[ROBUST] Saved:", out_robust_csv, "\n")
print(robust_tbl)

cat("\nSaved under:\n  ", normalizePath(exports_dir, winslash = "/", mustWork = FALSE), "\n",
    "events      : ", basename(out_events_csv),  "\n",
    "summary     : ", basename(out_summary_csv), "\n",
    "robust      : ", basename(out_robust_csv),  "\n",
    "debug       : ", basename(out_debug_csv),   "\n",
    "gesture_qc  : ", basename(out_gesture_qc),  "\n",
    sep = "")
cat("DONE:", TAG, "\n")
