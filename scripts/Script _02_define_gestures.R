# ============================================================
# 02_define_gestures.R  (GitHub-ready / dplyr-1.1+ SAFE)
# Input : exports/wrist_tracks/wrist_tracks_<TAG>.csv
# Output: exports/<TAG>/
#   - gesture_events_<TAG>.csv
#   - summary_<TAG>.csv
#   - robust_<TAG>.csv
#   - debug_missing_reason_<TAG>_q<q>.csv   (reserved for downstream use)

#
# Usage examples:
#   Rscript scripts/02_define_gestures.R
#   Rscript scripts/02_define_gestures.R m08
#   Rscript scripts/02_define_gestures.R m08 30 0.94 3
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/gesture_project
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
})

# ----------------------------
# (0) PROJECT ROOT (portable)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG <- if (length(args) >= 1) args[[1]] else "m08"      # <-- change or pass via CLI
fps <- if (length(args) >= 2) as.integer(args[[2]]) else 30L

q_main   <- if (length(args) >= 3) as.numeric(args[[3]]) else 0.94
min_len  <- if (length(args) >= 4) as.integer(args[[4]]) else 3L
q_robust <- c(0.94, 0.95, 0.96)

# root priority:
#  1) env var GESTURE_PROJECT_ROOT
#  2) current working directory (repo root if you run from there)
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

exports_root <- file.path(project_root, "exports")
in_dir       <- file.path(exports_root, "wrist_tracks")
csv_path     <- file.path(in_dir, paste0("wrist_tracks_", TAG, ".csv"))

exports_dir <- file.path(exports_root, TAG)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

out_events_csv  <- file.path(exports_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(exports_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(exports_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(exports_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))

# ----------------------------
# (1) READ + BASIC STRUCTURE
# ----------------------------
if (!file.exists(csv_path)) {
  stop(
    "Input not found: ", csv_path, "\n",
    "Expected file at: exports/wrist_tracks/wrist_tracks_", TAG, ".csv\n",
    "Tip: run from repo root OR set env var GESTURE_PROJECT_ROOT."
  )
}

dat <- read_csv(csv_path, show_col_types = FALSE)

if (!"frame" %in% names(dat)) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!"t_sec" %in% names(dat)) dat <- dat %>% mutate(t_sec = frame / fps)
if (!"teacher_present" %in% names(dat)) dat <- dat %>% mutate(teacher_present = TRUE)

needed <- c("lw_x","lw_y","rw_x","rw_y")
miss <- setdiff(needed, names(dat))
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "))

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
    hand_speed = ifelse(is.infinite(hand_speed), NA_real_, hand_speed)
  )

video_minutes <- max(dat2$t_sec, na.rm = TRUE) / 60

# ----------------------------
# (3) EVENT EXTRACTION (SAFE)
# ----------------------------
extract_events <- function(data, q, min_len, fps) {
  
  hs <- data$hand_speed
  if (all(is.na(hs))) return(tibble())
  
  thr <- as.numeric(quantile(hs, q, na.rm = TRUE, type = 7))
  if (!is.finite(thr)) return(tibble())
  
  tmp <- data %>%
    arrange(frame) %>%
    mutate(
      above = !is.na(hand_speed) & hand_speed > thr,
      run   = cumsum(above != lag(above, default = FALSE))
    )
  
  events <- tmp %>%
    filter(above) %>%
    group_by(run) %>%
    reframe(
      start_frame     = min(frame),
      end_frame       = max(frame),
      duration_frames = n(),
      start_sec       = min(t_sec),
      end_sec         = max(t_sec),
      duration_sec    = duration_frames / fps,
      mean_speed      = mean(hand_speed, na.rm = TRUE),
      max_speed       = max(hand_speed, na.rm = TRUE),
      peak_frame      = frame[which.max(replace(hand_speed, is.na(hand_speed), -Inf))],
      peak_sec        = t_sec[which.max(replace(hand_speed, is.na(hand_speed), -Inf))]
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
      EventID, run,
      start_frame, end_frame, duration_frames,
      start_sec, end_sec, duration_sec,
      mean_speed, max_speed,
      peak_frame, peak_sec, peak_sec_proxy,
      q, threshold
    )
  
  events
}

# ----------------------------
# (4) MAIN EVENTS
# ----------------------------
gesture_events <- extract_events(dat2, q_main, min_len, fps)
write_csv(gesture_events, out_events_csv)

summary_tbl <- tibble(
  TAG = TAG,
  q = q_main,
  n_events = nrow(gesture_events),
  events_per_min = ifelse(video_minutes > 0, nrow(gesture_events) / video_minutes, NA_real_),
  mean_amplitude = mean(gesture_events$max_speed, na.rm = TRUE),
  sd_amplitude   = sd(gesture_events$max_speed, na.rm = TRUE),
  mean_duration  = mean(gesture_events$duration_sec, na.rm = TRUE),
  sd_duration    = sd(gesture_events$duration_sec, na.rm = TRUE)
)

write_csv(summary_tbl, out_summary_csv)

# ----------------------------
# (5) ROBUSTNESS
# ----------------------------
robust_tbl <- bind_rows(lapply(q_robust, function(qq) {
  ev <- extract_events(dat2, qq, min_len, fps)
  tibble(
    TAG = TAG,
    q = qq,
    n_events = nrow(ev),
    events_per_min = ifelse(video_minutes > 0, nrow(ev) / video_minutes, NA_real_),
    mean_amplitude = mean(ev$max_speed, na.rm = TRUE),
    mean_duration  = mean(ev$duration_sec, na.rm = TRUE)
  )
}))

write_csv(robust_tbl, out_robust_csv)

cat(
  "DONE:", TAG, "\n",
  "Input :", normalizePath(csv_path, winslash = "/", mustWork = FALSE), "\n",
  "Output:", normalizePath(exports_dir, winslash = "/", mustWork = FALSE), "\n"
)
