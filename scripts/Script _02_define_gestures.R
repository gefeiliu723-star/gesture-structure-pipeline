# ============================================================
# 02_define_gestures.R  (FINAL+DEBUG / dplyr-1.1+ SAFE)
#
# Input  : exports/wrist_tracks/wrist_tracks_<TAG>.csv  (default)
#          or pose_csv/wrist_tracks_<TAG>.csv
#          or exports/<TAG>/wrist_tracks_<TAG>.csv
#
# Output : exports/<TAG>/:
#   - gesture_events_<TAG>.csv
#   - summary_<TAG>.csv
#   - robust_<TAG>.csv
#   - debug_missing_reason_<TAG>_q<q>.csv
#
# Usage:
#   Rscript scripts/02_define_gestures.R
#   Rscript scripts/02_define_gestures.R m08
#   Rscript scripts/02_define_gestures.R m08 30 0.94 3
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# (0) USER SETTINGS (or CLI)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG <- if (length(args) >= 1) args[[1]] else "m01"
fps <- if (length(args) >= 2) as.integer(args[[2]]) else 30L

q_main   <- if (length(args) >= 3) as.numeric(args[[3]]) else 0.94
min_len  <- if (length(args) >= 4) as.integer(args[[4]]) else 3L
q_robust <- c(0.94, 0.95, 0.96)

# ----------------------------
# (0.1) PORTABLE PROJECT PATHS
# ----------------------------
proj_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

exports_root <- file.path(proj_root, "exports")
exports_dir  <- file.path(exports_root, TAG)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

# ----------------------------
# (0.2) LOCATE INPUT wrist_tracks_<TAG>.csv
# ----------------------------
find_wrist_tracks <- function(root, tag) {
  candidates <- c(
    file.path(root, "exports", "wrist_tracks", paste0("wrist_tracks_", tag, ".csv")),
    file.path(root, "pose_csv",               paste0("wrist_tracks_", tag, ".csv")),
    file.path(root, "exports", tag,           paste0("wrist_tracks_", tag, ".csv")),
    file.path(root, "exports",               paste0("wrist_tracks_", tag, ".csv"))
  )
  hit <- candidates[file.exists(candidates)]
  if (length(hit) == 0) {
    stop(
      "Cannot find wrist tracks CSV for TAG=", tag, "\n",
      "Tried:\n  - ", paste(candidates, collapse = "\n  - "), "\n",
      "Tip: run from repo root or set env var GESTURE_PROJECT_ROOT."
    )
  }
  hit[[1]]
}

csv_path <- find_wrist_tracks(proj_root, TAG)

# ----------------------------
# OUTPUT PATHS
# ----------------------------
out_events_csv  <- file.path(exports_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(exports_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(exports_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(exports_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))

# ----------------------------
# (1) READ + BASIC STRUCTURE
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE)

if (!"frame" %in% names(dat)) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!"t_sec" %in% names(dat)) dat <- dat %>% mutate(t_sec = frame / fps)
if (!"teacher_present" %in% names(dat)) dat <- dat %>% mutate(teacher_present = TRUE)

needed <- c("lw_x","lw_y","rw_x","rw_y")
miss <- setdiff(needed, names(dat))
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "))

# ---- (1.1) SANITY CHECK
sanity <- dat %>%
  summarise(
    TAG = TAG,
    total_frames = n(),
    lw_detect_ratio = mean(!is.na(lw_x)),
    rw_detect_ratio = mean(!is.na(rw_x)),
    any_wrist_detect_ratio = mean(!is.na(lw_x) | !is.na(rw_x)),
    teacher_present_ratio = mean(teacher_present),
    .groups = "drop"
  )
print(sanity)

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

print(head(gesture_events, 10))
print(summary_tbl)

# ----------------------------
# (4.5) DEBUG: Missing-Event Diagnostics (window-based)
# ----------------------------
thr_main <- suppressWarnings(as.numeric(quantile(dat2$hand_speed, q_main, na.rm = TRUE, type = 7)))
win_sec <- 1.0

debug_windows <- dat2 %>%
  mutate(
    above_thr = is.finite(thr_main) & !is.na(hand_speed) & (hand_speed > thr_main),
    hand_na = is.na(hand_speed),
    teacher_off = !teacher_present,
    win_id = floor(t_sec / win_sec)
  ) %>%
  group_by(win_id) %>%
  summarise(
    TAG = TAG,
    q = q_main,
    threshold = thr_main,
    t_start = min(t_sec, na.rm = TRUE),
    t_end   = max(t_sec, na.rm = TRUE),
    n_frames = n(),
    na_rate = mean(hand_na),
    teacher_off_rate = mean(teacher_off),
    above_thr_rate = mean(above_thr),
    max_run_above = {
      r <- rle(above_thr)
      if (any(r$values)) max(r$lengths[r$values]) else 0L
    },
    .groups = "drop"
  ) %>%
  mutate(
    missing_reason = case_when(
      na_rate > 0.30 ~ "NA_dropout",
      teacher_off_rate > 0.30 ~ "teacher_gating",
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
cat("\n[DEBUG] Saved missing-reason table:\n", out_debug_csv, "\n")

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
print(robust_tbl)

cat("\nSaved:\n",
    "read   :", csv_path,        "\n",
    "events :", out_events_csv,  "\n",
    "summary:", out_summary_csv, "\n",
    "robust :", out_robust_csv,  "\n",
    "debug  :", out_debug_csv,   "\n")

cat("DONE:", TAG, "\n")
