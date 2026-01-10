# ============================================================
# 02_define_gestures.R  (STABLE / dplyr-1.1+ SAFE)
# Input : wrist_tracks_<TAG>.csv
# Output (exports/<TAG>/):
#   - gesture_events_<TAG>.csv
#   - summary_<TAG>.csv
#   - robust_<TAG>.csv
#   - debug_missing_reason_<TAG>_q<q>.csv
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
})

# ----------------------------
# (0) USER SETTINGS (PUBLIC TEMPLATE)
# ----------------------------
# IMPORTANT:
# - Do NOT commit personal local paths.
# - Configure via environment variables OR edit placeholders below.

get_env_or <- function(key, default) {
  v <- Sys.getenv(key, unset = "")
  if (nzchar(v)) v else default
}

is_placeholder_path <- function(x) {
  # detects "<PATH_TO_...>" style placeholders
  grepl("^<PATH_TO_", x)
}

# ---- identifiers (safe to commit) ----
TAG <- get_env_or("GESTURE_TAG", "mXX")

# ---- analysis parameters (not sensitive) ----
fps      <- 30L
q_main   <- 0.94
q_robust <- c(0.94, 0.95, 0.96)
min_len  <- 3L
win_sec  <- 1.0

# ---- input wrist_tracks directory ----
# Expected file: wrist_tracks_<TAG>.csv
in_dir <- get_env_or(
  "GESTURE_WRIST_DIR",
  "<PATH_TO_EXPORTS>/wrist_tracks"
)

# ---- output root ----
exports_root <- get_env_or(
  "GESTURE_EXPORTS_ROOT",
  "<PATH_TO_EXPORTS>"
)

# ---- SAFETY: prevent placeholder paths from silently passing ----
if (is_placeholder_path(in_dir) || is_placeholder_path(exports_root)) {
  stop(
    "Path placeholders detected.\n",
    "Please set env vars before running:\n",
    "  - GESTURE_WRIST_DIR (e.g., D:/gesture_project/exports/wrist_tracks)\n",
    "  - GESTURE_EXPORTS_ROOT (e.g., D:/gesture_project/exports)\n",
    "Current values:\n",
    "  GESTURE_WRIST_DIR    = ", in_dir, "\n",
    "  GESTURE_EXPORTS_ROOT = ", exports_root
  )
}

csv_path <- file.path(in_dir, paste0("wrist_tracks_", TAG, ".csv"))
if (!file.exists(csv_path)) {
  stop(
    "Input CSV not found:\n  ", csv_path, "\n\n",
    "Expected file name pattern: wrist_tracks_<TAG>.csv\n",
    "Check:\n",
    "  - TAG (GESTURE_TAG) = ", TAG, "\n",
    "  - in_dir (GESTURE_WRIST_DIR) points to the wrist_tracks folder."
  )
}

# ---- output directory ----
exports_dir <- file.path(exports_root, TAG)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

# ---- output files (FIX: define these!) ----
out_events_csv  <- file.path(exports_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(exports_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(exports_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(exports_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))

# ----------------------------
# (1) READ + BASIC STRUCTURE CHECK
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE)

# ensure frame / t_sec exist
if (!("frame" %in% names(dat))) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!("t_sec" %in% names(dat))) dat <- dat %>% mutate(t_sec = frame / fps)

# If teacher_present not available, assume TRUE (so debug won't crash)
if (!("teacher_present" %in% names(dat))) dat <- dat %>% mutate(teacher_present = TRUE)

# Ensure wrist columns exist
needed_cols <- c("lw_x","lw_y","rw_x","rw_y")
missing_cols <- setdiff(needed_cols, names(dat))
if (length(missing_cols) > 0) {
  stop("Missing required columns in input CSV: ", paste(missing_cols, collapse = ", "))
}

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
# (2) SPEED (LEFT, RIGHT, COMBINED)
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
    
    hand_speed = pmax(lw_speed, rw_speed, na.rm = TRUE)
  ) %>%
  mutate(hand_speed = ifelse(is.infinite(hand_speed), NA_real_, hand_speed))

video_minutes <- max(dat2$t_sec, na.rm = TRUE) / 60

# ----------------------------
# (3) EVENT EXTRACTION FUNCTION
# NOTE: use reframe() (dplyr 1.1+) because events are a multi-row output per group/run.
# ----------------------------
extract_events <- function(data, q, min_len = 3L, fps = 30L) {
  
  hs <- data$hand_speed
  if (all(is.na(hs))) return(tibble())
  
  thr <- suppressWarnings(as.numeric(quantile(hs, q, na.rm = TRUE, type = 7)))
  if (!is.finite(thr)) return(tibble())
  
  tmp <- data %>%
    arrange(frame) %>%
    mutate(
      above = !is.na(hand_speed) & (hand_speed > thr),
      run = cumsum(above != lag(above, default = FALSE))
    )
  
  events <- tmp %>%
    filter(above) %>%
    group_by(run) %>%
    reframe(
      start_frame = min(frame),
      end_frame   = max(frame),
      duration_frames = n(),
      start_sec = min(t_sec),
      end_sec   = max(t_sec),
      duration_sec = duration_frames / fps,
      mean_speed = mean(hand_speed, na.rm = TRUE),
      max_speed  = max(hand_speed,  na.rm = TRUE)
    ) %>%
    filter(duration_frames >= min_len) %>%
    mutate(q = q, threshold = thr)
  
  events
}

# ----------------------------
# (4) MAIN EVENTS (q = q_main)
# ----------------------------
gesture_events <- extract_events(dat2, q = q_main, min_len = min_len, fps = fps)
write_csv(gesture_events, out_events_csv)

summary_stats <- tibble(
  TAG = TAG,
  q = q_main,
  n_events = nrow(gesture_events),
  events_per_min = ifelse(video_minutes > 0, nrow(gesture_events) / video_minutes, NA_real_),
  mean_amplitude = if (nrow(gesture_events) > 0) mean(gesture_events$max_speed, na.rm = TRUE) else NA_real_,
  sd_amplitude   = if (nrow(gesture_events) > 1) sd(gesture_events$max_speed,   na.rm = TRUE) else NA_real_,
  mean_duration  = if (nrow(gesture_events) > 0) mean(gesture_events$duration_sec, na.rm = TRUE) else NA_real_,
  sd_duration    = if (nrow(gesture_events) > 1) sd(gesture_events$duration_sec,   na.rm = TRUE) else NA_real_
)
write_csv(summary_stats, out_summary_csv)

print(head(gesture_events, 10))
print(summary_stats)

# ----------------------------
# (4.5) DEBUG: Missing-Event Diagnostics (window-based)
# ----------------------------
thr_main <- suppressWarnings(as.numeric(quantile(dat2$hand_speed, q_main, na.rm = TRUE, type = 7)))

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

# mark ANY window overlapped by an event as has_event
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
cat("\n[DEBUG] Saved missing-reason table:\n  ", out_debug_csv, "\n", sep = "")

# ----------------------------
# (5) ROBUSTNESS
# ----------------------------
robust_tbl <- lapply(q_robust, function(qq) {
  ev <- extract_events(dat2, q = qq, min_len = min_len, fps = fps)
  tibble(
    TAG = TAG,
    q = qq,
    n_events = nrow(ev),
    events_per_min = ifelse(video_minutes > 0, nrow(ev) / video_minutes, NA_real_),
    mean_amplitude = if (nrow(ev) > 0) mean(ev$max_speed, na.rm = TRUE) else NA_real_,
    mean_duration  = if (nrow(ev) > 0) mean(ev$duration_sec, na.rm = TRUE) else NA_real_
  )
}) %>% bind_rows()

write_csv(robust_tbl, out_robust_csv)
print(robust_tbl)

cat("\nSaved:\n",
    "  events :", out_events_csv,  "\n",
    "  summary:", out_summary_csv, "\n",
    "  robust :", out_robust_csv,  "\n",
    "  debug  :", out_debug_csv,   "\n", sep = "")
