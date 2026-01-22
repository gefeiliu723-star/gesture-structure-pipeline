# ============================================================
# 02_define_gestures.R  (GitHub-ready / dplyr-1.1+ SAFE)
#
# Input:
#   exports/wrist_tracks/wrist_tracks_<TAG>.csv
#
# Output (exports/<TAG>/):
#   - gesture_events_<TAG>.csv
#   - summary_<TAG>.csv
#   - robust_<TAG>.csv
#   - debug_missing_reason_<TAG>_q<q>.csv   (WINDOW-LEVEL TABLE)
#
# Usage:
#   Rscript scripts/02_define_gestures.R
#   Rscript scripts/02_define_gestures.R l07
#   Rscript scripts/02_define_gestures.R l07 30 0.94 3
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
#   (default: current working directory)

# NOTE 02-A
#   Gesture events are defined as sustained local maxima in hand speed, not communicative acts.
#   The event definition is intentionally agnostic to gesture type or semantic content.

# NOTE 02-B
#   Quantile thresholds are computed within-video to preserve relative salience under heterogeneous recording conditions.

# NOTE 02-C（防 reviewer）
#   Windows without detected events are explicitly categorized (e.g., teacher absence, insufficient run length), rather than treated as missing data.
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tibble)
})

# ----------------------------
# (0) SETTINGS / CLI
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG     <- if (length(args) >= 1) args[[1]] else "l07"
fps     <- if (length(args) >= 2) suppressWarnings(as.numeric(args[[2]])) else 30
q_main  <- if (length(args) >= 3) suppressWarnings(as.numeric(args[[3]])) else 0.94
min_len <- if (length(args) >= 4) suppressWarnings(as.integer(args[[4]])) else 3L

q_robust <- c(0.94, 0.95, 0.96)

stopifnot(is.character(TAG), nzchar(TAG))
stopifnot(is.finite(fps), fps > 0)
stopifnot(is.finite(q_main), q_main > 0, q_main < 1)
stopifnot(is.finite(min_len), min_len >= 1)

# ----------------------------
# (0.1) PATHS (GitHub-ready)
# ----------------------------
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
wrist_dir    <- file.path(exports_root, "wrist_tracks")
tag_dir      <- file.path(exports_root, TAG)

dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)

csv_path <- file.path(wrist_dir, paste0("wrist_tracks_", TAG, ".csv"))
if (!file.exists(csv_path)) {
  stop("Input not found: ", csv_path,
       "\nExpected: exports/wrist_tracks/wrist_tracks_<TAG>.csv",
       call. = FALSE)
}

out_events_csv  <- file.path(tag_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(tag_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(tag_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(tag_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))

# ----------------------------
# (1) READ + BASIC STRUCTURE
# ----------------------------
dat <- read_csv(csv_path, show_col_types = FALSE, progress = FALSE)

if (!"frame" %in% names(dat)) dat <- dat %>% mutate(frame = row_number() - 1L)
if (!"t_sec" %in% names(dat)) dat <- dat %>% mutate(t_sec = frame / fps)
if (!"teacher_present" %in% names(dat)) dat <- dat %>% mutate(teacher_present = TRUE)

needed <- c("lw_x","lw_y","rw_x","rw_y")
miss <- setdiff(needed, names(dat))
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "), call. = FALSE)

# n_frames (for template compatibility downstream)
n_frames_val <- if ("n_frames" %in% names(dat)) suppressWarnings(as.numeric(dat$n_frames[1])) else NA_real_
if (!is.finite(n_frames_val)) n_frames_val <- max(dat$frame, na.rm = TRUE) + 1L

# ----------------------------
# (1.5) SANITY COERCION
# ----------------------------
as_num_safe <- function(x) suppressWarnings(as.numeric(x))
dat <- dat %>%
  mutate(across(any_of(c("frame","t_sec","lw_x","lw_y","rw_x","rw_y")), as_num_safe))

if (!all(is.finite(dat$frame))) stop("Non-finite values found in `frame` after coercion.", call. = FALSE)
if (!all(is.finite(dat$t_sec))) stop("Non-finite values found in `t_sec` after coercion.", call. = FALSE)

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
# (2.5) DEBUG MISSING (WINDOW-LEVEL TABLE)
# columns:
# win_id TAG q threshold t_start t_end n_frames
# na_rate teacher_off_rate above_thr_rate max_run_above
# missing_reason has_event
# ----------------------------
win_sec  <- 10
step_sec <- 10
win_n    <- max(1L, as.integer(round(win_sec * fps)))
step_n   <- max(1L, as.integer(round(step_sec * fps)))

teacher_present_vec <- dat2$teacher_present
teacher_present_vec <- ifelse(is.na(teacher_present_vec), FALSE, as.logical(teacher_present_vec))

hs <- dat2$hand_speed
thr_main <- if (all(is.na(hs))) NA_real_ else as.numeric(quantile(hs, q_main, na.rm = TRUE, type = 7))

max_run_len <- function(x) {
  if (length(x) == 0) return(0L)
  x <- as.logical(x)
  if (all(!x, na.rm = TRUE)) return(0L)
  r <- rle(x)
  if (!any(r$values)) return(0L)
  max(r$lengths[r$values], na.rm = TRUE)
}

N <- nrow(dat2)
starts <- seq(1L, N, by = step_n)
ends   <- pmin(starts + win_n - 1L, N)

debug_missing_tbl <- bind_rows(lapply(seq_along(starts), function(i) {
  s <- starts[[i]]
  e <- ends[[i]]
  w <- dat2[s:e, , drop = FALSE]
  
  n_frames <- nrow(w)
  t_start  <- suppressWarnings(min(w$t_sec, na.rm = TRUE))
  t_end    <- suppressWarnings(max(w$t_sec, na.rm = TRUE))
  
  hs_w <- w$hand_speed
  na_rate <- mean(is.na(hs_w))
  
  tp_w <- teacher_present_vec[s:e]
  teacher_off_rate <- mean(!tp_w)
  
  above <- if (!is.finite(thr_main)) rep(FALSE, n_frames) else (!is.na(hs_w) & hs_w > thr_main)
  above_thr_rate <- mean(above)
  max_run_above  <- max_run_len(above)
  
  tibble(
    win_id = sprintf("W%04d", i),
    TAG = TAG,
    q = q_main,
    threshold = thr_main,
    t_start = t_start,
    t_end = t_end,
    n_frames = n_frames,
    na_rate = na_rate,
    teacher_off_rate = teacher_off_rate,
    above_thr_rate = above_thr_rate,
    max_run_above = max_run_above
  )
}))

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
  
  tmp %>%
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
    distinct(start_frame, end_frame, .keep_all = TRUE) %>%
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
}

# ----------------------------
# (4) MAIN EVENTS + TEMPLATE ENRICH
# ----------------------------
gesture_events <- extract_events(dat2, q_main, min_len, fps)

template_extra <- dat2 %>%
  select(any_of(c(
    "frame","t_sec",
    "teacher_id","conf_det","x1","y1","x2","y2","frame_w","frame_h",
    "ls_x","ls_y","ls_z","ls_vis","rs_x","rs_y","rs_z","rs_vis",
    "le_x","le_y","le_z","le_vis","re_x","re_y","re_z","re_vis",
    "lw_x","lw_y","lw_z","lw_vis","rw_x","rw_y","rw_z","rw_vis"
  )))

gesture_events_out <- gesture_events %>%
  left_join(template_extra, by = c("peak_frame" = "frame")) %>%
  mutate(
    fps = fps,
    n_frames = n_frames_val,
    frame = peak_frame,
    t_sec = peak_sec
  )

template_head <- c(
  "EventID","frame","t_sec","teacher_id","conf_det","x1","y1","x2","y2","frame_w","frame_h","fps","n_frames",
  "ls_x","ls_y","ls_z","ls_vis","rs_x","rs_y","rs_z","rs_vis",
  "le_x","le_y","le_z","le_vis","re_x","re_y","re_z","re_vis",
  "lw_x","lw_y","lw_z","lw_vis","rw_x","rw_y","rw_z","rw_vis"
)
keep_cols <- c(template_head, setdiff(names(gesture_events_out), template_head))
gesture_events_out <- gesture_events_out %>% select(any_of(keep_cols))

write_csv(gesture_events_out, out_events_csv)

# ----------------------------
# (4.1) FINALIZE DEBUG MISSING (has_event + missing_reason) + WRITE
# ----------------------------
if (nrow(debug_missing_tbl) > 0) {
  if (nrow(gesture_events) == 0) {
    debug_missing_tbl$has_event <- FALSE
  } else {
    ev_s <- as.numeric(gesture_events$start_sec)
    ev_e <- as.numeric(gesture_events$end_sec)
    
    debug_missing_tbl$has_event <- vapply(seq_len(nrow(debug_missing_tbl)), function(i) {
      ws <- as.numeric(debug_missing_tbl$t_start[i])
      we <- as.numeric(debug_missing_tbl$t_end[i])
      if (!is.finite(ws) || !is.finite(we)) return(FALSE)
      any(is.finite(ev_s) & is.finite(ev_e) & (ev_s <= we) & (ev_e >= ws))
    }, logical(1))
  }
  
  debug_missing_tbl$missing_reason <- vapply(seq_len(nrow(debug_missing_tbl)), function(i) {
    if (isTRUE(debug_missing_tbl$has_event[i])) return("has_event")
    if (!is.finite(debug_missing_tbl$threshold[i])) return("threshold_non_finite_or_all_na")
    if (isTRUE(debug_missing_tbl$na_rate[i] >= 0.95)) return("hand_speed_mostly_na")
    if (isTRUE(debug_missing_tbl$teacher_off_rate[i] >= 0.95)) return("teacher_off_mostly")
    if (isTRUE(debug_missing_tbl$above_thr_rate[i] == 0)) return("no_frames_above_threshold")
    if (isTRUE(debug_missing_tbl$max_run_above[i] < min_len)) return(paste0("runs_above_exist_but_shorter_than_min_len_", min_len))
    "no_event_unknown"
  }, character(1))
}

debug_missing_tbl <- debug_missing_tbl %>%
  select(
    win_id, TAG, q, threshold, t_start, t_end, n_frames,
    na_rate, teacher_off_rate, above_thr_rate, max_run_above,
    missing_reason, has_event
  )

write_csv(debug_missing_tbl, out_debug_csv)

# ----------------------------
# (4.2) SUMMARY
# ----------------------------
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

cat("DONE:", TAG, "\n",
    "Input :", normalizePath(csv_path, winslash = "/", mustWork = FALSE), "\n",
    "Events:", normalizePath(out_events_csv, winslash = "/", mustWork = FALSE), "\n",
    "Debug :", normalizePath(out_debug_csv, winslash = "/", mustWork = FALSE), "\n")
