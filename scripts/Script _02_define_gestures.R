# ============================================================
# 02_define_gestures.R  (GitHub-clean / dplyr-1.1+ SAFE)
#
# Input  (fixed):
#   exports/wrist_tracks/wrist_tracks_<TAG>.csv
#
# Output (fixed) -> exports/<TAG>/:
#   - gesture_events_<TAG>.csv        (template-ready: peak-frame snapshot + event stats)
#   - summary_<TAG>.csv
#   - robust_<TAG>.csv
#   - debug_missing_reason_<TAG>_q<q>.csv   (path reserved; not written here)
#
# CLI:
#   Rscript scripts/02_define_gestures.R m02 30 0.94 3 0.94,0.95,0.96
#
# Env (optional):
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(dplyr)
  library(tibble)
  library(stringr)
})

# ----------------------------
# (0) SETTINGS (CLI or defaults)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG <- if (length(args) >= 1 && nzchar(args[[1]])) args[[1]] else "m02"
fps <- if (length(args) >= 2 && nzchar(args[[2]])) as.integer(args[[2]]) else 30L

q_main  <- if (length(args) >= 3 && nzchar(args[[3]])) as.numeric(args[[3]]) else 0.94
min_len <- if (length(args) >= 4 && nzchar(args[[4]])) as.integer(args[[4]]) else 3L

q_robust <- c(0.94, 0.95, 0.96)
if (length(args) >= 5 && nzchar(args[[5]])) {
  q_robust <- suppressWarnings(as.numeric(strsplit(args[[5]], ",", fixed = TRUE)[[1]]))
  q_robust <- q_robust[is.finite(q_robust)]
  if (length(q_robust) == 0) q_robust <- c(0.94, 0.95, 0.96)
}

stopifnot(is.finite(fps), fps > 0)
stopifnot(is.finite(q_main), q_main > 0, q_main < 1)
stopifnot(is.finite(min_len), min_len >= 1)

# project root (GitHub clean)
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = "")
if (!nzchar(project_root)) project_root <- getwd()

exports_root <- file.path(project_root, "exports")
in_dir  <- file.path(exports_root, "wrist_tracks")
csv_path <- file.path(in_dir, paste0("wrist_tracks_", TAG, ".csv"))

exports_dir <- file.path(exports_root, TAG)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

out_events_csv  <- file.path(exports_dir, paste0("gesture_events_", TAG, ".csv"))
out_summary_csv <- file.path(exports_dir, paste0("summary_", TAG, ".csv"))
out_robust_csv  <- file.path(exports_dir, paste0("robust_", TAG, ".csv"))
out_debug_csv   <- file.path(exports_dir, paste0("debug_missing_reason_", TAG, "_q", q_main, ".csv"))

if (!file.exists(csv_path)) {
  stop("Missing input: ", csv_path,
       "\nExpected: exports/wrist_tracks/wrist_tracks_<TAG>.csv",
       "\nTAG=", TAG,
       "\nproject_root=", project_root)
}

message("TAG = ", TAG)
message("Input : ", csv_path)
message("Output: ", exports_dir)

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

# ADDED: n_frames (used by template fields)
n_frames_val <- if ("n_frames" %in% names(dat)) suppressWarnings(as.numeric(dat$n_frames[1])) else NA_real_
if (!is.finite(n_frames_val)) n_frames_val <- max(dat$frame, na.rm = TRUE) + 1L

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
# (3) EVENT EXTRACTION (SAFE)  -- keep logic
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
    # remove duplicated events (same window produced more than once)
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
  
  events
}

# ----------------------------
# (4) MAIN EVENTS
# ----------------------------
gesture_events <- extract_events(dat2, q_main, min_len, fps)

# build peak-frame snapshot columns for template
template_cols <- c(
  "frame","t_sec",
  "teacher_id","conf_det","x1","y1","x2","y2","frame_w","frame_h",
  "ls_x","ls_y","ls_z","ls_vis","rs_x","rs_y","rs_z","rs_vis",
  "le_x","le_y","le_z","le_vis","re_x","re_y","re_z","re_vis",
  "lw_x","lw_y","lw_z","lw_vis","rw_x","rw_y","rw_z","rw_vis"
)
# ensure missing cols exist (avoid select errors)
for (nm in setdiff(template_cols, names(dat2))) dat2[[nm]] <- NA

template_extra <- dat2 %>% select(any_of(template_cols))

gesture_events_out <- gesture_events %>%
  left_join(template_extra, by = c("peak_frame" = "frame")) %>%
  mutate(
    fps = as.integer(fps),
    n_frames = n_frames_val,
    frame = peak_frame,  # template wants peak frame
    t_sec = peak_sec     # template wants peak t_sec
  )

# strict header order requested (template head first)
template_head <- c(
  "EventID","frame","t_sec","teacher_id","conf_det","x1","y1","x2","y2","frame_w","frame_h","fps","n_frames",
  "ls_x","ls_y","ls_z","ls_vis","rs_x","rs_y","rs_z","rs_vis",
  "le_x","le_y","le_z","le_vis","re_x","re_y","re_z","re_vis",
  "lw_x","lw_y","lw_z","lw_vis","rw_x","rw_y","rw_z","rw_vis"
)
# keep your original event stats after template head
keep_cols <- c(template_head, setdiff(names(gesture_events_out), template_head))
gesture_events_out <- gesture_events_out %>% select(any_of(keep_cols))

write_csv(gesture_events_out, out_events_csv)
message("Saved: ", out_events_csv)

# ----------------------------
# SUMMARY
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
# ROBUSTNESS
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

cat("DONE:", TAG, "\n")
