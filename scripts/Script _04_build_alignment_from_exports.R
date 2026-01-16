# ============================================================
# 04_build_alignment_from_template.R  (GitHub-ready / NO PATH / NO TAG EDIT)
#   - DO NOT TOUCH Script 3
#   - Methods countsflow computed from wrist_tracks + watch_points
#
# RUN (recommended):
#   Rscript scripts/04_build_alignment_from_template.R \
#     exports/wrist_tracks/wrist_tracks_m10.csv \
#     exports/m10/watch_points_m10_teacher.csv \
#     exports/m10/gesture_events_m10.csv
#
# Optional args:
#   [fps] [dt_seconds]
# Example:
#   Rscript scripts/04_build_alignment_from_template.R \
#     exports/wrist_tracks/wrist_tracks_m10.csv \
#     exports/m10/watch_points_m10_teacher.csv \
#     exports/m10/gesture_events_m10.csv 30 1.5
#
# Template auto-find:
#   exports/_alignment_template_with_formulas.xlsx
#   exports/alignment_template_with_formulas.xlsx
#
# Output:
#   exports/<TAG>/<TAG>_alignment.xlsx
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(readr)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# (0) CLI (NO TAG EDIT)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 3) {
  stop(
    "Usage:\n",
    "  Rscript scripts/04_build_alignment_from_template.R <wrist_tracks_csv> <watch_points_csv> <gesture_events_csv> [fps] [dt_seconds]\n\n",
    "Example:\n",
    "  Rscript scripts/04_build_alignment_from_template.R exports/wrist_tracks/wrist_tracks_m10.csv exports/m10/watch_points_m10_teacher.csv exports/m10/gesture_events_m10.csv 30 1.5\n",
    call. = FALSE
  )
}

wrist_csv  <- args[[1]]
watch_csv  <- args[[2]]
gesture_csv <- args[[3]]

fps        <- if (length(args) >= 4) as.numeric(args[[4]]) else 30
dt_seconds <- if (length(args) >= 5) as.numeric(args[[5]]) else 1.5
dt_grid    <- c(0.5, 1.0, 1.5)  # reserved (not used here)

stopifnot(is.finite(fps), fps > 0)
stopifnot(is.finite(dt_seconds), dt_seconds > 0)

for (p in c(wrist_csv, watch_csv, gesture_csv)) {
  if (!file.exists(p)) stop("Missing input file:\n  ", p, call. = FALSE)
}

# ----------------------------
# (0.1) PROJECT ROOT (repo root)
# ----------------------------
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = "")
if (!nzchar(project_root)) project_root <- getwd()
exports_root <- file.path(project_root, "exports")
dir.create(exports_root, showWarnings = FALSE, recursive = TRUE)

# ----------------------------
# (0.2) AUTO TAG (from input filenames)
# ----------------------------
extract_tag <- function(path) {
  b <- tolower(basename(path))
  m <- str_match(b, "(_|-)m([0-9]+)\\b")[,2]
  if (!is.na(m) && nzchar(m)) return(paste0("m", m))
  # fallback: try wrist_tracks_<tag>.csv
  m2 <- str_match(b, "wrist[_-]?tracks[_-]?([a-z0-9]+)\\.(csv|txt)$")[,2]
  if (!is.na(m2) && nzchar(m2)) return(m2)
  NA_character_
}

tag_candidates <- na.omit(c(extract_tag(wrist_csv), extract_tag(watch_csv), extract_tag(gesture_csv)))
TAG <- if (length(tag_candidates) > 0) tag_candidates[[1]] else "tag"

tag_dir <- file.path(exports_root, TAG)
dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)

# ----------------------------
# (0.3) TEMPLATE (auto-find in exports/)
# ----------------------------
template_candidates <- c(
  file.path(exports_root, "_alignment_template_with_formulas.xlsx"),
  file.path(exports_root, "alignment_template_with_formulas.xlsx")
)
template_xlsx <- template_candidates[file.exists(template_candidates)][1]
if (is.na(template_xlsx) || !nzchar(template_xlsx)) {
  stop(
    "Template not found in exports/. Tried:\n  - ",
    paste(template_candidates, collapse = "\n  - "),
    call. = FALSE
  )
}

out_xlsx <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

message("\n[INFO] project_root : ", normalizePath(project_root, winslash = "/", mustWork = FALSE))
message("[INFO] TAG          : ", TAG)
message("[INFO] Template     : ", normalizePath(template_xlsx, winslash = "/", mustWork = FALSE))
message("[INFO] WristTracks  : ", normalizePath(wrist_csv, winslash = "/", mustWork = FALSE))
message("[INFO] WatchPoints  : ", normalizePath(watch_csv, winslash = "/", mustWork = FALSE))
message("[INFO] GestureEvents: ", normalizePath(gesture_csv, winslash = "/", mustWork = FALSE))
message("[INFO] Output       : ", normalizePath(out_xlsx, winslash = "/", mustWork = FALSE), "\n")

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

apply_rename_map <- function(df, rename_map) {
  for (old in names(rename_map)) {
    new <- rename_map[[old]]
    if (old %in% names(df) && !(new %in% names(df))) names(df)[names(df) == old] <- new
  }
  df
}

safe_read_csv <- function(path) {
  readr::read_csv(path, show_col_types = FALSE, progress = FALSE)
}

ensure_cols <- function(df, cols, default = NA) {
  for (cc in cols) if (!(cc %in% names(df))) df[[cc]] <- default
  df
}

write_df_under_header <- function(wb, sheet, df,
                                  max_rows_to_clear = 2000,
                                  max_cols_to_clear = 60) {
  if (is.null(df)) return(invisible(FALSE))
  if (!(sheet %in% names(wb))) return(invisible(FALSE))
  
  df <- as.data.frame(df)
  
  openxlsx::writeData(
    wb, sheet,
    x = matrix("", nrow = max_rows_to_clear, ncol = max_cols_to_clear),
    startRow = 2, startCol = 1,
    colNames = FALSE, rowNames = FALSE
  )
  openxlsx::writeData(
    wb, sheet,
    x = df,
    startRow = 2, startCol = 1,
    colNames = FALSE, rowNames = FALSE,
    keepNA = TRUE
  )
  invisible(TRUE)
}

# ----------------------------
# (2) READ + NORMALIZE
# ----------------------------
gesture_df <- safe_read_csv(gesture_csv)
struct_df  <- safe_read_csv(watch_csv)
wrist_df   <- safe_read_csv(wrist_csv)

names(gesture_df) <- std_names(names(gesture_df))
names(struct_df)  <- std_names(names(struct_df))
names(wrist_df)   <- std_names(names(wrist_df))

# ----------------------------
# (2A) GestureEvents schema
# ----------------------------
expected_gesture <- c(
  "eventid","run","start_frame","end_frame","duration_frames",
  "start_sec","end_sec","duration_sec",
  "mean_speed","max_speed",
  "peak_frame","peak_sec","peak_sec_proxy",
  "q","threshold"
)

gesture_df <- apply_rename_map(gesture_df, c(
  "event_id" = "eventid",
  "id" = "eventid",
  "mean_vel" = "mean_speed",
  "max_vel" = "max_speed",
  "peak_t" = "peak_sec",
  "peakframe" = "peak_frame"
))

if (!("eventid" %in% names(gesture_df))) gesture_df$eventid <- sprintf("E%04d", seq_len(nrow(gesture_df)))
if (!("start_sec" %in% names(gesture_df)) && "start_frame" %in% names(gesture_df)) gesture_df$start_sec <- as.numeric(gesture_df$start_frame) / fps
if (!("end_sec" %in% names(gesture_df)) && "end_frame" %in% names(gesture_df)) gesture_df$end_sec <- as.numeric(gesture_df$end_frame) / fps
if (!("duration_sec" %in% names(gesture_df)) && all(c("start_sec","end_sec") %in% names(gesture_df))) gesture_df$duration_sec <- as.numeric(gesture_df$end_sec) - as.numeric(gesture_df$start_sec)
if (!("duration_frames" %in% names(gesture_df)) && all(c("start_frame","end_frame") %in% names(gesture_df))) gesture_df$duration_frames <- as.numeric(gesture_df$end_frame) - as.numeric(gesture_df$start_frame)

if (!("peak_sec" %in% names(gesture_df))) {
  if (all(c("start_sec","end_sec") %in% names(gesture_df))) gesture_df$peak_sec <- (as.numeric(gesture_df$start_sec) + as.numeric(gesture_df$end_sec)) / 2
  else gesture_df$peak_sec <- NA_real_
}
if (!("peak_sec_proxy" %in% names(gesture_df))) gesture_df$peak_sec_proxy <- as.numeric(gesture_df$peak_sec)
if (!("peak_frame" %in% names(gesture_df))) gesture_df$peak_frame <- NA_real_

gesture_df <- ensure_cols(gesture_df, expected_gesture, default = NA) %>%
  select(all_of(expected_gesture)) %>%
  mutate(
    eventid = as.character(eventid),
    peak_sec = suppressWarnings(as.numeric(peak_sec)),
    start_sec = suppressWarnings(as.numeric(start_sec)),
    end_sec = suppressWarnings(as.numeric(end_sec)),
    max_speed = suppressWarnings(as.numeric(max_speed)),
    duration_sec = suppressWarnings(as.numeric(duration_sec))
  ) %>%
  distinct(eventid, .keep_all = TRUE) %>%
  arrange(peak_sec)

# ----------------------------
# (2B) WatchPoints schema
# ----------------------------
expected_struct <- c(
  "pointid","time_s","time_start","time_end",
  "teacher_ratio","speed","transcript_cue","macro","micro_codes","confidence_1to3","notes"
)

struct_df <- apply_rename_map(struct_df, c(
  "point_id" = "pointid",
  "id" = "pointid",
  "time" = "time_s",
  "time_sec" = "time_s",
  "t_sec" = "time_s",
  "macro_code" = "macro",
  "micro" = "micro_codes",
  "confidence" = "confidence_1to3",
  "cue" = "transcript_cue"
))

if (!("pointid" %in% names(struct_df))) struct_df$pointid <- sprintf("P%03d", seq_len(nrow(struct_df)))
if (!("time_start" %in% names(struct_df)) && "time_s" %in% names(struct_df)) struct_df$time_start <- struct_df$time_s
if (!("time_end" %in% names(struct_df)) && "time_s" %in% names(struct_df)) struct_df$time_end <- struct_df$time_s

struct_df <- ensure_cols(struct_df, expected_struct, default = NA) %>%
  select(all_of(expected_struct)) %>%
  mutate(
    pointid = as.character(pointid),
    time_s = suppressWarnings(as.numeric(time_s)),
    teacher_ratio = suppressWarnings(as.numeric(teacher_ratio)),
    confidence_1to3 = suppressWarnings(as.numeric(confidence_1to3))
  ) %>%
  filter(is.finite(time_s)) %>%
  arrange(time_s)

# ----------------------------
# (2C) WristTracks checks (Methods countsflow)
# ----------------------------
# Script 3 gating params (MUST MATCH Script 3)
conf_min         <- 0.35
shoulder_vis_min <- 0.45
shoulder_ok_mode <- "both"  # "both" or "either"

need_wrist_cols <- c("conf_det", "ls_vis", "rs_vis")
missing_wrist <- setdiff(need_wrist_cols, names(wrist_df))
if (length(missing_wrist) > 0) {
  stop("WristTracks missing required cols: ",
       paste(missing_wrist, collapse = ", "),
       "\nNeed conf_det, ls_vis, rs_vis (from Script 3 schema).",
       call. = FALSE)
}

if (!("t_sec" %in% names(wrist_df))) {
  if ("frame" %in% names(wrist_df)) wrist_df$t_sec <- as.numeric(wrist_df$frame) / fps
  else wrist_df$t_sec <- NA_real_
}

# ----------------------------
# (3) AUTO-MATCH (nearest peak within dt)
# ----------------------------
event_peaks <- gesture_df$peak_sec

match_point <- function(t0) {
  if (!is.finite(t0) || length(event_peaks) == 0 || all(!is.finite(event_peaks))) return(list(ok = FALSE))
  j <- which.min(abs(event_peaks - t0))
  lag <- as.numeric(event_peaks[j] - t0)
  ok  <- is.finite(lag) && abs(lag) <= dt_seconds
  if (!ok) return(list(ok = FALSE))
  
  ev <- gesture_df[j, , drop = FALSE]
  list(
    ok = TRUE,
    eventid = as.character(ev$eventid),
    event_peak_sec = as.numeric(ev$peak_sec),
    lag_sec = lag,
    event_start_sec = as.numeric(ev$start_sec),
    event_end_sec = as.numeric(ev$end_sec),
    event_max_speed = as.numeric(ev$max_speed),
    event_duration_sec = as.numeric(ev$duration_sec)
  )
}

m_list <- lapply(struct_df$time_s, match_point)

aligned_tbl <- struct_df %>%
  mutate(
    EventID_auto = vapply(m_list, function(x) if (isTRUE(x$ok)) x$eventid else NA_character_, character(1)),
    event_peak_sec = vapply(m_list, function(x) if (isTRUE(x$ok)) x$event_peak_sec else NA_real_, numeric(1)),
    lag_sec        = vapply(m_list, function(x) if (isTRUE(x$ok)) x$lag_sec else NA_real_, numeric(1)),
    event_start_sec= vapply(m_list, function(x) if (isTRUE(x$ok)) x$event_start_sec else NA_real_, numeric(1)),
    event_end_sec  = vapply(m_list, function(x) if (isTRUE(x$ok)) x$event_end_sec else NA_real_, numeric(1)),
    event_max_speed= vapply(m_list, function(x) if (isTRUE(x$ok)) x$event_max_speed else NA_real_, numeric(1)),
    event_duration_sec = vapply(m_list, function(x) if (isTRUE(x$ok)) x$event_duration_sec else NA_real_, numeric(1))
  ) %>%
  select(
    pointid, time_s,
    teacher_ratio, macro, micro_codes, confidence_1to3,
    EventID_auto,
    event_peak_sec, lag_sec,
    event_start_sec, event_end_sec,
    event_max_speed, event_duration_sec
  )

# ----------------------------
# (4) SUMMARY
# ----------------------------
q_val <- NA_real_
if ("q" %in% names(gesture_df)) {
  q_tmp <- suppressWarnings(as.numeric(gesture_df$q))
  q_tmp <- q_tmp[is.finite(q_tmp)]
  if (length(q_tmp) > 0) q_val <- q_tmp[[1]]
}

n_events <- nrow(gesture_df)

total_sec <- suppressWarnings(max(as.numeric(wrist_df$t_sec), na.rm = TRUE))
if (!is.finite(total_sec)) total_sec <- suppressWarnings(max(as.numeric(gesture_df$end_sec), na.rm = TRUE))
total_min <- total_sec / 60
events_per_min <- ifelse(is.finite(total_min) && total_min > 0, n_events / total_min, NA_real_)

mean_amp <- suppressWarnings(mean(as.numeric(gesture_df$max_speed), na.rm = TRUE))
sd_amp   <- suppressWarnings(sd(as.numeric(gesture_df$max_speed), na.rm = TRUE))
mean_dur <- suppressWarnings(mean(as.numeric(gesture_df$duration_sec), na.rm = TRUE))
sd_dur   <- suppressWarnings(sd(as.numeric(gesture_df$duration_sec), na.rm = TRUE))

summary_row <- tibble(
  TAG = TAG,
  q = q_val,
  n_events = n_events,
  events_per_min = events_per_min,
  mean_amplitude = mean_amp,
  sd_amplitude = sd_amp,
  mean_duration = mean_dur,
  sd_duration = sd_dur
)

# ----------------------------
# (4.1) COUNTSFLOW (Methods sheet)
# ----------------------------
conf_ok <- !is.na(wrist_df$conf_det) & (as.numeric(wrist_df$conf_det) >= conf_min)
ls_ok   <- !is.na(wrist_df$ls_vis)   & (as.numeric(wrist_df$ls_vis)   >= shoulder_vis_min)
rs_ok   <- !is.na(wrist_df$rs_vis)   & (as.numeric(wrist_df$rs_vis)   >= shoulder_vis_min)

shoulders_ok <- if (shoulder_ok_mode == "either") (ls_ok | rs_ok) else (ls_ok & rs_ok)
teacher_present <- conf_ok & shoulders_ok

total_frames           <- nrow(wrist_df)
conf_ok_frames         <- sum(conf_ok, na.rm = TRUE)
shoulders_ok_frames    <- sum(shoulders_ok, na.rm = TRUE)
teacher_present_frames <- sum(teacher_present, na.rm = TRUE)
n_candidates           <- teacher_present_frames
n_watch_final          <- nrow(struct_df)

# ----------------------------
# (5) LOAD TEMPLATE + WRITE
# ----------------------------
wb <- openxlsx::loadWorkbook(template_xlsx)

if ("Inputs" %in% names(wb)) {
  openxlsx::writeData(wb, "Inputs", x = TAG,        startCol = 2, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Inputs", x = fps,        startCol = 2, startRow = 3, colNames = FALSE)
  openxlsx::writeData(wb, "Inputs", x = dt_seconds, startCol = 2, startRow = 4, colNames = FALSE)
}

write_df_under_header(wb, "GestureEvents",    gesture_df, max_rows_to_clear = 3000,  max_cols_to_clear = 40)
write_df_under_header(wb, "StructuralPoints", struct_df,  max_rows_to_clear = 2000,  max_cols_to_clear = 40)
write_df_under_header(wb, "WristTracks",      wrist_df,   max_rows_to_clear = 60000, max_cols_to_clear = 80)

# DebugMissing is optional: only write if sheet exists AND user has a file
# (We keep it simple: not required for running)
# If you want to support it, pass a 4th csv path and write it here.

if ("Alignment" %in% names(wb)) {
  n_points  <- nrow(aligned_tbl)
  start_row <- 2
  max_clear <- max(500, n_points + 50)
  
  cols_to_clear <- c(1,2,4,5,6,7,8,9,10,12,13,14,15)
  for (cc in cols_to_clear) {
    openxlsx::writeData(wb, "Alignment", x = rep("", max_clear),
                        startCol = cc, startRow = start_row, colNames = FALSE)
  }
  
  if (n_points > 0) {
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$pointid,            startCol = 1,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$time_s,             startCol = 2,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$EventID_auto,       startCol = 4,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_peak_sec,     startCol = 5,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$lag_sec,            startCol = 6,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_start_sec,    startCol = 7,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_end_sec,      startCol = 8,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_max_speed,    startCol = 9,  startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_duration_sec, startCol = 10, startRow = start_row, colNames = FALSE)
    
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$teacher_ratio,      startCol = 12, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$macro,              startCol = 13, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$micro_codes,        startCol = 14, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$confidence_1to3,    startCol = 15, startRow = start_row, colNames = FALSE)
  }
}

if ("Summary" %in% names(wb)) {
  openxlsx::writeData(wb, "Summary", x = rep("", 80), startCol = 1, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Summary", x = summary_row, startCol = 1, startRow = 2, colNames = TRUE)
}

if ("Methods" %in% names(wb)) {
  openxlsx::writeData(wb, "Methods", x = TAG,        startCol = 3, startRow = 2,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = fps,        startCol = 3, startRow = 3,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = dt_seconds, startCol = 3, startRow = 4,  colNames = FALSE)
  
  openxlsx::writeData(wb, "Methods", x = total_frames,           startCol = 3, startRow = 5,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = teacher_present_frames, startCol = 3, startRow = 6,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = conf_ok_frames,         startCol = 3, startRow = 7,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = shoulders_ok_frames,    startCol = 3, startRow = 8,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = n_candidates,           startCol = 3, startRow = 9,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = n_watch_final,          startCol = 3, startRow = 10, colNames = FALSE)
  
  openxlsx::writeData(wb, "Methods", x = n_events,       startCol = 3, startRow = 11, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = events_per_min, startCol = 3, startRow = 12, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = mean_amp,       startCol = 3, startRow = 13, colNames = FALSE)
}

# Robustness table (optional)
robust_csv <- file.path(tag_dir, paste0("robust_", TAG, ".csv"))
if (file.exists(robust_csv) && "Robustness" %in% names(wb)) {
  robust_df <- readr::read_csv(robust_csv, show_col_types = FALSE, progress = FALSE)
  names(robust_df) <- std_names(names(robust_df))
  
  max_clear <- max(100, nrow(robust_df) + 20)
  openxlsx::writeData(wb, "Robustness",
                      x = matrix("", nrow = max_clear, ncol = ncol(robust_df)),
                      startRow = 2, startCol = 1, colNames = FALSE, rowNames = FALSE)
  openxlsx::writeData(wb, "Robustness",
                      x = robust_df,
                      startRow = 2, startCol = 1, colNames = FALSE, rowNames = FALSE, keepNA = TRUE)
}

# ----------------------------
# (6) SAVE
# ----------------------------
openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)

matched_n <- sum(!is.na(aligned_tbl$EventID_auto))

message("\nDONE -> ", normalizePath(out_xlsx, winslash = "/", mustWork = FALSE))
message(sprintf("AUTO-MATCH: %d points | matched %d within dt=%.3fs",
                nrow(aligned_tbl), matched_n, dt_seconds))
message(sprintf("Methods countsflow: total=%d | teacher_present=%d | conf_ok=%d | shoulders_ok=%d | n_candidates=%d | n_watch_final=%d",
                total_frames, teacher_present_frames, conf_ok_frames, shoulders_ok_frames, n_candidates, n_watch_final))
