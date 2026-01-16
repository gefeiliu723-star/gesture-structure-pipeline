# ============================================================
# 04_build_alignment_from_template_with_formulas.R (FINAL CLEAN)
#
# Clean rules:
#   - TAG is REQUIRED (no default m02)
#   - No absolute paths
#   - Project root: env GESTURE_PROJECT_ROOT (fallback: getwd())
#   - Template path MUST be provided:
#       (1) CLI: --template path/to/template.xlsx
#       (2) env: ALIGNMENT_TEMPLATE_XLSX=path/to/template.xlsx
#     (No guessing template names)
#
# Inputs (expected under exports/<TAG>/):
#   - gesture_events_<TAG>.csv
#   - watch_points_<TAG>_teacher.csv   (or watch_points / structural_points variants)
#   - teacher_qc_<TAG>.csv             (optional)
#   - debug_missing_reason_<TAG>_q*.csv (optional)
#   - robust_<TAG>.csv                 (optional; fills Robustness A:F)
#   - wrist_tracks_<TAG>.csv           (optional; either exports/wrist_tracks/ or exports/<TAG>/)
#
# Output:
#   - exports/<TAG>/<TAG>_alignment.xlsx
#
# Usage:
#   Rscript scripts/04_build_alignment_from_template_with_formulas.R <TAG> [fps] [dt_seconds] --template <xlsx>
# Example:
#   Rscript scripts/04_build_alignment_from_template_with_formulas.R m02 30 1.5 --template exports/my_any_name.xlsx
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(readr)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# (0) ARGS (TAG required) + template required
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

get_arg_value <- function(flag, args) {
  i <- which(args == flag)
  if (length(i) == 0) return(NA_character_)
  if (i[[1]] >= length(args)) return(NA_character_)
  args[[i[[1]] + 1]]
}

if (length(args) < 1 || !nzchar(args[[1]])) {
  stop(
    "Missing TAG.\n",
    "Usage:\n  Rscript scripts/04_build_alignment_from_template_with_formulas.R <TAG> [fps] [dt_seconds] --template <xlsx>\n",
    call. = FALSE
  )
}

TAG <- args[[1]]
fps <- if (length(args) >= 2 && nzchar(args[[2]]) && args[[2]] != "--template") as.numeric(args[[2]]) else 30
dt_seconds <- if (length(args) >= 3 && nzchar(args[[3]]) && args[[3]] != "--template") as.numeric(args[[3]]) else 1.5

stopifnot(is.finite(fps), fps > 0, is.finite(dt_seconds), dt_seconds > 0)

template_xlsx <- get_arg_value("--template", args)
if (is.na(template_xlsx) || !nzchar(template_xlsx)) {
  template_xlsx <- Sys.getenv("ALIGNMENT_TEMPLATE_XLSX", unset = NA_character_)
}
if (is.na(template_xlsx) || !nzchar(template_xlsx) || !file.exists(template_xlsx)) {
  stop(
    "Template not found.\n",
    "Provide via:\n",
    "  (1) CLI: --template <path>\n",
    "  (2) env: ALIGNMENT_TEMPLATE_XLSX=<path>\n",
    call. = FALSE
  )
}

# ----------------------------
# (0.1) PROJECT ROOT (portable)
# ----------------------------
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
tag_dir      <- file.path(exports_root, TAG)
dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)

out_xlsx <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

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

stop_if_missing <- function(path, what = "path") {
  if (!file.exists(path) && !dir.exists(path)) {
    stop(sprintf("Missing %s: %s", what, path), call. = FALSE)
  }
}

pick_file <- function(dir, patterns, optional = FALSE) {
  if (!dir.exists(dir)) {
    if (optional) return(NA_character_)
    stop("Folder not found: ", dir, call. = FALSE)
  }
  files <- list.files(dir, full.names = TRUE, recursive = FALSE)
  if (length(files) == 0) {
    if (optional) return(NA_character_)
    stop("No files in: ", dir, call. = FALSE)
  }
  
  low_base <- tolower(basename(files))
  hit <- character(0)
  for (pat in patterns) {
    idx <- which(str_detect(low_base, pat))
    if (length(idx) > 0) { hit <- files[idx]; break }
  }
  
  if (length(hit) == 0) {
    if (optional) return(NA_character_)
    stop("Could not find file in ", dir, " with patterns: ",
         paste(patterns, collapse = " | "), call. = FALSE)
  }
  
  if (length(hit) > 1) {
    hit <- hit[order(file.info(hit)$mtime, decreasing = TRUE)]
  }
  hit[[1]]
}

pick_file_multi <- function(dirs, patterns, optional = FALSE) {
  for (d in dirs) {
    p <- pick_file(d, patterns, optional = TRUE)
    if (!is.na(p) && file.exists(p)) return(p)
  }
  if (optional) return(NA_character_)
  stop("Could not find required file across dirs: ", paste(dirs, collapse = " , "),
       " | patterns: ", paste(patterns, collapse = " | "), call. = FALSE)
}

safe_read_csv <- function(path) {
  if (is.na(path) || !nzchar(path) || !file.exists(path)) return(NULL)
  readr::read_csv(path, show_col_types = FALSE, progress = FALSE)
}

ensure_cols <- function(df, cols, default = NA) {
  for (cc in cols) if (!(cc %in% names(df))) df[[cc]] <- default
  df
}

apply_rename_map <- function(df, rename_map) {
  for (old in names(rename_map)) {
    new <- rename_map[[old]]
    if (old %in% names(df) && !(new %in% names(df))) {
      names(df)[names(df) == old] <- new
    }
  }
  df
}

# write df under header row (row 1) without overwriting headers
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

num1 <- function(x) {
  v <- suppressWarnings(as.numeric(x))
  v <- v[is.finite(v)]
  if (length(v) == 0) return(NA_real_)
  v[[1]]
}

# ----------------------------
# (2) LOCATE INPUT FILES
# ----------------------------
stop_if_missing(exports_root, "exports folder")
stop_if_missing(tag_dir, paste0("exports/", TAG, " folder"))
stop_if_missing(template_xlsx, "template xlsx")

gesture_csv <- pick_file(
  tag_dir,
  patterns = c("^gesture_events_.*\\.csv$", "gesture.*events.*\\.csv$")
)

struct_csv <- pick_file(
  tag_dir,
  patterns = c("^watch_points_.*\\.csv$", "watch[_-]?points.*\\.csv$",
               "^structural_points_.*\\.csv$", "structural[_-]?points.*\\.csv$")
)

wrist_csv <- pick_file_multi(
  dirs = c(file.path(exports_root, "wrist_tracks"), tag_dir),
  patterns = c(paste0("^wrist_tracks_", tolower(TAG), "\\.csv$"),
               paste0("wrist[_-]?tracks.*", tolower(TAG), ".*\\.csv$")),
  optional = TRUE
)

teacher_qc_csv <- pick_file(
  tag_dir,
  patterns = c("^teacher_qc_.*\\.csv$", "teacher[_-]?qc.*\\.csv$"),
  optional = TRUE
)

debug_csv <- pick_file(
  tag_dir,
  patterns = c("^debug_missing_reason_.*\\.csv$", "debug[_-]?missing.*\\.csv$"),
  optional = TRUE
)

robust_csv <- pick_file(
  tag_dir,
  patterns = c("^robust_.*\\.csv$", "robust.*\\.csv$"),
  optional = TRUE
)

message("Using files:")
message("  Template     : ", template_xlsx)
message("  GestureEvents: ", gesture_csv)
message("  StructuralPts: ", struct_csv)
message("  WristTracks  : ", wrist_csv)
message("  TeacherQC    : ", teacher_qc_csv)
message("  DebugMissing : ", debug_csv)
message("  Robust       : ", robust_csv)

# ----------------------------
# (3) READ + NORMALIZE TABLES
# ----------------------------
gesture_df <- safe_read_csv(gesture_csv)
struct_df  <- safe_read_csv(struct_csv)
wrist_df   <- safe_read_csv(wrist_csv)
qc_df      <- safe_read_csv(teacher_qc_csv)
debug_df   <- safe_read_csv(debug_csv)
robust_df  <- safe_read_csv(robust_csv)

stopifnot(!is.null(gesture_df), !is.null(struct_df))

names(gesture_df) <- std_names(names(gesture_df))
names(struct_df)  <- std_names(names(struct_df))
if (!is.null(wrist_df))  names(wrist_df)  <- std_names(names(wrist_df))
if (!is.null(qc_df))     names(qc_df)     <- std_names(names(qc_df))
if (!is.null(debug_df))  names(debug_df)  <- std_names(names(debug_df))
if (!is.null(robust_df)) names(robust_df) <- std_names(names(robust_df))

# ----------------------------
# (3A) GestureEvents schema
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
  "eventid_" = "eventid",
  "event_id_" = "eventid",
  "id" = "eventid",
  "mean_vel" = "mean_speed",
  "max_vel" = "max_speed",
  "peak_t" = "peak_sec",
  "peakframe" = "peak_frame"
))

if (!("eventid" %in% names(gesture_df))) {
  gesture_df$eventid <- sprintf("E%04d", seq_len(nrow(gesture_df)))
}

if (!("start_sec" %in% names(gesture_df)) && "start_frame" %in% names(gesture_df)) {
  gesture_df$start_sec <- as.numeric(gesture_df$start_frame) / fps
}
if (!("end_sec" %in% names(gesture_df)) && "end_frame" %in% names(gesture_df)) {
  gesture_df$end_sec <- as.numeric(gesture_df$end_frame) / fps
}
if (!("duration_sec" %in% names(gesture_df)) && all(c("start_sec","end_sec") %in% names(gesture_df))) {
  gesture_df$duration_sec <- as.numeric(gesture_df$end_sec) - as.numeric(gesture_df$start_sec)
}
if (!("duration_frames" %in% names(gesture_df)) && all(c("start_frame","end_frame") %in% names(gesture_df))) {
  gesture_df$duration_frames <- as.numeric(gesture_df$end_frame) - as.numeric(gesture_df$start_frame)
}

if (!("peak_sec" %in% names(gesture_df))) {
  if (all(c("start_sec","end_sec") %in% names(gesture_df))) {
    gesture_df$peak_sec <- (as.numeric(gesture_df$start_sec) + as.numeric(gesture_df$end_sec)) / 2
  } else {
    gesture_df$peak_sec <- NA_real_
  }
}
if (!("peak_sec_proxy" %in% names(gesture_df))) gesture_df$peak_sec_proxy <- as.numeric(gesture_df$peak_sec)
if (!("peak_frame" %in% names(gesture_df))) gesture_df$peak_frame <- NA_real_

gesture_df <- ensure_cols(gesture_df, expected_gesture, default = NA) %>%
  select(all_of(expected_gesture)) %>%
  mutate(
    eventid = as.character(eventid),
    peak_sec = as.numeric(peak_sec),
    start_sec = as.numeric(start_sec),
    end_sec = as.numeric(end_sec),
    max_speed = as.numeric(max_speed),
    duration_sec = as.numeric(duration_sec)
  ) %>%
  distinct(eventid, .keep_all = TRUE) %>%
  arrange(peak_sec)

# ----------------------------
# (3B) StructuralPoints schema
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
  "t_struct_sec" = "time_s",
  "macro_code" = "macro",
  "micro" = "micro_codes",
  "confidence" = "confidence_1to3",
  "cue" = "transcript_cue"
))

if (!("pointid" %in% names(struct_df))) {
  struct_df$pointid <- sprintf("P%03d", seq_len(nrow(struct_df)))
}
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
# (4) AUTO-MATCH helper (dt parameter)
# ----------------------------
match_one <- function(t_point, events_df, dt) {
  if (!is.finite(t_point) || nrow(events_df) == 0) return(NULL)
  lag_vec <- events_df$peak_sec - t_point
  j <- which.min(abs(lag_vec))
  ev <- events_df[j, , drop = FALSE]
  lag <- as.numeric(ev$peak_sec) - t_point
  ok <- is.finite(lag) && is.finite(dt) && abs(lag) <= dt
  
  list(
    eventid = as.character(ev$eventid),
    event_peak_sec = as.numeric(ev$peak_sec),
    lag_sec = as.numeric(lag),
    event_start_sec = as.numeric(ev$start_sec),
    event_end_sec = as.numeric(ev$end_sec),
    event_max_speed = as.numeric(ev$max_speed),
    event_duration_sec = as.numeric(ev$duration_sec),
    ok = ok
  )
}

aligned_tbl <- struct_df %>%
  rowwise() %>%
  mutate(.m = list(match_one(time_s, gesture_df, dt_seconds))) %>%
  mutate(
    EventID_auto = ifelse(!is.null(.m) && .m$ok, .m$eventid, NA_character_),
    event_peak_sec = ifelse(!is.null(.m) && .m$ok, .m$event_peak_sec, NA_real_),
    lag_sec        = ifelse(!is.null(.m) && .m$ok, .m$lag_sec, NA_real_),
    event_start_sec= ifelse(!is.null(.m) && .m$ok, .m$event_start_sec, NA_real_),
    event_end_sec  = ifelse(!is.null(.m) && .m$ok, .m$event_end_sec, NA_real_),
    event_max_speed= ifelse(!is.null(.m) && .m$ok, .m$event_max_speed, NA_real_),
    event_duration_sec = ifelse(!is.null(.m) && .m$ok, .m$event_duration_sec, NA_real_)
  ) %>%
  ungroup() %>%
  select(
    pointid, time_s,
    teacher_ratio, macro, micro_codes, confidence_1to3,
    EventID_auto,
    event_peak_sec, lag_sec,
    event_start_sec, event_end_sec,
    event_max_speed, event_duration_sec
  )

# ----------------------------
# (5) SUMMARY + COUNTSFLOW
# ----------------------------
q_val <- NA_real_
if ("q" %in% names(gesture_df)) {
  q_val <- suppressWarnings(as.numeric(gesture_df$q))
  q_val <- q_val[is.finite(q_val)]
  if (length(q_val) > 0) q_val <- q_val[[1]]
}

n_events <- nrow(gesture_df)

total_sec <- NA_real_
if (!is.null(wrist_df) && "t_sec" %in% names(wrist_df)) {
  total_sec <- suppressWarnings(max(as.numeric(wrist_df$t_sec), na.rm = TRUE))
}
if (!is.finite(total_sec)) {
  total_sec <- suppressWarnings(max(as.numeric(gesture_df$end_sec), na.rm = TRUE))
}
total_min <- total_sec / 60
events_per_min <- ifelse(is.finite(total_min) && total_min > 0, n_events / total_min, NA_real_)

mean_amp <- suppressWarnings(mean(as.numeric(gesture_df$max_speed), na.rm = TRUE))
mean_dur <- suppressWarnings(mean(as.numeric(gesture_df$duration_sec), na.rm = TRUE))

summary_row <- tibble(
  TAG = TAG,
  q = q_val,
  n_events = n_events,
  events_per_min = events_per_min,
  mean_amplitude = mean_amp,
  mean_duration = mean_dur
)

countsflow <- list(
  total_frames = NA_real_,
  teacher_present_frames = NA_real_,
  conf_ok_frames = NA_real_,
  shoulders_ok_frames = NA_real_,
  n_candidates = NA_real_,
  n_watch_final = NA_real_
)

if (!is.null(qc_df) && nrow(qc_df) > 0) {
  qc_nms <- names(qc_df)
  if ("total_frames" %in% qc_nms) countsflow$total_frames <- num1(qc_df$total_frames)
  if ("teacher_present_frames" %in% qc_nms) countsflow$teacher_present_frames <- num1(qc_df$teacher_present_frames)
  if ("conf_ok_frames" %in% qc_nms) countsflow$conf_ok_frames <- num1(qc_df$conf_ok_frames)
  if ("shoulders_ok_frames" %in% qc_nms) countsflow$shoulders_ok_frames <- num1(qc_df$shoulders_ok_frames)
  if ("n_candidates" %in% qc_nms) countsflow$n_candidates <- num1(qc_df$n_candidates)
  if ("n_watch_final" %in% qc_nms) countsflow$n_watch_final <- num1(qc_df$n_watch_final)
}
if (!is.null(wrist_df) && nrow(wrist_df) > 0 && !is.finite(countsflow$total_frames)) {
  countsflow$total_frames <- nrow(wrist_df)
}

# ----------------------------
# (6) LOAD TEMPLATE + WRITE
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
write_df_under_header(wb, "DebugMissing",     debug_df,   max_rows_to_clear = 8000,  max_cols_to_clear = 50)

# ----------------------------
# (7) Alignment (write values only; keep template formulas elsewhere)
# ----------------------------
if ("Alignment" %in% names(wb)) {
  n_points  <- nrow(aligned_tbl)
  start_row <- 2
  max_clear <- max(500, n_points + 50)
  
  cols_to_clear <- c(1,2,4,5,6,7,8,9,10,12,13,14,15)
  for (cc in cols_to_clear) {
    openxlsx::writeData(
      wb, "Alignment",
      x = rep("", max_clear),
      startCol = cc, startRow = start_row,
      colNames = FALSE
    )
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
  openxlsx::writeData(wb, "Methods", x = TAG,        startCol = 3, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = fps,        startCol = 3, startRow = 3, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = dt_seconds, startCol = 3, startRow = 4, colNames = FALSE)
  
  openxlsx::writeData(wb, "Methods", x = countsflow$total_frames,           startCol = 3, startRow = 5,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = countsflow$teacher_present_frames, startCol = 3, startRow = 6,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = countsflow$conf_ok_frames,         startCol = 3, startRow = 7,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = countsflow$shoulders_ok_frames,    startCol = 3, startRow = 8,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = countsflow$n_candidates,           startCol = 3, startRow = 9,  colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = countsflow$n_watch_final,          startCol = 3, startRow = 10, colNames = FALSE)
  
  openxlsx::writeData(wb, "Methods", x = n_events,       startCol = 3, startRow = 11, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = events_per_min, startCol = 3, startRow = 12, colNames = FALSE)
  openxlsx::writeData(wb, "Methods", x = mean_amp,       startCol = 3, startRow = 13, colNames = FALSE)
}

# ----------------------------
# (7B) Robustness sheet
# Fill ONLY A:F from exports/<TAG>/robust_<TAG>.csv
# Keep H:I (dt_seconds + Note) untouched.
# ----------------------------
if (!is.null(robust_df) && "Robustness" %in% names(wb)) {
  
  # Expect columns from Script02 robust output:
  # TAG,q,n_events,events_per_min,mean_amplitude,mean_duration
  # Normalize & pick those six in order.
  robust_df <- robust_df %>%
    mutate(across(everything(), ~ .x)) # no-op, keep as tibble
  
  # ensure required columns exist (case normalized already)
  need <- c("tag","q","n_events","events_per_min","mean_amplitude","mean_duration")
  miss <- setdiff(need, names(robust_df))
  if (length(miss) == 0) {
    
    out_r <- robust_df %>%
      select(all_of(need)) %>%
      mutate(tag = as.character(tag)) %>%
      as.data.frame()
    
    start_row <- 2
    max_clear <- max(100, nrow(out_r) + 20)
    
    # clear A:F only (1..6)
    openxlsx::writeData(
      wb, "Robustness",
      x = matrix("", nrow = max_clear, ncol = 6),
      startRow = start_row, startCol = 1,
      colNames = FALSE, rowNames = FALSE
    )
    
    # write A:F only, no headers (header already in template)
    openxlsx::writeData(
      wb, "Robustness",
      x = out_r,
      startRow = start_row, startCol = 1,
      colNames = FALSE, rowNames = FALSE,
      keepNA = TRUE
    )
    
    message("Robustness A:F filled from: ", basename(robust_csv))
  } else {
    message("Robustness CSV found but missing required columns: ", paste(miss, collapse = ", "))
  }
}

# ----------------------------
# (8) SAVE
# ----------------------------
openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)

matched_n <- sum(!is.na(aligned_tbl$EventID_auto))
message("DONE -> ", normalizePath(out_xlsx, winslash = "/", mustWork = FALSE))
message(sprintf("AUTO-MATCH: %d points | matched %d within dt=%.3fs",
                nrow(aligned_tbl), matched_n, dt_seconds))
