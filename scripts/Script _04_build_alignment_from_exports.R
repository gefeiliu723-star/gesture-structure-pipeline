## ============================================================
# build_alignment_from_template_with_formulas.R  (GitHub-ready / AUTO-FILL)
#
# Repo layout expected (run from repo root):
#   exports/
#     alignment_template_with_formulas.xlsx
#     wrist_tracks/   (optional; wrist_tracks_<TAG>.csv may live here)
#     <TAG>/
#       gesture_events_<TAG>.csv
#       watch_points_<TAG>.csv   (or structural_points... etc)
#
# Usage:
#   Rscript scripts/build_alignment_from_template_with_formulas.R
#   Rscript scripts/build_alignment_from_template_with_formulas.R m08
#   Rscript scripts/build_alignment_from_template_with_formulas.R m08 30 1.5
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
## ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(readr)
  library(dplyr)
  library(stringr)
  library(tools)
})

# ----------------------------
# (0) USER SETTINGS (EDIT ONLY HERE)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG <- if (length(args) >= 1) args[[1]] else "m08"
fps <- if (length(args) >= 2) as.numeric(args[[2]]) else 30
dt_seconds <- if (length(args) >= 3) as.numeric(args[[3]]) else 1.5

dt_grid <- c(0.5, 1.0, 1.5)

# Portable project root:
#  1) env var GESTURE_PROJECT_ROOT
#  2) current working directory (assumed repo root)
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

exports_root <- file.path(project_root, "exports")
tag_dir      <- file.path(exports_root, TAG)

template_xlsx <- file.path(exports_root, "alignment_template_with_formulas.xlsx")
out_xlsx      <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

# ----------------------------
# (1) HELPERS
# ----------------------------
stop_if_missing <- function(path, what = "file/folder") {
  if (!file.exists(path) && !dir.exists(path)) {
    stop(sprintf("Missing %s: %s", what, path), call. = FALSE)
  }
}

std_names <- function(x) {
  x %>%
    str_replace_all("\\.+", "_") %>%
    str_replace_all("[^A-Za-z0-9_]", "_") %>%
    str_replace_all("_+", "_") %>%
    tolower()
}

find_one_in_dir <- function(dir, patterns, optional = FALSE) {
  if (!dir.exists(dir)) {
    if (optional) return(NA_character_)
    stop(sprintf("Folder not found: %s", dir), call. = FALSE)
  }
  files <- list.files(dir, full.names = TRUE, recursive = FALSE)
  hit <- character(0)
  for (pat in patterns) {
    hit <- files[str_detect(tolower(basename(files)), pat)]
    if (length(hit) > 0) break
  }
  if (length(hit) == 0) {
    if (optional) return(NA_character_)
    stop(sprintf("Could not find required file in %s. Tried patterns: %s",
                 dir, paste(patterns, collapse = " | ")), call. = FALSE)
  }
  if (length(hit) > 1) {
    hit <- hit[order(file.info(hit)$mtime, decreasing = TRUE)]
  }
  hit[[1]]
}

find_one_multi <- function(dirs, patterns, optional = FALSE) {
  for (d in dirs) {
    p <- find_one_in_dir(d, patterns, optional = TRUE)
    if (!is.na(p) && file.exists(p)) return(p)
  }
  if (optional) return(NA_character_)
  stop(sprintf("Could not find required file. Tried dirs: %s | patterns: %s",
               paste(dirs, collapse = " , "),
               paste(patterns, collapse = " | ")), call. = FALSE)
}

safe_read_csv <- function(path) {
  if (is.na(path) || !nzchar(path) || !file.exists(path)) return(NULL)
  readr::read_csv(path, show_col_types = FALSE, progress = FALSE)
}

# Write df under existing header row (row 1) WITHOUT writing header
write_df_under_header <- function(wb, sheet, df,
                                  max_rows_to_clear = 2000,
                                  max_cols_to_clear = 60) {
  if (is.null(df)) return(invisible(FALSE))
  if (!(sheet %in% names(wb))) return(invisible(FALSE))
  
  df <- as.data.frame(df)
  
  openxlsx::writeData(
    wb, sheet = sheet,
    x = matrix("", nrow = max_rows_to_clear, ncol = max_cols_to_clear),
    startRow = 2, startCol = 1, colNames = FALSE, rowNames = FALSE,
    keepNA = FALSE
  )
  
  openxlsx::writeData(
    wb, sheet = sheet, x = df,
    startRow = 2, startCol = 1,
    colNames = FALSE, rowNames = FALSE,
    keepNA = TRUE
  )
  invisible(TRUE)
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

ensure_cols <- function(df, cols, default = NA) {
  for (cc in cols) if (!(cc %in% names(df))) df[[cc]] <- default
  df
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
stop_if_missing(template_xlsx, "template xlsx")
stop_if_missing(tag_dir, paste0("tag folder exports/", TAG))

gesture_csv <- find_one_in_dir(
  tag_dir,
  patterns = c(
    "gesture.*events.*\\.csv$",
    "gesture_events.*\\.csv$",
    "gestureevents.*\\.csv$"
  )
)

struct_csv <- find_one_in_dir(
  tag_dir,
  patterns = c(
    "watch[_-]?points.*\\.csv$",
    "structural[_-]?points.*\\.csv$",
    "structure.*points.*\\.csv$",
    "discourse.*points.*\\.csv$"
  )
)

wrist_candidates <- c(
  tag_dir,
  file.path(exports_root, "wrist_tracks"),
  exports_root
)
wrist_csv <- find_one_multi(
  wrist_candidates,
  patterns = c(
    paste0("wrist[_-]?tracks.*", tolower(TAG), ".*\\.csv$"),
    paste0("wristtracks.*", tolower(TAG), ".*\\.csv$"),
    "wrist[_-]?tracks.*\\.csv$",
    "wristtracks.*\\.csv$",
    "tracks.*wrist.*\\.csv$"
  ),
  optional = TRUE
)

teacher_qc_csv <- find_one_in_dir(
  tag_dir,
  patterns = c(
    "teacher[_-]?qc.*\\.csv$",
    "teacherqc.*\\.csv$",
    "qc.*teacher.*\\.csv$"
  ),
  optional = TRUE
)

debug_candidates <- c(tag_dir, exports_root)
debug_csv <- find_one_multi(
  debug_candidates,
  patterns = c(
    paste0("debug[_-]?missing.*", tolower(TAG), ".*\\.csv$"),
    "debug[_-]?missing.*\\.csv$",
    "debugmissing.*\\.csv$",
    "debug_missing_reason.*\\.csv$"
  ),
  optional = TRUE
)

message("Using files:")
message("  GestureEvents : ", gesture_csv)
message("  StructuralPts : ", struct_csv)
message("  WristTracks   : ", wrist_csv)
message("  TeacherQC     : ", teacher_qc_csv)
message("  DebugMissing  : ", debug_csv)

# ----------------------------
# (3) READ + NORMALIZE
# ----------------------------
gesture_df <- safe_read_csv(gesture_csv)
struct_df  <- safe_read_csv(struct_csv)
wrist_df   <- safe_read_csv(wrist_csv)
qc_df      <- safe_read_csv(teacher_qc_csv)
debug_df   <- safe_read_csv(debug_csv)

if (is.null(gesture_df) || is.null(struct_df)) {
  stop("GestureEvents and StructuralPoints CSV are required.", call. = FALSE)
}

names(gesture_df) <- std_names(names(gesture_df))
names(struct_df)  <- std_names(names(struct_df))
if (!is.null(wrist_df)) names(wrist_df) <- std_names(names(wrist_df))
if (!is.null(qc_df))    names(qc_df)    <- std_names(names(qc_df))
if (!is.null(debug_df)) names(debug_df) <- std_names(names(debug_df))

# ----------------------------
# (3A) GestureEvents normalize to required schema
# ----------------------------
expected_gesture <- c(
  "eventid","run","start_frame","end_frame","duration_frames",
  "start_sec","end_sec","duration_sec",
  "mean_speed","max_speed",
  "peak_frame","peak_sec","peak_sec_proxy",
  "q","threshold"
)

rename_map_gesture <- c(
  "event_id"   = "eventid",
  "eventid"    = "eventid",
  "eventid_"   = "eventid",
  "event_id_"  = "eventid",
  "id"         = "eventid",
  "eventid_manual" = "eventid",
  "trial"      = "run",
  "segment"    = "run",
  "duration_f" = "duration_frames",
  "duration"   = "duration_sec",
  "dur"        = "duration_sec",
  "mean_vel"   = "mean_speed",
  "max_vel"    = "max_speed",
  "peak_t"     = "peak_sec",
  "peakframe"  = "peak_frame"
)

gesture_df <- apply_rename_map(gesture_df, rename_map_gesture)

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
if (!("peak_sec_proxy" %in% names(gesture_df))) {
  gesture_df$peak_sec_proxy <- as.numeric(gesture_df$peak_sec)
}

if (!("peak_frame" %in% names(gesture_df)) && "peak_fram" %in% names(gesture_df)) {
  gesture_df$peak_frame <- gesture_df$peak_fram
}
if (!("peak_frame" %in% names(gesture_df))) {
  gesture_df$peak_frame <- NA_real_
}

gesture_df <- ensure_cols(gesture_df, expected_gesture, default = NA)

gesture_df <- gesture_df %>%
  select(all_of(expected_gesture)) %>%
  mutate(
    eventid = as.character(eventid),
    peak_sec = as.numeric(peak_sec),
    peak_sec_proxy = as.numeric(peak_sec_proxy),
    start_sec = as.numeric(start_sec),
    end_sec = as.numeric(end_sec),
    max_speed = as.numeric(max_speed),
    duration_sec = as.numeric(duration_sec)
  ) %>%
  distinct(eventid, .keep_all = TRUE) %>%
  arrange(peak_sec)

# ----------------------------
# (3B) StructuralPoints normalize
# ----------------------------
expected_struct <- c(
  "pointid","time_s","time_start","time_end",
  "teacher_ratio","speed","transcript_cue","macro","micro_codes","confidence_1to3","notes"
)

rename_map_struct <- c(
  "point_id"        = "pointid",
  "id"              = "pointid",
  "watchpointid"    = "pointid",
  "t"               = "time_s",
  "time"            = "time_s",
  "time_sec"        = "time_s",
  "t_sec"           = "time_s",
  "time_start_s"    = "time_start",
  "time_end_s"      = "time_end",
  "teacher_present" = "teacher_ratio",
  "teacher_ratio_in_clip" = "teacher_ratio",
  "macro_code"      = "macro",
  "micro"           = "micro_codes",
  "micro_code"      = "micro_codes",
  "confidence"      = "confidence_1to3",
  "transcript"      = "transcript_cue",
  "cue"             = "transcript_cue"
)

struct_df <- apply_rename_map(struct_df, rename_map_struct)

if (!("pointid" %in% names(struct_df))) {
  struct_df$pointid <- sprintf("P%03d", seq_len(nrow(struct_df)))
}
if (!("time_start" %in% names(struct_df)) && "time_s" %in% names(struct_df)) {
  struct_df$time_start <- struct_df$time_s
}
if (!("time_end" %in% names(struct_df)) && "time_s" %in% names(struct_df)) {
  struct_df$time_end <- struct_df$time_s
}

struct_df <- ensure_cols(struct_df, expected_struct, default = NA)

struct_df <- struct_df %>%
  select(all_of(expected_struct)) %>%
  mutate(
    pointid = as.character(pointid),
    time_s  = as.numeric(time_s),
    teacher_ratio = as.numeric(teacher_ratio)
  ) %>%
  arrange(time_s)

# ----------------------------
# (4) AUTO-MATCH: point -> nearest event by peak_sec
# ----------------------------
match_one <- function(t_point, events_df) {
  if (is.na(t_point) || nrow(events_df) == 0) return(NULL)
  lag_vec <- events_df$peak_sec - t_point
  j <- which.min(abs(lag_vec))
  ev <- events_df[j, , drop = FALSE]
  lag <- as.numeric(ev$peak_sec) - t_point
  list(
    eventid = as.character(ev$eventid),
    event_peak_sec = as.numeric(ev$peak_sec),
    lag_sec = as.numeric(lag),
    event_start_sec = as.numeric(ev$start_sec),
    event_end_sec = as.numeric(ev$end_sec),
    event_max_speed = as.numeric(ev$max_speed),
    event_duration_sec = as.numeric(ev$duration_sec),
    ok = is.finite(lag) && abs(lag) <= dt_seconds
  )
}

aligned_tbl <- struct_df %>%
  rowwise() %>%
  mutate(.m = list(match_one(time_s, gesture_df))) %>%
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
# (5) SUMMARY METRICS
# ----------------------------
q_val <- NA
if ("q" %in% names(gesture_df)) {
  q_val <- suppressWarnings(as.numeric(gesture_df$q)) %>% unique()
  q_val <- q_val[!is.na(q_val)]
  if (length(q_val) > 1) q_val <- q_val[[1]]
  if (length(q_val) == 0) q_val <- NA
}

n_events <- nrow(gesture_df)

total_sec <- NA_real_
if (!is.null(wrist_df) && "t_sec" %in% names(wrist_df)) {
  total_sec <- suppressWarnings(max(as.numeric(wrist_df$t_sec), na.rm = TRUE))
}
if (!is.finite(total_sec) || is.na(total_sec)) {
  total_sec <- suppressWarnings(max(as.numeric(gesture_df$end_sec), na.rm = TRUE))
}
total_min <- as.numeric(total_sec) / 60
events_per_min <- ifelse(is.finite(total_min) && total_min > 0, n_events / total_min, NA_real_)

mean_amp <- suppressWarnings(mean(as.numeric(gesture_df$max_speed), na.rm = TRUE))
sd_amp   <- suppressWarnings(sd(as.numeric(gesture_df$max_speed), na.rm = TRUE))
mean_dur <- suppressWarnings(mean(as.numeric(gesture_df$duration_sec), na.rm = TRUE))
sd_dur   <- suppressWarnings(sd(as.numeric(gesture_df$duration_sec), na.rm = TRUE))

summary_row <- data.frame(
  TAG = TAG,
  q = q_val,
  n_events = n_events,
  events_per_min = events_per_min,
  mean_amplitude = mean_amp,
  sd_amplitude = sd_amp,
  mean_duration = mean_dur,
  sd_duration = sd_dur,
  stringsAsFactors = FALSE
)

# ----------------------------
# (5.5) COUNTSFLOW (TeacherQC)
# ----------------------------
countsflow <- list(
  total_frames = NA_real_,
  teacher_present_frames = NA_real_,
  conf_ok_frames = NA_real_,
  shoulders_ok_frames = NA_real_,
  n_candidates = NA_real_,
  n_watch_final = NA_real_
)

if (!is.null(qc_df) && nrow(qc_df) > 0) {
  
  if ("total_frames" %in% names(qc_df)) countsflow$total_frames <- num1(qc_df$total_frames)
  if ("n_candidates" %in% names(qc_df))  countsflow$n_candidates  <- num1(qc_df$n_candidates)
  if ("n_watch_final" %in% names(qc_df)) countsflow$n_watch_final <- num1(qc_df$n_watch_final)
  
  if ("teacher_present_frames" %in% names(qc_df)) countsflow$teacher_present_frames <- num1(qc_df$teacher_present_frames)
  if ("conf_ok_frames" %in% names(qc_df))         countsflow$conf_ok_frames         <- num1(qc_df$conf_ok_frames)
  if ("shoulders_ok_frames" %in% names(qc_df))    countsflow$shoulders_ok_frames    <- num1(qc_df$shoulders_ok_frames)
  
  pick_col <- function(df, candidates) {
    for (nm in candidates) if (nm %in% names(df)) return(nm)
    return(NA_character_)
  }
  
  totalF <- countsflow$total_frames
  if (is.finite(totalF) && !is.na(totalF)) {
    
    col_teacher_ratio <- pick_col(qc_df, c(
      "teacher_present_ratio","teacher_pr","teacher_pr_ratio","teacher_present_ra",
      "teacher_present_r","teacher_ratio","teacher_present"
    ))
    
    col_conf_ratio <- pick_col(qc_df, c(
      "conf_ok_ratio","conf_ok_ra","conf_ok_r","conf_ratio"
    ))
    
    col_sh_ratio <- pick_col(qc_df, c(
      "shoulders_ok_ratio","shoulders_ok_ra","shoulders_ok_r",
      "shoulders_ratio","shoulders_ok"
    ))
    
    if (is.na(countsflow$teacher_present_frames) && !is.na(col_teacher_ratio)) {
      r <- num1(qc_df[[col_teacher_ratio]])
      if (is.finite(r)) countsflow$teacher_present_frames <- round(r * totalF)
    }
    if (is.na(countsflow$conf_ok_frames) && !is.na(col_conf_ratio)) {
      r <- num1(qc_df[[col_conf_ratio]])
      if (is.finite(r)) countsflow$conf_ok_frames <- round(r * totalF)
    }
    if (is.na(countsflow$shoulders_ok_frames) && !is.na(col_sh_ratio)) {
      r <- num1(qc_df[[col_sh_ratio]])
      if (is.finite(r)) countsflow$shoulders_ok_frames <- round(r * totalF)
    }
  }
}

# Optional last-resort fallbacks from wrist_df
if (!is.null(wrist_df) && nrow(wrist_df) > 0) {
  
  if (is.na(countsflow$total_frames)) countsflow$total_frames <- nrow(wrist_df)
  
  if (is.na(countsflow$teacher_present_frames) && "teacher_present" %in% names(wrist_df)) {
    countsflow$teacher_present_frames <- sum(as.logical(wrist_df$teacher_present), na.rm = TRUE)
  }
  if (is.na(countsflow$conf_ok_frames) && "conf_ok" %in% names(wrist_df)) {
    countsflow$conf_ok_frames <- sum(as.logical(wrist_df$conf_ok), na.rm = TRUE)
  }
  if (is.na(countsflow$shoulders_ok_frames) && "shoulders_ok" %in% names(wrist_df)) {
    countsflow$shoulders_ok_frames <- sum(as.logical(wrist_df$shoulders_ok), na.rm = TRUE)
  }
}

# ----------------------------
# (6) LOAD TEMPLATE + WRITE
# ----------------------------
wb <- openxlsx::loadWorkbook(template_xlsx)

# Inputs
if ("Inputs" %in% names(wb)) {
  openxlsx::writeData(wb, "Inputs", x = TAG,        startCol = 2, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Inputs", x = fps,        startCol = 2, startRow = 3, colNames = FALSE)
  openxlsx::writeData(wb, "Inputs", x = dt_seconds, startCol = 2, startRow = 4, colNames = FALSE)
}

# Data sheets
write_df_under_header(wb, "GestureEvents",    gesture_df, max_rows_to_clear = 3000,  max_cols_to_clear = 30)
write_df_under_header(wb, "StructuralPoints", struct_df,  max_rows_to_clear = 2000,  max_cols_to_clear = 30)
write_df_under_header(wb, "WristTracks",      wrist_df,   max_rows_to_clear = 50000, max_cols_to_clear = 60)
write_df_under_header(wb, "DebugMissing",     debug_df,   max_rows_to_clear = 5000,  max_cols_to_clear = 30)

# ----------------------------
# (7) Alignment: fill key input columns (AUTO)
# ----------------------------
if ("Alignment" %in% names(wb)) {
  
  n_points  <- nrow(aligned_tbl)
  start_row <- 2
  max_clear <- max(500, n_points + 50)
  
  cols_to_clear <- c(1,2,4,5,6,7,8,9,10,12,13,14,15)  # A,B,D,E,F,G,H,I,J,L,M,N,O
  for (cc in cols_to_clear) {
    openxlsx::writeData(
      wb, "Alignment",
      x = rep("", max_clear),
      startCol = cc, startRow = start_row,
      colNames = FALSE
    )
  }
  
  if (n_points > 0) {
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$pointid,
                        startCol = 1, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$time_s,
                        startCol = 2, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$EventID_auto,
                        startCol = 4, startRow = start_row, colNames = FALSE)
    
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_peak_sec,
                        startCol = 5, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$lag_sec,
                        startCol = 6, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_start_sec,
                        startCol = 7, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_end_sec,
                        startCol = 8, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_max_speed,
                        startCol = 9, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$event_duration_sec,
                        startCol = 10, startRow = start_row, colNames = FALSE)
    
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$teacher_ratio,
                        startCol = 12, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$macro,
                        startCol = 13, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$micro_codes,
                        startCol = 14, startRow = start_row, colNames = FALSE)
    openxlsx::writeData(wb, "Alignment", x = aligned_tbl$confidence_1to3,
                        startCol = 15, startRow = start_row, colNames = FALSE)
  }
}

# Summary
if ("Summary" %in% names(wb)) {
  openxlsx::writeData(wb, "Summary", x = rep("", 40),
                      startCol = 1, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Summary", x = summary_row,
                      startCol = 1, startRow = 2, colNames = TRUE)
}

# Robustness
if ("Robustness" %in% names(wb)) {
  rob_rows <- 2:(1 + length(dt_grid))
  for (i in seq_along(dt_grid)) {
    r <- rob_rows[i]
    openxlsx::writeData(wb, "Robustness", x = TAG,            startCol = 1, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = q_val,          startCol = 2, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = n_events,       startCol = 3, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = events_per_min, startCol = 4, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = mean_amp,       startCol = 5, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = mean_dur,       startCol = 6, startRow = r, colNames = FALSE)
    openxlsx::writeData(wb, "Robustness", x = dt_grid[i],     startCol = 8, startRow = r, colNames = FALSE)
  }
}

# Methods
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
# (8) SAVE
# ----------------------------
dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)
openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)

message("DONE -> ", normalizePath(out_xlsx, winslash = "/", mustWork = FALSE))
message(sprintf("AUTO-MATCH: %d points, matched %d within dt=%.3fs",
                nrow(aligned_tbl),
                sum(!is.na(aligned_tbl$EventID_auto)),
                dt_seconds))
