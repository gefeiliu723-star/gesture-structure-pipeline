## ============================================================
# build_alignment_from_template_with_formulas.R  (FULL VERSION)
# Note:
#  - Goal: Use alignment_template_with_formulas.xlsx (keep formulas/formatting) to build <TAG>_alignment.xlsx
#  - Principle: Do NOT overwrite formula cells in Alignment; only write data blocks.
#  - Robust:
#      * flexible file name matching (TAG suffix OK)
#      * schema normalization + auto-fill missing columns (NA or derived)
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
TAG  <- ifelse(length(args) >= 1, args[[1]], "m01")

fps        <- 30
dt_seconds <- 1.5
dt_grid    <- c(0.5, 1.0, 1.5)

# project root = repo root (assumes script is under scripts/)
get_script_dir <- function() {
  # 1) Rscript --file=...
  cmd <- commandArgs(trailingOnly = FALSE)
  file_arg <- grep("^--file=", cmd, value = TRUE)
  if (length(file_arg) == 1) return(dirname(normalizePath(sub("^--file=", "", file_arg))))
  
  # 2) RStudio Source
  if (!is.null(sys.frame(1)$ofile)) return(dirname(normalizePath(sys.frame(1)$ofile)))
  
  # 3) fallback
  getwd()
}

script_dir   <- get_script_dir()
project_root <- normalizePath(file.path(script_dir, ".."), winslash = "/", mustWork = FALSE)

exports_root <- file.path(project_root, "exports")
tag_dir      <- file.path(exports_root, TAG)

template_xlsx <- file.path(exports_root, "alignment_template_with_formulas.xlsx")
out_xlsx      <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

# ----------------------------
# (1) HELPERS
# ----------------------------

stop_if_missing <- function(path, what = "file") {
  if (!file.exists(path)) stop(sprintf("Missing %s: %s", what, path), call. = FALSE)
}

std_names <- function(x) {
  x %>%
    str_replace_all("\\.+", "_") %>%
    str_replace_all("[^A-Za-z0-9_]", "_") %>%
    str_replace_all("_+", "_") %>%
    tolower()
}

# Find one file by flexible regex patterns
find_one <- function(dir, patterns, optional = FALSE) {
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

safe_read_csv <- function(path) {
  if (is.na(path) || !nzchar(path) || !file.exists(path)) return(NULL)
  readr::read_csv(path, show_col_types = FALSE, progress = FALSE)
}

# Clear & write df under existing header row (row 1)
write_df_under_header <- function(wb, sheet, df, max_rows_to_clear = 2000, max_cols_to_clear = 60) {
  if (is.null(df)) return(invisible(FALSE))
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

# Rename columns using a named character vector: c("old"="new")
apply_rename_map <- function(df, rename_map) {
  for (old in names(rename_map)) {
    new <- rename_map[[old]]
    if (old %in% names(df) && !(new %in% names(df))) {
      names(df)[names(df) == old] <- new
    }
  }
  df
}

# Ensure columns exist; missing columns will be created with NA (or provided default)
ensure_cols <- function(df, cols, default = NA) {
  for (cc in cols) {
    if (!(cc %in% names(df))) df[[cc]] <- default
  }
  df
}

# ----------------------------
# (2) LOCATE INPUT FILES (flexible names)
# ----------------------------

if (!dir.exists(tag_dir)) {
  message("❌ Tag folder not found: ", tag_dir)
  message("Create it and put required CSVs inside, e.g.:")
  message("  exports/", TAG, "/gesture_events_", TAG, ".csv")
  message("  exports/", TAG, "/watch_points_", TAG, "_teacher.csv")
  message("")
  message("Optional:")
  message("  exports/", TAG, "/wrist_tracks_", TAG, ".csv")
  message("  exports/", TAG, "/teacher_qc_", TAG, ".csv")
  quit(status = 1)
}

if (!file.exists(template_xlsx)) {
  message("❌ Template not found: ", template_xlsx)
  message("Expected location (in repo): exports/alignment_template_with_formulas.xlsx")
  quit(status = 1)
}

gesture_csv <- find_one(
  tag_dir,
  patterns = c(
    "gesture.*events.*\\.csv$",
    "gesture_events.*\\.csv$",
    "gestureevents.*\\.csv$"
  )
)

# Your "StructuralPoints" is actually watch_points_*.csv
struct_csv <- find_one(
  tag_dir,
  patterns = c(
    "watch[_-]?points.*\\.csv$",
    "structural[_-]?points.*\\.csv$",
    "structure.*points.*\\.csv$",
    "discourse.*points.*\\.csv$"
  )
)

# Optional: wrist tracks (support tag suffix)
wrist_csv <- find_one(
  tag_dir,
  patterns = c(
    "wrist[_-]?tracks.*\\.csv$",
    "wristtracks.*\\.csv$",
    "tracks.*wrist.*\\.csv$"
  ),
  optional = TRUE
)

# Optional: teacher QC (support teacher_qc_m01.csv)
teacher_qc_csv <- find_one(
  tag_dir,
  patterns = c(
    "teacher[_-]?qc.*\\.csv$",
    "teacherqc.*\\.csv$",
    "qc.*teacher.*\\.csv$"
  ),
  optional = TRUE
)

message("Using files:")
message("  GestureEvents : ", gesture_csv)
message("  StructuralPts : ", struct_csv)
message("  WristTracks   : ", wrist_csv)
message("  TeacherQC     : ", teacher_qc_csv)

# ----------------------------
# (3) READ + NORMALIZE SCHEMAS
# ----------------------------
gesture_df <- safe_read_csv(gesture_csv)
struct_df  <- safe_read_csv(struct_csv)
wrist_df   <- safe_read_csv(wrist_csv)
qc_df      <- safe_read_csv(teacher_qc_csv)

if (is.null(gesture_df) || is.null(struct_df)) {
  stop("GestureEvents and StructuralPoints CSV are required.", call. = FALSE)
}

names(gesture_df) <- std_names(names(gesture_df))
names(struct_df)  <- std_names(names(struct_df))
if (!is.null(wrist_df)) names(wrist_df) <- std_names(names(wrist_df))
if (!is.null(qc_df))    names(qc_df)    <- std_names(names(qc_df))

# ----------------------------
# (3A) GestureEvents: robust mapping + derive missing
# ----------------------------
# Target (template) columns:
expected_gesture <- c(
  "eventid","run","start_frame","end_frame","duration_frames",
  "start_sec","end_sec","duration_sec",
  "mean_speed","max_speed","q","threshold","peak_sec_proxy"
)

rename_map_gesture <- c(
  "event_id"  = "eventid",
  "id"        = "eventid",
  "trial"     = "run",
  "segment"   = "run",
  "start"     = "start_sec",
  "end"       = "end_sec",
  "dur"       = "duration_sec",
  "duration"  = "duration_sec",
  "max_vel"   = "max_speed",
  "mean_vel"  = "mean_speed",
  "quantile"  = "q",
  "peak_sec"  = "peak_sec_proxy",
  "peak_t"    = "peak_sec_proxy"
)
gesture_df <- apply_rename_map(gesture_df, rename_map_gesture)

# If eventid missing -> create deterministic IDs
if (!("eventid" %in% names(gesture_df))) {
  gesture_df$eventid <- sprintf("E%04d", seq_len(nrow(gesture_df)))
}

# Derive seconds/frames if possible
# start_sec/end_sec
if (!("start_sec" %in% names(gesture_df)) && "start_frame" %in% names(gesture_df)) {
  gesture_df$start_sec <- as.numeric(gesture_df$start_frame) / fps
}
if (!("end_sec" %in% names(gesture_df)) && "end_frame" %in% names(gesture_df)) {
  gesture_df$end_sec <- as.numeric(gesture_df$end_frame) / fps
}

# duration_sec
if (!("duration_sec" %in% names(gesture_df)) && all(c("start_sec","end_sec") %in% names(gesture_df))) {
  gesture_df$duration_sec <- as.numeric(gesture_df$end_sec) - as.numeric(gesture_df$start_sec)
}
# duration_frames
if (!("duration_frames" %in% names(gesture_df)) && all(c("start_frame","end_frame") %in% names(gesture_df))) {
  gesture_df$duration_frames <- as.numeric(gesture_df$end_frame) - as.numeric(gesture_df$start_frame)
}

# peak_sec_proxy (if missing -> use midpoint)
if (!("peak_sec_proxy" %in% names(gesture_df))) {
  if (all(c("start_sec","end_sec") %in% names(gesture_df))) {
    gesture_df$peak_sec_proxy <- (as.numeric(gesture_df$start_sec) + as.numeric(gesture_df$end_sec)) / 2
  } else if (all(c("start_frame","end_frame") %in% names(gesture_df))) {
    gesture_df$peak_sec_proxy <- (as.numeric(gesture_df$start_frame) + as.numeric(gesture_df$end_frame)) / (2 * fps)
  } else {
    gesture_df$peak_sec_proxy <- NA_real_
  }
}

# Ensure ALL expected columns exist (fill missing with NA)
gesture_df <- ensure_cols(gesture_df, expected_gesture, default = NA)

# Select + de-duplicate
gesture_df <- gesture_df %>%
  select(all_of(expected_gesture)) %>%
  mutate(eventid = as.character(eventid)) %>%
  distinct(eventid, .keep_all = TRUE) %>%
  arrange(suppressWarnings(as.numeric(str_extract(eventid, "\\d+"))), eventid)

# ----------------------------
# (3B) StructuralPoints (watch_points): robust mapping + fill missing
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

# If pointid missing -> create IDs
if (!("pointid" %in% names(struct_df))) {
  struct_df$pointid <- sprintf("P%03d", seq_len(nrow(struct_df)))
}

# If time_start/time_end missing but time_s exists -> create window endpoints as same instant
if (!("time_start" %in% names(struct_df)) && "time_s" %in% names(struct_df)) {
  struct_df$time_start <- struct_df$time_s
}
if (!("time_end" %in% names(struct_df)) && "time_s" %in% names(struct_df)) {
  struct_df$time_end <- struct_df$time_s
}

# Ensure all expected cols exist
struct_df <- ensure_cols(struct_df, expected_struct, default = NA)

struct_df <- struct_df %>%
  select(all_of(expected_struct)) %>%
  mutate(pointid = as.character(pointid)) %>%
  arrange(as.numeric(time_s))

# ----------------------------
# (3C) Wrist tracks (optional) - keep as is, but reorder if matches
# ----------------------------
expected_wrist <- c(
  "frame","t_sec","teacher_id","conf_det","x1","y1","x2","y2",
  "ls_x","ls_y","ls_z","ls_vis",
  "rs_x","rs_y","rs_z","rs_vis",
  "le_x","le_y","le_z","le_vis",
  "re_x","re_y","re_z","re_vis",
  "lw_x","lw_y","lw_z","lw_vis",
  "rw_x","rw_y","rw_z","rw_vis"
)
if (!is.null(wrist_df)) {
  if (all(expected_wrist %in% names(wrist_df))) {
    wrist_df <- wrist_df %>% select(all_of(expected_wrist))
  }
}

# ----------------------------
# (4) SUMMARY METRICS
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

# Optional CountsFlow from teacher QC
countsflow <- list(
  total_frames = NA,
  teacher_present_frames = NA,
  conf_ok_frames = NA,
  shoulders_ok_frames = NA,
  n_candidates = NA,
  n_watch_final = NA
)

if (!is.null(qc_df)) {
  for (nm in names(countsflow)) {
    if (nm %in% names(qc_df)) countsflow[[nm]] <- qc_df[[nm]][[1]]
  }
}

# ----------------------------
# (5) LOAD TEMPLATE + WRITE
# ----------------------------
wb <- openxlsx::loadWorkbook(template_xlsx)

# Inputs
openxlsx::writeData(wb, "Inputs", x = TAG,        startCol = 2, startRow = 2, colNames = FALSE)
openxlsx::writeData(wb, "Inputs", x = fps,        startCol = 2, startRow = 3, colNames = FALSE)
openxlsx::writeData(wb, "Inputs", x = dt_seconds, startCol = 2, startRow = 4, colNames = FALSE)

# Sheets
write_df_under_header(wb, "GestureEvents",    gesture_df, max_rows_to_clear = 3000,  max_cols_to_clear = 30)
write_df_under_header(wb, "StructuralPoints", struct_df,  max_rows_to_clear = 2000,  max_cols_to_clear = 30)
if (!is.null(wrist_df)) {
  write_df_under_header(wb, "WristTracks", wrist_df, max_rows_to_clear = 50000, max_cols_to_clear = 40)
}

# Alignment prefill (only typed columns; keep formulas intact)
n_points  <- nrow(struct_df)
start_row <- 2
max_clear <- 500

# Clear A,B,D,L,P only
openxlsx::writeData(wb, "Alignment", x = rep("", max_clear), startCol = 1,  startRow = start_row, colNames = FALSE) # A
openxlsx::writeData(wb, "Alignment", x = rep("", max_clear), startCol = 2,  startRow = start_row, colNames = FALSE) # B
openxlsx::writeData(wb, "Alignment", x = rep("", max_clear), startCol = 4,  startRow = start_row, colNames = FALSE) # D
openxlsx::writeData(wb, "Alignment", x = rep("", max_clear), startCol = 12, startRow = start_row, colNames = FALSE) # L
openxlsx::writeData(wb, "Alignment", x = rep("", max_clear), startCol = 16, startRow = start_row, colNames = FALSE) # P

if (n_points > 0) {
  openxlsx::writeData(wb, "Alignment", x = struct_df$pointid,       startCol = 1,  startRow = start_row, colNames = FALSE)
  openxlsx::writeData(wb, "Alignment", x = struct_df$time_s,        startCol = 2,  startRow = start_row, colNames = FALSE)
  openxlsx::writeData(wb, "Alignment", x = struct_df$teacher_ratio, startCol = 12, startRow = start_row, colNames = FALSE)
}

# Summary (row 2)
openxlsx::writeData(wb, "Summary", x = rep("", 40), startCol = 1, startRow = 2, colNames = FALSE)
openxlsx::writeData(wb, "Summary", x = summary_row, startCol = 1, startRow = 2, colNames = TRUE)

# Robustness
rob_rows <- 2:(1 + length(dt_grid))
for (i in seq_along(dt_grid)) {
  r <- rob_rows[i]
  openxlsx::writeData(wb, "Robustness", x = TAG,            startCol = 1, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = q_val,          startCol = 2, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = n_events,       startCol = 3, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = events_per_min, startCol = 4, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = mean_amp,       startCol = 5, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = mean_dur,       startCol = 6, startRow = r, colNames = FALSE)
  openxlsx::writeData(wb, "Robustness", x = dt_grid[i],     startCol = 8, startRow = r, colNames = FALSE) # col H
}

# Methods values (col C)
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

# ----------------------------
# (6) SAVE
# ----------------------------
if (!dir.exists(tag_dir)) dir.create(tag_dir, recursive = TRUE)
openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)

message("DONE -> ", out_xlsx)

# EN:
#  - Change TAG / paths / fps / dt_seconds only.
#  - Required: gesture_events*.csv + watch_points*.csv
## ============================================================
