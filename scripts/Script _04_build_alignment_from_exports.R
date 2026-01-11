## ============================================================
# build_alignment_from_template_with_formulas.R  (GITHUB FRIENDLY / NO-DATA)
# CN:
#  - Repo-friendly: NO hard-coded absolute paths (default repo root)
#  - Usage:
#       Rscript scripts/build_alignment_from_template_with_formulas.R m01
#    (TAG optional; default m01)
#  - Inputs expected (local, NOT committed):
#       exports/<TAG>/gesture_events*.csv
#       exports/<TAG>/watch_points*.csv
#    Optional:
#       exports/<TAG>/teacher_qc*.csv
#       exports/<TAG>/wrist_tracks*.csv   (or exports/wrist_tracks/*TAG*.csv)
#       exports/<TAG>/debug_missing_reason*.csv (or exports/*)
#  - Template (committed):
#       exports/alignment_template_with_formulas.xlsx
#
# EN:
#  - Goal: Use alignment_template_with_formulas.xlsx (keep formulas/formatting)
#          to build exports/<TAG>/<TAG>_alignment.xlsx
#  - Principle: Do NOT overwrite formula cells in Alignment; only write data blocks.
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(readr)
  library(dplyr)
  library(stringr)
})

# ----------------------------
# (0) USER SETTINGS
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)
TAG  <- ifelse(length(args) >= 1, args[[1]], "m01")

fps        <- 30
dt_seconds <- 1.5
dt_grid    <- c(0.5, 1.0, 1.5)

# repo root = parent of scripts/ (robust for Rscript/RStudio)
get_script_dir <- function() {
  cmd <- commandArgs(trailingOnly = FALSE)
  file_arg <- grep("^--file=", cmd, value = TRUE)
  if (length(file_arg) == 1) return(dirname(normalizePath(sub("^--file=", "", file_arg), winslash = "/")))
  if (!is.null(sys.frame(1)$ofile)) return(dirname(normalizePath(sys.frame(1)$ofile, winslash = "/")))
  getwd()
}

script_dir   <- get_script_dir()
project_root <- normalizePath(file.path(script_dir, ".."), winslash = "/", mustWork = FALSE)

exports_root  <- file.path(project_root, "exports")
tag_dir       <- file.path(exports_root, TAG)

template_xlsx <- file.path(exports_root, "alignment_template_with_formulas.xlsx")
out_xlsx      <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

# ----------------------------
# (1) HELPERS
# ----------------------------
std_names <- function(x) {
  x %>%
    str_replace_all("\\.+", "_") %>%
    str_replace_all("[^A-Za-z0-9_]", "_") %>%
    str_replace_all("_+", "_") %>%
    tolower()
}

# find one file in ONE dir
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

# find one file across MULTIPLE dirs
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

# write df under header row (row 1); only if sheet exists
write_df_under_header <- function(wb, sheet, df, max_rows_to_clear = 2000, max_cols_to_clear = 60) {
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
if (!dir.exists(tag_dir)) {
  message("❌ Tag folder not found: ", tag_dir)
  message("Create it and put required CSVs inside, e.g.:")
  message("  exports/", TAG, "/gesture_events_", TAG, ".csv")
  message("  exports/", TAG, "/watch_points_", TAG, "_teacher.csv")
  quit(status = 1)
}
if (!file.exists(template_xlsx)) {
  message("❌ Template not found: ", template_xlsx)
  message("Expected committed file: exports/alignment_template_with_formulas.xlsx")
  quit(status = 1)
}

gesture_csv <- find_one_in_dir(
  tag_dir,
  patterns = c("gesture.*events.*\\.csv$", "gesture_events.*\\.csv$", "gestureevents.*\\.csv$")
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

# WristTracks: tag_dir -> exports/wrist_tracks -> exports_root
wrist_candidates <- c(tag_dir, file.path(exports_root, "wrist_tracks"), exports_root)
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
  patterns = c("teacher[_-]?qc.*\\.csv$", "teacherqc.*\\.csv$", "qc.*teacher.*\\.csv$"),
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
# (3A) GestureEvents schema normalize
# ----------------------------
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

if (!("eventid" %in% names(gesture_df))) gesture_df$eventid <- sprintf("E%04d", seq_len(nrow(gesture_df)))

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
if (!("peak_sec_proxy" %in% names(gesture_df))) {
  if (all(c("start_sec","end_sec") %in% names(gesture_df))) {
    gesture_df$peak_sec_proxy <- (as.numeric(gesture_df$start_sec) + as.numeric(gesture_df$end_sec)) / 2
  } else if (all(c("start_frame","end_frame") %in% names(gesture_df))) {
    gesture_df$peak_sec_proxy <- (as.numeric(gesture_df$start_frame) + as.numeric(gesture_df$end_frame)) / (2 * fps)
  } else {
    gesture_df$peak_sec_proxy <- NA_real_
  }
}

gesture_df <- ensure_cols(gesture_df, expected_gesture, default = NA)

gesture_df <- gesture_df %>%
  select(all_of(expected_gesture)) %>%
  mutate(eventid = as.character(eventid)) %>%
  distinct(eventid, .keep_all = TRUE) %>%
  arrange(suppressWarnings(as.numeric(str_extract(eventid, "\\d+"))), eventid)

# ----------------------------
# (3B) StructuralPoints schema normalize
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

if (!("pointid" %in% names(struct_df))) struct_df$pointid <- sprintf("P%03d", seq_len(nrow(struct_df)))
if (!("time_start" %in% names(struct_df)) && "time_s" %in% names(struct_df)) struct_df$time_start <- struct_df$time_s
if (!("time_end" %in% names(struct_df)) && "time_s" %in% names(struct_df))   struct_df$time_end   <- struct_df$time_s

struct_df <- ensure_cols(struct_df, expected_struct, default = NA)

struct_df <- struct_df %>%
  select(all_of(expected_struct)) %>%
  mutate(pointid = as.character(pointid)) %>%
  arrange(as.numeric(time_s))

# ----------------------------
# (3C) WristTracks (optional) reorder if matches template
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
  if (all(expected_wrist %in% names(wrist_df))) wrist_df <- wrist_df %>% select(all_of(expected_wrist))
}

# ----------------------------
# (4) SUMMARY METRICS
# ----------------------------
q_val <- NA_real_
if ("q" %in% names(gesture_df)) {
  q_val <- suppressWarnings(as.numeric(gesture_df$q)) %>% unique()
  q_val <- q_val[!is.na(q_val)]
  if (length(q_val) > 1) q_val <- q_val[[1]]
  if (length(q_val) == 0) q_val <- NA_real_
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
# (4B) COUNTSFLOW (prefer TeacherQC; fallback to WristTracks)
# ----------------------------
countsflow <- list(
  total_frames = NA_real_,
  teacher_present_frames = NA_real_,
  conf_ok_frames = NA_real_,
  shoulders_ok_frames = NA_real_,
  n_candidates = NA_real_,
  n_watch_final = NA_real_
)

# 1) TeacherQC preferred (if has real values)
if (!is.null(qc_df)) {
  if ("total_frames" %in% names(qc_df))            countsflow$total_frames           <- num1(qc_df$total_frames)
  if ("teacher_present_frames" %in% names(qc_df))  countsflow$teacher_present_frames <- num1(qc_df$teacher_present_frames)
  if ("conf_ok_frames" %in% names(qc_df))          countsflow$conf_ok_frames         <- num1(qc_df$conf_ok_frames)
  if ("shoulders_ok_frames" %in% names(qc_df))     countsflow$shoulders_ok_frames    <- num1(qc_df$shoulders_ok_frames)
  if ("n_candidates" %in% names(qc_df))            countsflow$n_candidates           <- num1(qc_df$n_candidates)
  if ("n_watch_final" %in% names(qc_df))           countsflow$n_watch_final          <- num1(qc_df$n_watch_final)
}

# 2) Fallback from WristTracks (robust against missing TeacherQC fields)
if (!is.null(wrist_df)) {
  
  # total frames
  if (!is.finite(countsflow$total_frames) || is.na(countsflow$total_frames)) {
    if ("frame" %in% names(wrist_df)) {
      countsflow$total_frames <- suppressWarnings(max(as.numeric(wrist_df$frame), na.rm = TRUE)) + 1
      if (!is.finite(countsflow$total_frames)) countsflow$total_frames <- nrow(wrist_df)
    } else {
      countsflow$total_frames <- nrow(wrist_df)
    }
  }
  
  # teacher present frames = has conf_det (+ bbox if available)
  if (!is.finite(countsflow$teacher_present_frames) || is.na(countsflow$teacher_present_frames)) {
    if ("conf_det" %in% names(wrist_df)) {
      conf <- suppressWarnings(as.numeric(wrist_df$conf_det))
      present <- is.finite(conf)
      
      has_bbox <- all(c("x1","y1","x2","y2") %in% names(wrist_df))
      if (has_bbox) {
        x1 <- suppressWarnings(as.numeric(wrist_df$x1))
        y1 <- suppressWarnings(as.numeric(wrist_df$y1))
        x2 <- suppressWarnings(as.numeric(wrist_df$x2))
        y2 <- suppressWarnings(as.numeric(wrist_df$y2))
        present <- present & is.finite(x1) & is.finite(y1) & is.finite(x2) & is.finite(y2)
      }
      
      countsflow$teacher_present_frames <- sum(present, na.rm = TRUE)
    }
  }
  
  # conf ok frames (threshold adjustable)
  if (!is.finite(countsflow$conf_ok_frames) || is.na(countsflow$conf_ok_frames)) {
    if ("conf_det" %in% names(wrist_df)) {
      conf <- suppressWarnings(as.numeric(wrist_df$conf_det))
      countsflow$conf_ok_frames <- sum(is.finite(conf) & conf >= 0.50, na.rm = TRUE)
    }
  }
  
  # shoulders ok frames (visibility threshold adjustable)
  if (!is.finite(countsflow$shoulders_ok_frames) || is.na(countsflow$shoulders_ok_frames)) {
    if (all(c("ls_vis","rs_vis") %in% names(wrist_df))) {
      ls <- suppressWarnings(as.numeric(wrist_df$ls_vis))
      rs <- suppressWarnings(as.numeric(wrist_df$rs_vis))
      countsflow$shoulders_ok_frames <- sum(is.finite(ls) & is.finite(rs) & ls >= 0.50 & rs >= 0.50, na.rm = TRUE)
    }
  }
}

# ----------------------------
# (5) LOAD TEMPLATE + WRITE
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
write_df_under_header(wb, "DebugMissing",     debug_df,   max_rows_to_clear = 5000,  max_cols_to_clear = 60)

# Alignment prefill (only typed columns; keep formulas intact)
if ("Alignment" %in% names(wb)) {
  n_points  <- nrow(struct_df)
  start_row <- 2
  max_clear <- 500
  
  # Clear typed columns only (A,B,D,L,P) so formulas elsewhere survive
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
}

# Summary
if ("Summary" %in% names(wb)) {
  openxlsx::writeData(wb, "Summary", x = rep("", 40), startCol = 1, startRow = 2, colNames = FALSE)
  openxlsx::writeData(wb, "Summary", x = summary_row, startCol = 1, startRow = 2, colNames = TRUE)
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
    openxlsx::writeData(wb, "Robustness", x = dt_grid[i],     startCol = 8, startRow = r, colNames = FALSE) # H
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
# (6) SAVE
# ----------------------------
if (!dir.exists(tag_dir)) dir.create(tag_dir, recursive = TRUE)
openxlsx::saveWorkbook(wb, out_xlsx, overwrite = TRUE)

message("DONE -> ", out_xlsx)
