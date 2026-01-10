# 04) build_alignment_from_exports_by_tag_FINAL_BLUEHEADERS_NO_COUNTSFLOW.R
# ============================================================
# CN:
#  - Template: exports/alignment_template_with_formulas.xlsx (keep format/formulas)
#  - Import CSVs under exports/<TAG>/ into sheets
#  - De-duplicate GestureEvents
#  - Auto-fill Inputs: TAG / fps / Δt_seconds (write from row 2)
#  - CountsFlow: DO NOT keep standalone sheet; compute counts_df and export to Methods
#  - Robustness: write dt_grid to the right side (from H1)
#  - Methods: export Inputs + CountsFlow summary
#  - ALL HEADERS: blue background + white bold text
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
})

# ----------------------------
# (0) USER SETTINGS (PUBLIC TEMPLATE)
# ----------------------------
# IMPORTANT:
# - Do NOT commit personal local paths.
# - Configure via environment variables OR replace placeholders below.

get_env_or <- function(key, default) {
  v <- Sys.getenv(key, unset = "")
  if (nzchar(v)) v else default
}

# ---- identifiers (safe to commit) ----
TAG <- get_env_or("GESTURE_TAG", "mXX")

# ---- parameters (safe to commit) ----
fps        <- 30
dt_seconds <- 1.5
dt_grid    <- c(0.5, 1.0, 1.5)

# ---- roots ----
# Option A (recommended): env vars
#   GESTURE_ROOT         : project root (contains exports/)
#   GESTURE_EXPORTS_ROOT : exports root (if you want to point directly to exports/)
project_root <- get_env_or("GESTURE_ROOT", "<PATH_TO_PROJECT_ROOT>")
exports_root <- get_env_or("GESTURE_EXPORTS_ROOT", file.path(project_root, "exports"))

# ---- template + output ----
template_xlsx <- file.path(exports_root, "alignment_template.xlsx")
tag_dir       <- file.path(exports_root, TAG)
out_xlsx      <- file.path(tag_dir, paste0(TAG, "_alignment.xlsx"))

# ---- wrist tracks candidates (keep your original fallback logic) ----
wrist_candidates <- c(
  file.path(exports_root, "wrist_tracks", paste0("wrist_tracks_", TAG, ".csv")),
  file.path(tag_dir, paste0("wrist_tracks_", TAG, ".csv"))
)

# ----------------------------
# (1) Helper functions + Styles
# ----------------------------
stop_if_missing <- function(paths_named) {
  paths <- unlist(paths_named, use.names = TRUE)
  ok <- file.exists(paths)
  if (!all(ok)) {
    cat("Missing files:\n")
    for (k in names(paths)[!ok]) cat(" - ", k, ": ", paths[[k]], "\n", sep = "")
    stop("Some required files are missing. Generate them first.")
  }
}

read_csv_safe <- function(path) {
  read.csv(path, stringsAsFactors = FALSE, check.names = FALSE)
}

# ---- Header style (blue background, white bold text) ----
header_style_blue <- createStyle(
  fontColour     = "#FFFFFF",
  fgFill         = "#1F4E79",
  halign         = "center",
  valign         = "center",
  textDecoration = "bold",
  border         = "Bottom"
)

# Apply style to header row (row = 1)
style_header_blue <- function(wb, sheet, ncol, start_col = 1, header_row = 1) {
  if (is.null(ncol) || is.na(ncol) || ncol < 1) return(invisible(NULL))
  addStyle(
    wb,
    sheet,
    style = header_style_blue,
    rows  = header_row,
    cols  = start_col:(start_col + ncol - 1),
    gridExpand = TRUE,
    stack = TRUE
  )
}

# Write df and ALWAYS blue-style the header row (row 1)
write_df <- function(wb, sheet, df) {
  if (!(sheet %in% names(wb))) addWorksheet(wb, sheet)
  writeData(wb, sheet, df, colNames = TRUE)
  style_header_blue(wb, sheet, ncol(df), start_col = 1, header_row = 1)
  # optional QoL:
  setColWidths(wb, sheet, cols = 1:ncol(df), widths = "auto")
}

get1 <- function(df, col) {
  if (is.null(df) || !is.data.frame(df) || nrow(df) < 1) return(NA)
  if (!(col %in% names(df))) return(NA)
  df[[col]][1]
}

# ----------------------------
# (1.1) GestureEvents 去重
# ----------------------------
dedupe_gesture_events <- function(df) {
  n0 <- nrow(df)
  df <- unique(df)
  
  if ("EventID" %in% names(df)) {
    if ("max_speed" %in% names(df)) {
      df <- df[order(df$EventID, -df$max_speed), ]
    } else {
      df <- df[order(df$EventID), ]
    }
    df <- df[!duplicated(df$EventID), ]
  }
  
  message(sprintf("GestureEvents de-dup: %d → %d rows", n0, nrow(df)))
  df
}

# ----------------------------
# (2) Input CSV paths
# ----------------------------
in_files <- list(
  gesture_events = file.path(tag_dir, paste0("gesture_events_", TAG, ".csv")),
  watch_points   = file.path(tag_dir, paste0("watch_points_", TAG, "_teacher.csv")),
  teacher_qc     = file.path(tag_dir, paste0("teacher_qc_", TAG, ".csv")),
  summary        = file.path(tag_dir, paste0("summary_", TAG, ".csv")),
  robust         = file.path(tag_dir, paste0("robust_", TAG, ".csv")),
  debug_missing  = file.path(tag_dir, paste0("debug_missing_reason_", TAG, "_q0.94.csv"))
)

wrist_path <- wrist_candidates[file.exists(wrist_candidates)][1]
if (is.na(wrist_path) || !nzchar(wrist_path)) wrist_path <- wrist_candidates[1]
in_files$wrist_tracks <- wrist_path

# ----------------------------
# (3) Validate
# ----------------------------
stop_if_missing(c(template_xlsx = template_xlsx))
dir.create(tag_dir, showWarnings = FALSE, recursive = TRUE)
stop_if_missing(in_files)

# ----------------------------
# (4) Load template
# ----------------------------
wb <- loadWorkbook(template_xlsx)

# ----------------------------
# (4.1) Auto-fill Inputs sheet (write from row 2, keep template header row)
# ----------------------------
if ("Inputs" %in% names(wb)) {
  inputs_df <- data.frame(
    Parameter = c("TAG", "fps", "Δt_seconds"),
    Value     = c(TAG, fps, dt_seconds),
    Note      = c(
      "video tag",
      "frames per second",
      "alignment half-window (seconds)"
    ),
    stringsAsFactors = FALSE
  )
  
  # write values starting at row 2 (do NOT overwrite header row)
  writeData(
    wb,
    sheet    = "Inputs",
    x        = inputs_df,
    startRow = 2,
    startCol = 1,
    colNames = FALSE
  )
  
  # force Inputs header (row 1, first 3 cols) to be blue
  style_header_blue(wb, "Inputs", ncol(inputs_df), start_col = 1, header_row = 1)
  setColWidths(wb, "Inputs", cols = 1:ncol(inputs_df), widths = "auto")
  
} else {
  message("NOTE: 'Inputs' sheet not found in template; skipping Inputs auto-fill.")
}

# ----------------------------
# (5) Read data
# ----------------------------
df_gesture <- read_csv_safe(in_files$gesture_events)
df_points  <- read_csv_safe(in_files$watch_points)
df_qc      <- read_csv_safe(in_files$teacher_qc)
df_sum     <- read_csv_safe(in_files$summary)
df_rob     <- read_csv_safe(in_files$robust)
df_dbg     <- read_csv_safe(in_files$debug_missing)
df_wrist   <- read_csv_safe(in_files$wrist_tracks)

df_gesture <- dedupe_gesture_events(df_gesture)

# ----------------------------
# (6) Populate data sheets (ALL headers blue)
# ----------------------------
write_df(wb, "GestureEvents",    df_gesture)
write_df(wb, "StructuralPoints", df_points)
write_df(wb, "TeacherQC",        df_qc)
write_df(wb, "Summary",          df_sum)
write_df(wb, "Robustness",       df_rob)
write_df(wb, "DebugMissing",     df_dbg)
write_df(wb, "WristTracks",      df_wrist)

# ----------------------------
# (l) CountsFlow：compute counts_df (robust to column-name variants)
# ----------------------------

# helper: try multiple possible column names
get_any1 <- function(df, cols) {
  for (cc in cols) {
    if (!is.null(df) && is.data.frame(df) && nrow(df) >= 1 && (cc %in% names(df))) {
      return(df[[cc]][1])
    }
  }
  return(NA)
}

# total_frames (may be stored as total_frames or n_frames etc.)
total_frames <- get_any1(df_qc, c("total_frames", "n_frames", "frames_total"))
if (is.na(total_frames)) total_frames <- NA

# ---- teacher-present
teacher_present_frames <- get_any1(df_qc, c("teacher_present_frames", "teacher_frames", "teacher_present_n"))
teacher_present_ratio  <- get_any1(df_qc, c("teacher_present_ratio", "teacher_ratio", "teacher_pres_ratio", "teacher_present_rati"))

# ---- conf ok
conf_ok_frames <- get_any1(df_qc, c("conf_ok_frames", "conf_frames", "conf_ok_n"))
conf_ok_ratio  <- get_any1(df_qc, c("conf_ok_ratio", "conf_ratio", "conf_ok_rati", "conf_ok_r"))

# ---- shoulders ok
shoulders_ok_frames <- get_any1(df_qc, c("shoulders_ok_frames", "shoulder_ok_frames", "shoulders_ok_n"))
shoulders_ok_ratio  <- get_any1(df_qc, c("shoulders_ok_ratio", "shoulder_ok_ratio", "shoulders_ok_rati"))

# ---- candidates / watch kept (these seem correct already)
n_candidates  <- get_any1(df_qc, c("n_candidates", "candidates", "n_candidate_frames"))
n_watch_final <- get_any1(df_qc, c("n_watch_final", "n_watch_kept", "watch_final", "n_watch"))

# If frames are missing but ratios exist, derive frames using total_frames
derive_frames <- function(frames, ratio, total_frames) {
  if (!is.na(frames)) return(frames)
  if (!is.na(ratio) && !is.na(total_frames)) return(round(as.numeric(ratio) * as.numeric(total_frames)))
  return(NA)
}

teacher_present_frames <- derive_frames(teacher_present_frames, teacher_present_ratio, total_frames)
conf_ok_frames         <- derive_frames(conf_ok_frames,         conf_ok_ratio,         total_frames)
shoulders_ok_frames    <- derive_frames(shoulders_ok_frames,    shoulders_ok_ratio,    total_frames)

# gesture events
n_events       <- nrow(df_gesture)
events_per_min <- get_any1(df_sum, c("events_per_min", "events_per_minute", "event_rate_per_min"))
mean_amplitude <- get_any1(df_sum, c("mean_amplitude", "mean_peak_speed", "mean_max_speed", "mean_speed"))

counts_df <- data.frame(
  Metric = c(
    "total_frames",
    "teacher_present_frames",
    "conf_ok_frames",
    "shoulders_ok_frames",
    "n_candidates",
    "n_watch_final",
    "n_events",
    "events_per_min",
    "mean_amplitude"
  ),
  Value = c(
    total_frames,
    teacher_present_frames,
    conf_ok_frames,
    shoulders_ok_frames,
    n_candidates,
    n_watch_final,
    n_events,
    events_per_min,
    mean_amplitude
  ),
  stringsAsFactors = FALSE
)

# ensure CountsFlow sheet is NOT in output
if ("CountsFlow" %in% names(wb)) {
  removeWorksheet(wb, "CountsFlow")
}


# ----------------------------
# (6.2) Robustness dt grid -> write to the right (H1)
# ----------------------------
{
  sheet_rb <- "Robustness"
  if (!(sheet_rb %in% names(wb))) addWorksheet(wb, sheet_rb)
  
  rb_grid <- data.frame(
    dt_seconds = dt_grid,
    Note = c(
      "narrow window (stricter alignment)",
      "medium window",
      "default / wider window"
    ),
    stringsAsFactors = FALSE
  )
  
  # write at H1 (col 8)
  writeData(wb, sheet_rb, rb_grid, startRow = 1, startCol = 8, colNames = TRUE)
  
  # style that mini-table header too (row 1, cols 8..)
  style_header_blue(wb, sheet_rb, ncol(rb_grid), start_col = 8, header_row = 1)
  setColWidths(wb, sheet_rb, cols = 8:(8 + ncol(rb_grid) - 1), widths = "auto")
}

# ----------------------------
# (6.3) Methods sheet: Inputs + CountsFlow summary (header blue)
# ----------------------------
{
  methods_sheet <- "Methods"
  if (methods_sheet %in% names(wb)) removeWorksheet(wb, methods_sheet)
  addWorksheet(wb, methods_sheet)
  
  inputs_block <- data.frame(
    Section = rep("Inputs", 3),
    Item    = c("TAG", "fps", "Δt_seconds"),
    Value   = c(TAG, fps, dt_seconds),
    stringsAsFactors = FALSE
  )
  
  cf_block <- data.frame(
    Section = rep("CountsFlow", nrow(counts_df)),
    Item    = counts_df$Metric,
    Value   = counts_df$Value,
    stringsAsFactors = FALSE
  )
  
  methods_df <- rbind(inputs_block, cf_block)
  
  writeData(wb, methods_sheet, methods_df, startRow = 1, startCol = 1, colNames = TRUE)
  style_header_blue(wb, methods_sheet, ncol(methods_df), start_col = 1, header_row = 1)
  freezePane(wb, methods_sheet, firstRow = TRUE)
  setColWidths(wb, methods_sheet, cols = 1:ncol(methods_df), widths = "auto")
}

# ----------------------------
# (7) Save
# ----------------------------
saveWorkbook(wb, out_xlsx, overwrite = TRUE)

cat(
  "✅ Alignment workbook generated\n",
  "TAG: ", TAG, "\n",
  "fps: ", fps, "\n",
  "Δt_seconds: ", dt_seconds, "\n",
  "Template: ", template_xlsx, "\n",
  "Output: ", out_xlsx, "\n",
  "Header style: BLUE (#1F4E79) across all sheets\n",
  "CountsFlow sheet: REMOVED (summary kept in Methods)\n",
  sep = ""
)
