# ============================================================
# Script 05 (PORTABLE / FINAL): Time-based enrichment around structural points
#
# NO TAG / NO absolute root.
# User provides files:
#   --align  path/to/<alignment>.xlsx   (REQUIRED)
#   --wrist  path/to/wrist_tracks_<TAG>.csv  (OPTIONAL; only needed for teacher_only)
#
# Observed:
#   - Structural points = t_struct_sec from:
#       * PRIORITY: Alignment sheet (has Confidence)
#       * fallback: StructuralPoints sheet (no Confidence -> conf3 skipped)
#   - Events = gesture event peaks from GestureEvents sheet
#       * Template note: GestureEvents$t_sec IS peak time (treated as peak_sec)
#   - Compute r_in vs r_out (OUT = rest of time)
#
# Nulls:
#   1) global:       sample n_points uniformly in [w, T_total-w]
#   2) teacher_only: sample n_points uniformly within teacher-present intervals
#                   (derived from wrist_tracks csv)
#
# Output:
#   <out_root>/<tag_or_filebase>/time_enrichment_<pointset>_<null_mode>/
#     - <prefix>_summary_overall.csv
#     - <prefix>_windows_union.csv
#     - <prefix>_events_within_flags.csv
#     - <prefix>_perm_overall.csv
#     - <prefix>_allowed_intervals.csv
#   plus MASTER summary:
#     <out_root>/<tag_or_filebase>/time_enrichment_MASTER_summary.csv
#
# CLI:
#   Rscript scripts/05_time_enrichment_portable.R --align l05_alignment.xlsx --wrist wrist_tracks_l05.csv --window 1.5 --perm 5000 --seed 1 --out results
# 
# NOTE 05-A
#   Lift is defined as the ratio of event rates inside versus outside structural windows, controlling for unequal window coverage.
# NOTE 05-B（null 的意义）
#   The teacher-only null controls for temporal availability of observable gesture, isolating alignment from mere presence.
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(stringr)
  library(readr)
  library(tibble)
})

# ----------------------------
# (0) CLI PARSE (simple)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

get_arg <- function(flag, default = NULL) {
  i <- match(flag, args)
  if (!is.na(i) && i < length(args)) return(args[[i + 1]])
  default
}

align_xlsx <- get_arg("--align", NULL)
wrist_csv  <- get_arg("--wrist", NULL)
window_s   <- as.numeric(get_arg("--window", "1.5"))
n_perm     <- as.integer(get_arg("--perm", "5000"))
seed       <- as.integer(get_arg("--seed", "1"))
out_root   <- get_arg("--out", "results")

stopifnot(!is.null(align_xlsx), nzchar(align_xlsx), file.exists(align_xlsx))
stopifnot(is.finite(window_s), window_s > 0)
stopifnot(is.finite(n_perm), n_perm >= 100)
stopifnot(is.finite(seed))

# derive a stable tag/prefix from filename (e.g., l05_alignment.xlsx -> l05)
base_name <- basename(align_xlsx)
tag_guess <- str_replace(base_name, "\\.xlsx$", "")
tag_guess <- str_replace(tag_guess, "_alignment$", "")
TAG <- tag_guess

dir.create(out_root, recursive = TRUE, showWarnings = FALSE)
results_root <- file.path(out_root, TAG)
dir.create(results_root, recursive = TRUE, showWarnings = FALSE)

message("Using align_xlsx: ", normalizePath(align_xlsx, winslash = "/"))
if (!is.null(wrist_csv)) message("Using wrist_csv : ", normalizePath(wrist_csv, winslash = "/"))
message("TAG prefix       : ", TAG)
message("window_s         : ", window_s, " | n_perm=", n_perm, " | seed=", seed)
message("out_root         : ", normalizePath(out_root, winslash = "/"))

# ----------------------------
# (1) HELPERS
# ----------------------------
clean_names <- function(nms) {
  nms %>% str_replace_all("\\s+", " ") %>% str_trim()
}

sheet_exists <- function(wb_path, sheet_name) {
  sheet_name %in% openxlsx::getSheetNames(wb_path)
}

pick_col <- function(df, candidates) {
  nms <- names(df)
  hit <- candidates[candidates %in% nms]
  if (length(hit) >= 1) return(hit[[1]])
  low <- tolower(nms)
  for (cand in candidates) {
    j <- which(low == tolower(cand))
    if (length(j) >= 1) return(nms[j[[1]]])
  }
  NA_character_
}

make_windows <- function(times, w) {
  tibble(start = times - w, end = times + w) %>%
    filter(is.finite(start), is.finite(end))
}

safe_clip_windows <- function(windows_df, t_min, t_max) {
  windows_df %>%
    mutate(start = pmax(start, t_min), end = pmin(end, t_max)) %>%
    filter(end > start)
}

union_intervals <- function(df_intervals) {
  if (nrow(df_intervals) == 0) return(df_intervals)
  x <- df_intervals %>% arrange(start, end)
  starts <- x$start
  ends   <- x$end
  
  out_s <- numeric(0); out_e <- numeric(0)
  cs <- starts[[1]]; ce <- ends[[1]]
  
  for (i in 2:length(starts)) {
    s <- starts[[i]]; e <- ends[[i]]
    if (s <= ce) {
      ce <- max(ce, e)
    } else {
      out_s <- c(out_s, cs); out_e <- c(out_e, ce)
      cs <- s; ce <- e
    }
  }
  out_s <- c(out_s, cs); out_e <- c(out_e, ce)
  tibble(start = out_s, end = out_e)
}

interval_length <- function(df_union) {
  if (nrow(df_union) == 0) return(0)
  sum(pmax(0, df_union$end - df_union$start))
}

count_in_union <- function(peaks_sorted, union_df) {
  if (length(peaks_sorted) == 0 || nrow(union_df) == 0) return(0L)
  total <- 0L
  for (i in seq_len(nrow(union_df))) {
    a <- union_df$start[[i]]
    b <- union_df$end[[i]]
    left  <- findInterval(a, peaks_sorted) + 1L
    right <- findInterval(b, peaks_sorted)
    if (right >= left) total <- total + (right - left + 1L)
  }
  total
}

flag_in_union <- function(peaks, union_df) {
  if (length(peaks) == 0) return(logical(0))
  if (nrow(union_df) == 0) return(rep(FALSE, length(peaks)))
  inw <- rep(FALSE, length(peaks))
  for (i in seq_len(nrow(union_df))) {
    inw <- inw | (peaks >= union_df$start[[i]] & peaks <= union_df$end[[i]])
  }
  inw
}

sample_from_intervals <- function(intervals_df, n) {
  stopifnot(n >= 1)
  if (nrow(intervals_df) == 0) stop("No allowed intervals to sample from.")
  lens <- pmax(0, intervals_df$end - intervals_df$start)
  total <- sum(lens)
  if (total <= 0) stop("Allowed intervals total length <= 0.")
  u <- runif(n, min = 0, max = total)
  cum <- cumsum(lens)
  out <- numeric(n)
  for (i in seq_len(n)) {
    k <- which(cum >= u[[i]])[1]
    prev <- if (k == 1) 0 else cum[[k - 1]]
    out[[i]] <- intervals_df$start[[k]] + (u[[i]] - prev)
  }
  out
}

build_teacher_allowed_intervals <- function(wrist_csv, T_total, w,
                                            conf_min = 0.50,
                                            shoulder_ok_mode = c("either", "both")) {
  shoulder_ok_mode <- match.arg(shoulder_ok_mode)
  stopifnot(file.exists(wrist_csv))
  
  dat <- readr::read_csv(wrist_csv, show_col_types = FALSE)
  names(dat) <- clean_names(names(dat))
  
  col_t <- pick_col(dat, c("time_sec", "t_sec", "time_s", "time"))
  if (is.na(col_t)) stop("wrist_tracks missing time column (time_sec/t_sec/time).")
  
  t <- suppressWarnings(as.numeric(dat[[col_t]]))
  ok_t <- is.finite(t)
  dat <- dat[ok_t, , drop = FALSE]
  t <- t[ok_t]
  
  col_conf <- pick_col(dat, c("conf_det", "conf", "det_conf"))
  conf_ok <- if (is.na(col_conf)) rep(TRUE, length(t)) else {
    x <- suppressWarnings(as.numeric(dat[[col_conf]]))
    is.finite(x) & (x >= conf_min)
  }
  
  col_ls <- pick_col(dat, c("ls_vis", "left_shoulder_vis", "ls_visible"))
  col_rs <- pick_col(dat, c("rs_vis", "right_shoulder_vis", "rs_visible"))
  
  shoulders_ok <- if (!is.na(col_ls) && !is.na(col_rs)) {
    ls <- suppressWarnings(as.numeric(dat[[col_ls]]))
    rs <- suppressWarnings(as.numeric(dat[[col_rs]]))
    ls_ok <- is.finite(ls) & (ls > 0.5)
    rs_ok <- is.finite(rs) & (rs > 0.5)
    if (shoulder_ok_mode == "both") (ls_ok & rs_ok) else (ls_ok | rs_ok)
  } else {
    rep(TRUE, length(t))
  }
  
  teacher_present <- conf_ok & shoulders_ok
  
  ord <- order(t)
  t <- pmin(pmax(t[ord], 0), T_total)
  teacher_present <- teacher_present[ord]
  teacher_present[t < 0 | t > T_total] <- FALSE
  
  if (!any(teacher_present)) {
    stop("No teacher_present TRUE in wrist_tracks under current criteria.")
  }
  
  r <- rle(teacher_present)
  ends <- cumsum(r$lengths)
  starts <- ends - r$lengths + 1
  idx_true <- which(r$values)
  
  intervals <- lapply(idx_true, function(k) {
    i1 <- starts[[k]]; i2 <- ends[[k]]
    tibble(start = t[[i1]], end = t[[i2]])
  }) %>% bind_rows()
  
  allowed <- intervals %>%
    mutate(start = pmax(start, w), end = pmin(end, T_total - w)) %>%
    filter(end > start) %>%
    union_intervals()
  
  if (interval_length(allowed) <= 0) {
    stop("Teacher-only allowed time is zero after clipping to [w, T_total-w].")
  }
  allowed
}

# ----------------------------
# (2) Load points (Alignment-first for Confidence)
# ----------------------------
read_points_alignment_first <- function(xlsx) {
  if (sheet_exists(xlsx, "Alignment")) {
    A <- openxlsx::read.xlsx(xlsx, sheet = "Alignment")
    names(A) <- clean_names(names(A))
    
    col_t    <- pick_col(A, c("t_struct_sec","t_sec","time_sec","time_s","time"))
    col_pid  <- pick_col(A, c("PointID","pointid","point_id"))
    col_conf <- pick_col(A, c("Confidence","conf","confidence"))
    
    if (is.na(col_t)) stop("Alignment sheet exists but missing t_struct_sec (or equivalent).")
    
    points <- tibble(
      PointID = if (!is.na(col_pid)) as.character(A[[col_pid]]) else NA_character_,
      t_struct_sec = suppressWarnings(as.numeric(A[[col_t]])),
      Confidence = if (!is.na(col_conf)) suppressWarnings(as.numeric(A[[col_conf]])) else NA_real_
    ) %>% filter(is.finite(t_struct_sec))
    
    return(list(points = points, conf_available = !is.na(col_conf), source = "Alignment"))
  }
  
  if (sheet_exists(xlsx, "StructuralPoints")) {
    SP <- openxlsx::read.xlsx(xlsx, sheet = "StructuralPoints")
    names(SP) <- clean_names(names(SP))
    
    col_t   <- pick_col(SP, c("t_struct_sec","t_sec","time_sec","time_s","time"))
    col_pid <- pick_col(SP, c("PointID","pointid","point_id"))
    
    if (is.na(col_t)) stop("StructuralPoints sheet exists but missing t_struct_sec (or equivalent).")
    
    points <- tibble(
      PointID = if (!is.na(col_pid)) as.character(SP[[col_pid]]) else NA_character_,
      t_struct_sec = suppressWarnings(as.numeric(SP[[col_t]])),
      Confidence = NA_real_
    ) %>% filter(is.finite(t_struct_sec))
    
    return(list(points = points, conf_available = FALSE, source = "StructuralPoints"))
  }
  
  stop("Neither Alignment nor StructuralPoints sheet exists.")
}

tmp <- read_points_alignment_first(align_xlsx)
points_df <- tmp$points
conf_available <- tmp$conf_available
message("Points loaded from: ", tmp$source, " | conf_available=", conf_available)

points_all <- points_df %>% distinct(t_struct_sec, .keep_all = TRUE)
points_conf3 <- if (conf_available) {
  points_df %>% filter(Confidence == 3) %>% distinct(t_struct_sec, .keep_all = TRUE)
} else {
  points_df[0, ]
}

message("n_points_all=", nrow(points_all), " | n_points_conf3=", nrow(points_conf3))

# ----------------------------
# (3) Load events (peaks) from GestureEvents
# ----------------------------
stopifnot(sheet_exists(align_xlsx, "GestureEvents"))
E <- openxlsx::read.xlsx(align_xlsx, sheet = "GestureEvents")
names(E) <- clean_names(names(E))

# Template rule:
# GestureEvents$t_sec IS peak time (peak_sec)
col_peak <- pick_col(E, c(
  "peak_sec", "t_sec",
  "t_peak", "tpeak",
  "t_peak_sec",
  "peak_sec_proxy",
  "peak", "peak_s", "peak_time"
))
if (is.na(col_peak)) stop("GestureEvents missing peak time column (peak_sec / t_sec / ...).")

col_eid <- pick_col(E, c("EventID","eventid","EventId","event_id"))

peaks <- suppressWarnings(as.numeric(E[[col_peak]]))
peaks <- peaks[is.finite(peaks)]
stopifnot(length(peaks) >= 1)

peaks_sorted <- sort(peaks)

# robust total time: cover both events and points
T_events <- max(peaks_sorted, na.rm = TRUE)
T_points <- max(points_all$t_struct_sec, na.rm = TRUE)
T_total  <- max(T_events, T_points)

if (!is.finite(T_total) || T_total <= 2 * window_s) {
  stop("T_total too small/invalid relative to window_s. T_total=", T_total, " window_s=", window_s)
}
if (T_points > T_events + 1e-6) {
  warning("Points extend beyond max event peak time; using T_total=max(events, points). ",
          "T_points=", signif(T_points,4), " T_events=", signif(T_events,4))
}

make_events_flags <- function(in_flag_vec) {
  out <- tibble(peak_sec = peaks_sorted, in_struct_window = in_flag_vec)
  if (!is.na(col_eid)) {
    EE <- tibble(
      EventID = as.character(E[[col_eid]]),
      peak_sec = suppressWarnings(as.numeric(E[[col_peak]]))
    ) %>% filter(is.finite(peak_sec)) %>% arrange(peak_sec)
    out <- bind_cols(EE %>% select(EventID), out)
  }
  out
}

# ----------------------------
# (4) RUN ONE: pointset × null_mode
# ----------------------------
run_one <- function(pointset_name, t_struct_vec, null_mode) {
  stopifnot(null_mode %in% c("global","teacher_only"))
  
  t_struct <- t_struct_vec[t_struct_vec >= 0 & t_struct_vec <= T_total]
  t_struct <- unique(t_struct)
  
  n_points_obs <- length(t_struct)
  if (n_points_obs == 0) {
    warning("Skipping: ", pointset_name, " × ", null_mode, " (n_points_obs=0)")
    return(NULL)
  }
  
  win_obs <- make_windows(t_struct, window_s) %>% safe_clip_windows(., 0, T_total)
  union_obs <- union_intervals(win_obs)
  
  T_in  <- interval_length(union_obs)
  T_out <- T_total - T_in
  
  N_in  <- count_in_union(peaks_sorted, union_obs)
  N_all <- length(peaks_sorted)
  N_out <- N_all - N_in
  
  r_in  <- ifelse(T_in  > 0, N_in / T_in, NA_real_)
  r_out <- ifelse(T_out > 0, N_out / T_out, NA_real_)
  lift  <- ifelse(is.finite(r_in) & is.finite(r_out) & r_out > 0, r_in / r_out, NA_real_)
  delta <- ifelse(is.finite(r_in) & is.finite(r_out), r_in - r_out, NA_real_)
  
  in_flag <- flag_in_union(peaks_sorted, union_obs)
  events_flags_df <- make_events_flags(in_flag)
  
  out_dir <- file.path(results_root, paste0("time_enrichment_", pointset_name, "_", null_mode))
  dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  
  # allowed intervals for permutation
  if (null_mode == "global") {
    allowed_intervals <- tibble(start = window_s, end = T_total - window_s)
  } else {
    if (is.null(wrist_csv) || !nzchar(wrist_csv)) {
      stop("teacher_only requested but --wrist not provided.")
    }
    if (!file.exists(wrist_csv)) stop("teacher_only requires existing wrist_csv: ", wrist_csv)
    
    allowed_intervals <- build_teacher_allowed_intervals(
      wrist_csv = wrist_csv,
      T_total = T_total,
      w = window_s,
      conf_min = 0.50,
      shoulder_ok_mode = "either"
    )
  }
  allowed_total <- interval_length(allowed_intervals)
  
  set.seed(seed)
  perm_rows <- vector("list", n_perm)
  
  for (b in seq_len(n_perm)) {
    t_perm <- if (null_mode == "global") {
      runif(n_points_obs, min = window_s, max = T_total - window_s)
    } else {
      sample_from_intervals(allowed_intervals, n_points_obs)
    }
    
    win_p <- make_windows(t_perm, window_s) %>% safe_clip_windows(., 0, T_total)
    union_p <- union_intervals(win_p)
    
    T_in_p  <- interval_length(union_p)
    T_out_p <- T_total - T_in_p
    
    N_in_p  <- count_in_union(peaks_sorted, union_p)
    N_out_p <- N_all - N_in_p
    
    r_in_p  <- ifelse(T_in_p  > 0, N_in_p / T_in_p, NA_real_)
    r_out_p <- ifelse(T_out_p > 0, N_out_p / T_out_p, NA_real_)
    lift_p  <- ifelse(is.finite(r_in_p) & is.finite(r_out_p) & r_out_p > 0, r_in_p / r_out_p, NA_real_)
    delta_p <- ifelse(is.finite(r_in_p) & is.finite(r_out_p), r_in_p - r_out_p, NA_real_)
    
    perm_rows[[b]] <- tibble(
      iter = b,
      pointset = pointset_name,
      null_mode = null_mode,
      n_points = n_points_obs,
      allowed_time = allowed_total,
      T_in = T_in_p, T_out = T_out_p,
      N_in = N_in_p, N_out = N_out_p,
      r_in = r_in_p, r_out = r_out_p,
      lift = lift_p,
      delta_rate = delta_p
    )
  }
  
  perm_df <- bind_rows(perm_rows)
  
  if (is.finite(delta)) {
    more_extreme <- sum(abs(perm_df$delta_rate) >= abs(delta), na.rm = TRUE)
    p_raw   <- more_extreme / n_perm
    p_plus1 <- (more_extreme + 1) / (n_perm + 1)
  } else {
    p_raw <- NA_real_
    p_plus1 <- NA_real_
  }
  
  summary_df <- tibble(
    tag = TAG,
    pointset = pointset_name,
    window_s = window_s,
    n_struct_points = n_points_obs,
    T_total = T_total,
    T_in = T_in, T_out = T_out,
    N_all = N_all, N_in = N_in, N_out = N_out,
    r_in = r_in, r_out = r_out,
    lift = lift,
    delta_rate = delta,
    null_mode = null_mode,
    allowed_time = allowed_total,
    n_perm = n_perm,
    seed = seed,
    p_value_raw = p_raw,
    p_value_plus1 = p_plus1
  )
  
  # write
  write_csv(summary_df, file.path(out_dir, paste0(TAG, "_summary_overall.csv")))
  write_csv(union_obs %>% mutate(tag = TAG, pointset = pointset_name, window_s = window_s),
            file.path(out_dir, paste0(TAG, "_windows_union.csv")))
  write_csv(events_flags_df, file.path(out_dir, paste0(TAG, "_events_within_flags.csv")))
  write_csv(perm_df, file.path(out_dir, paste0(TAG, "_perm_overall.csv")))
  write_csv(allowed_intervals, file.path(out_dir, paste0(TAG, "_allowed_intervals.csv")))
  
  message("DONE: ", pointset_name, " × ", null_mode,
          " | lift=", signif(lift, 4),
          " | delta=", signif(delta, 4),
          " | p+1=", signif(p_plus1, 4),
          " | out_dir=", normalizePath(out_dir, winslash = "/"))
  
  summary_df
}

# ----------------------------
# (5) RUN ALL COMBINATIONS
# ----------------------------
master <- list()

master[[length(master) + 1]] <- run_one("all_points", points_all$t_struct_sec, "global")
master[[length(master) + 1]] <- run_one("all_points", points_all$t_struct_sec, "teacher_only")

if (nrow(points_conf3) > 0) {
  master[[length(master) + 1]] <- run_one("conf3_only", points_conf3$t_struct_sec, "global")
  master[[length(master) + 1]] <- run_one("conf3_only", points_conf3$t_struct_sec, "teacher_only")
} else {
  warning("No conf3 points (or Confidence unavailable). Skipping conf3_only runs.")
}

master_df <- bind_rows(master)

master_out <- file.path(results_root, "time_enrichment_MASTER_summary.csv")
write_csv(master_df, master_out)

message("ALL DONE: ", TAG)
message("MASTER -> ", normalizePath(master_out, winslash = "/"))
