# ============================================================
# Script 05 (FINAL / CLEAN / GitHub-ready):
# Time-based enrichment around structural points (4 runs)
#
# CLEAN RULES:
#   - TAG is REQUIREDC (REQUIRED) / no default tag
#   - No absolute/local machine paths
#   - Root = Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
#   - No template involvement (reads alignment xlsx only)
#
# INPUT:
#   exports/<TAG>/<TAG>_alignment.xlsx
#   exports/wrist_tracks/wrist_tracks_<TAG>.csv   (required for teacher_only)
#
# OUTPUT:
#   results/<TAG>/time_enrichment_<pointset>_<null_mode>/
#   results/<TAG>/time_enrichment_MASTER_summary.csv
#
# CLI:
#   Rscript scripts/05_time_enrichment_FINAL_4runs.R <TAG> [window_s] [n_perm] [seed]
# Example:
#   Rscript scripts/05_time_enrichment_FINAL_4runs.R m02 1.5 5000 1
# ============================================================

suppressPackageStartupMessages({
  library(openxlsx)
  library(dplyr)
  library(stringr)
  library(readr)
  library(tibble)
})

# ----------------------------
# (0) SETTINGS / CLI  (TAG REQUIRED)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 1 || !nzchar(args[[1]])) {
  stop(
    "Missing TAG.\n",
    "Usage:\n  Rscript scripts/05_time_enrichment_FINAL_4runs.R <TAG> [window_s] [n_perm] [seed]\n",
    call. = FALSE
  )
}

TAG      <- args[[1]]
window_s <- if (length(args) >= 2 && nzchar(args[[2]])) as.numeric(args[[2]]) else 1.5
n_perm   <- if (length(args) >= 3 && nzchar(args[[3]])) as.integer(args[[3]]) else 5000L
seed     <- if (length(args) >= 4 && nzchar(args[[4]])) as.integer(args[[4]]) else 1L

stopifnot(is.finite(window_s), window_s > 0)
stopifnot(is.finite(n_perm), n_perm >= 100)

project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())
exports_root <- file.path(project_root, "exports")
results_root <- file.path(project_root, "results")

align_xlsx <- file.path(exports_root, TAG, paste0(TAG, "_alignment.xlsx"))
if (!file.exists(align_xlsx)) {
  stop("Cannot find alignment xlsx: ", align_xlsx, call. = FALSE)
}
message("Using align_xlsx: ", normalizePath(align_xlsx, winslash = "/"))

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

pick_col <- function(df, candidates) {
  nms <- names(df)
  low <- tolower(nms)
  for (cand in candidates) {
    j <- which(low == tolower(cand))
    if (length(j) >= 1) return(nms[j[[1]]])
  }
  NA_character_
}

as_num <- function(x) suppressWarnings(as.numeric(as.character(x)))

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
  
  out_s <- numeric(0); out_e <- numeric(0)
  cs <- x$start[[1]]; ce <- x$end[[1]]
  
  if (nrow(x) >= 2) {
    for (i in 2:nrow(x)) {
      s <- x$start[[i]]; e <- x$end[[i]]
      if (!is.finite(s) || !is.finite(e) || !is.finite(cs) || !is.finite(ce)) next
      if (s <= ce) ce <- max(ce, e)
      else {
        out_s <- c(out_s, cs); out_e <- c(out_e, ce)
        cs <- s; ce <- e
      }
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
  if (nrow(intervals_df) == 0) stop("No allowed intervals to sample from.", call. = FALSE)
  lens <- pmax(0, intervals_df$end - intervals_df$start)
  total <- sum(lens)
  if (total <= 0) stop("Allowed intervals total length <= 0.", call. = FALSE)
  
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
                                            shoulder_ok_mode = c("either", "both"),
                                            shoulder_vis_min = 0.50) {
  shoulder_ok_mode <- match.arg(shoulder_ok_mode)
  
  dat <- readr::read_csv(wrist_csv, show_col_types = FALSE)
  names(dat) <- std_names(names(dat))
  
  col_t <- pick_col(dat, c("time_sec","t_sec","time_s","time"))
  if (is.na(col_t)) stop("wrist_tracks missing time column (time_sec/t_sec/time).", call. = FALSE)
  
  t <- suppressWarnings(as.numeric(dat[[col_t]]))
  ok_t <- is.finite(t)
  dat <- dat[ok_t, , drop = FALSE]
  t <- t[ok_t]
  
  col_conf <- pick_col(dat, c("conf_det","conf","det_conf"))
  conf_ok <- if (is.na(col_conf)) rep(TRUE, length(t)) else {
    x <- suppressWarnings(as.numeric(dat[[col_conf]]))
    is.finite(x) & (x >= conf_min)
  }
  
  col_ls <- pick_col(dat, c("ls_vis","left_shoulder_vis","ls_visible"))
  col_rs <- pick_col(dat, c("rs_vis","right_shoulder_vis","rs_visible"))
  
  shoulders_ok <- if (!is.na(col_ls) && !is.na(col_rs)) {
    ls <- suppressWarnings(as.numeric(dat[[col_ls]]))
    rs <- suppressWarnings(as.numeric(dat[[col_rs]]))
    ls_ok <- is.finite(ls) & (ls >= shoulder_vis_min)
    rs_ok <- is.finite(rs) & (rs >= shoulder_vis_min)
    if (shoulder_ok_mode == "both") (ls_ok & rs_ok) else (ls_ok | rs_ok)
  } else {
    rep(TRUE, length(t))
  }
  
  teacher_present <- conf_ok & shoulders_ok
  
  ord <- order(t)
  t <- pmin(pmax(t[ord], 0), T_total)
  teacher_present <- teacher_present[ord]
  teacher_present[t < 0 | t > T_total] <- FALSE
  
  if (!any(teacher_present)) stop("No teacher_present TRUE in wrist_tracks under current criteria.", call. = FALSE)
  
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
  
  if (interval_length(allowed) <= 0) stop("Teacher-only allowed time is zero after clipping.", call. = FALSE)
  allowed
}

# ----------------------------
# (2) READ SHEETS ONCE
# ----------------------------
sheets <- openxlsx::getSheetNames(align_xlsx)

# ----------------------------
# (3) Load points (Alignment-first; fallback StructuralPoints)
# ----------------------------
read_points_alignment_first <- function(xlsx, sheets_vec) {
  if ("Alignment" %in% sheets_vec) {
    A <- openxlsx::read.xlsx(xlsx, sheet = "Alignment")
    names(A) <- std_names(names(A))
    
    col_t    <- pick_col(A, c("t_struct_sec","t_sec","time_sec","time_s","time"))
    col_pid  <- pick_col(A, c("pointid","point_id","id"))
    col_conf <- pick_col(A, c("confidence_1to3","confidence","conf"))
    
    if (is.na(col_t)) stop("Alignment exists but missing t_struct_sec (or equivalent).", call. = FALSE)
    
    points <- tibble(
      PointID = if (!is.na(col_pid)) as.character(A[[col_pid]]) else as.character(seq_len(nrow(A))),
      t_struct_sec = as_num(A[[col_t]]),
      Confidence = if (!is.na(col_conf)) as_num(A[[col_conf]]) else NA_real_
    ) %>% filter(is.finite(t_struct_sec))
    
    list(points = points, conf_available = !is.na(col_conf), source = "Alignment")
  } else if ("StructuralPoints" %in% sheets_vec) {
    SP <- openxlsx::read.xlsx(xlsx, sheet = "StructuralPoints")
    names(SP) <- std_names(names(SP))
    
    col_t   <- pick_col(SP, c("t_struct_sec","t_sec","time_sec","time_s","time"))
    col_pid <- pick_col(SP, c("pointid","point_id","id"))
    
    if (is.na(col_t)) stop("StructuralPoints exists but missing t_struct_sec (or equivalent).", call. = FALSE)
    
    points <- tibble(
      PointID = if (!is.na(col_pid)) as.character(SP[[col_pid]]) else as.character(seq_len(nrow(SP))),
      t_struct_sec = as_num(SP[[col_t]]),
      Confidence = NA_real_
    ) %>% filter(is.finite(t_struct_sec))
    
    list(points = points, conf_available = FALSE, source = "StructuralPoints")
  } else {
    stop("Neither Alignment nor StructuralPoints sheet exists.", call. = FALSE)
  }
}

tmp <- read_points_alignment_first(align_xlsx, sheets)
points_df <- tmp$points
conf_available <- tmp$conf_available
message("Points loaded from: ", tmp$source, " | conf_available=", conf_available)

points_all <- points_df %>% distinct(t_struct_sec, .keep_all = TRUE)
points_conf3 <- if (conf_available) {
  points_df %>% filter(Confidence == 3) %>% distinct(t_struct_sec, .keep_all = TRUE)
} else {
  points_df[0, ]
}

# ----------------------------
# (4) Load events (peaks)
# ----------------------------
if (!("GestureEvents" %in% sheets)) stop("Missing sheet: GestureEvents", call. = FALSE)

E <- openxlsx::read.xlsx(align_xlsx, sheet = "GestureEvents")
names(E) <- std_names(names(E))

col_peak <- pick_col(E, c("peak_sec","t_peak","t_peak_sec","peak","peak_s","start_sec"))
if (is.na(col_peak)) stop("GestureEvents must contain peak_sec (or equivalent).", call. = FALSE)

col_eid <- pick_col(E, c("eventid","event_id","eventid_auto","id"))

peaks <- as_num(E[[col_peak]])
peaks <- peaks[is.finite(peaks)]
if (length(peaks) < 1) stop("No finite peaks in GestureEvents.", call. = FALSE)

peaks_sorted <- sort(peaks)

# Use total duration from max(peaks, points) to be safe
T_total <- max(c(peaks_sorted, points_all$t_struct_sec), na.rm = TRUE)
if (!is.finite(T_total) || T_total <= 2 * window_s) stop("T_total too small relative to window_s.", call. = FALSE)

make_events_flags <- function(in_flag_vec) {
  out <- tibble(peak_sec = peaks_sorted, in_struct_window = in_flag_vec)
  
  if (!is.na(col_eid)) {
    EE <- tibble(
      EventID = as.character(E[[col_eid]]),
      peak_sec = as_num(E[[col_peak]])
    ) %>% filter(is.finite(peak_sec)) %>% arrange(peak_sec)
    
    out <- EE %>%
      left_join(out, by = "peak_sec") %>%
      mutate(in_struct_window = ifelse(is.na(in_struct_window), FALSE, in_struct_window))
  }
  
  out
}

# ----------------------------
# (5) RUN ONE: pointset × null_mode
# ----------------------------
run_one <- function(pointset_name, t_struct_vec, null_mode) {
  stopifnot(null_mode %in% c("global","teacher_only"))
  
  t_struct <- t_struct_vec[is.finite(t_struct_vec)]
  t_struct <- unique(t_struct[t_struct >= 0 & t_struct <= T_total])
  
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
  
  out_dir <- file.path(results_root, TAG, paste0("time_enrichment_", pointset_name, "_", null_mode))
  dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
  
  # allowed intervals
  if (null_mode == "global") {
    allowed_intervals <- tibble(start = window_s, end = T_total - window_s)
  } else {
    wrist_csv <- file.path(exports_root, "wrist_tracks", paste0("wrist_tracks_", TAG, ".csv"))
    if (!file.exists(wrist_csv)) {
      stop("teacher_only requires: ", wrist_csv, call. = FALSE)
    }
    
    allowed_intervals <- build_teacher_allowed_intervals(
      wrist_csv = wrist_csv,
      T_total = T_total,
      w = window_s,
      conf_min = 0.50,
      shoulder_ok_mode = "either",
      shoulder_vis_min = 0.50
    )
  }
  
  allowed_total <- interval_length(allowed_intervals)
  
  # permutation
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
  
  # p-values on |delta|, plus-one correction
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
  
  readr::write_csv(summary_df, file.path(out_dir, paste0(TAG, "_summary_overall.csv")))
  readr::write_csv(union_obs %>% mutate(tag = TAG, pointset = pointset_name, window_s = window_s),
                   file.path(out_dir, paste0(TAG, "_windows_union.csv")))
  readr::write_csv(events_flags_df, file.path(out_dir, paste0(TAG, "_events_within_flags.csv")))
  readr::write_csv(perm_df, file.path(out_dir, paste0(TAG, "_perm_overall.csv")))
  readr::write_csv(allowed_intervals, file.path(out_dir, paste0(TAG, "_allowed_intervals.csv")))
  
  message("DONE: ", pointset_name, " × ", null_mode,
          " | lift=", signif(lift, 4),
          " | delta=", signif(delta, 4),
          " | p+1=", signif(p_plus1, 4))
  
  summary_df
}

# ----------------------------
# (6) RUN ALL COMBINATIONS
# ----------------------------
master <- list()

master[[length(master) + 1]] <- run_one("all_points",  points_all$t_struct_sec,  "global")
master[[length(master) + 1]] <- run_one("all_points",  points_all$t_struct_sec,  "teacher_only")

if (nrow(points_conf3) > 0) {
  master[[length(master) + 1]] <- run_one("conf3_only", points_conf3$t_struct_sec, "global")
  master[[length(master) + 1]] <- run_one("conf3_only", points_conf3$t_struct_sec, "teacher_only")
} else {
  warning("No conf3 points found (or Confidence unavailable). Skipping conf3_only runs.")
}

master_df <- dplyr::bind_rows(master)

tag_results_dir <- file.path(results_root, TAG)
dir.create(tag_results_dir, recursive = TRUE, showWarnings = FALSE)
master_out <- file.path(tag_results_dir, "time_enrichment_MASTER_summary.csv")
readr::write_csv(master_df, master_out)

message("ALL DONE: ", TAG)
message("MASTER -> ", normalizePath(master_out, winslash = "/"))
