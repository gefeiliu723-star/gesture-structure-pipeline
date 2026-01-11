#!/user/bin/env Rscript
# ============================================================
# Script 05: Analysis (Node ±window) + Macro breakdown + Permutation
# INPUT:  <project_root>/exports/<tag>/<tag>_alignment.xlsx   (ONLY THIS)
# OUTPUT: <project_root>/results/<tag>/ with <tag>_ prefix
# NO KEEP, NO GATING, NO CONF FILTERS
# Permutation null: circular time-shift of structural node times
# ============================================================
# EN Fixes:
#  1) Prevent NA in logical contexts: hard-clean event times; skip NA/Inf in loops
#  2) Macro case-insensitive: macro / Macro supported
#  3) PointID case-insensitive: pointid / PointID supported
#  4) Structural time column mapping: time_s / time_sec / t_sec / time supported
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# CONFIG (GitHub-friendly)
# ----------------------------
# Repo layout assumption:
#   project_root/
#     exports/<tag>/<tag>_alignment.xlsx
#     results/<tag>/
#
# If you run from repo root, project_root="." works.
project_root <- Sys.getenv("MKAP_ROOT", unset = ".")
tag          <- Sys.getenv("MKAP_TAG",  unset = "m01")

window_s     <- as.numeric(Sys.getenv("MKAP_WINDOW_S", unset = "2"))

# Permutation settings
B_perm       <- as.integer(Sys.getenv("MKAP_B",    unset = "2000"))
seed_perm    <- as.integer(Sys.getenv("MKAP_SEED", unset = "42"))

# ----------------------------
# PATHS (relative)
# ----------------------------
export_dir <- file.path(project_root, "exports", tag)
result_dir <- file.path(project_root, "results", tag)
dir.create(result_dir, showWarnings = FALSE, recursive = TRUE)

xlsx_path  <- file.path(export_dir, paste0(tag, "_alignment.xlsx"))

out_file <- function(stem) file.path(result_dir, paste0(tag, "_", stem, ".csv"))

# ----------------------------
# HELPERS
# ----------------------------
safe_div <- function(a, b) ifelse(is.na(b) | b == 0, NA_real_, a / b)

merge_intervals <- function(df, start_col = "start", end_col = "end") {
  df <- df[order(df[[start_col]], df[[end_col]]), , drop = FALSE]
  if (nrow(df) == 0) return(tibble(start = numeric(0), end = numeric(0)))
  
  out_s <- numeric(0); out_e <- numeric(0)
  cur_s <- df[[start_col]][1]; cur_e <- df[[end_col]][1]
  
  if (nrow(df) >= 2) {
    for (i in 2:nrow(df)) {
      s <- df[[start_col]][i]; e <- df[[end_col]][i]
      if (is.na(s) || is.na(e) || is.na(cur_s) || is.na(cur_e)) next
      if (s <= cur_e) cur_e <- max(cur_e, e)
      else {
        out_s <- c(out_s, cur_s); out_e <- c(out_e, cur_e)
        cur_s <- s; cur_e <- e
      }
    }
  }
  out_s <- c(out_s, cur_s); out_e <- c(out_e, cur_e)
  tibble(start = out_s, end = out_e)
}

point_in_intervals <- function(t_vec, intervals) {
  if (nrow(intervals) == 0) return(rep(FALSE, length(t_vec)))
  res <- rep(FALSE, length(t_vec))
  j <- 1L
  for (i in seq_along(t_vec)) {
    t <- t_vec[i]
    if (!is.finite(t)) next  # HARDEN: skip NA/Inf
    
    while (j <= nrow(intervals) && is.finite(intervals$end[j]) && t > intervals$end[j]) {
      j <- j + 1L
    }
    if (j <= nrow(intervals) &&
        is.finite(intervals$start[j]) &&
        is.finite(intervals$end[j]) &&
        t >= intervals$start[j] &&
        t <= intervals$end[j]) {
      res[i] <- TRUE
    }
  }
  res
}

# ----------------------------
# CHECK INPUT
# ----------------------------
if (!file.exists(xlsx_path)) {
  stop("Missing input workbook: ", xlsx_path, "\n",
       "Expected: <project_root>/exports/<tag>/<tag>_alignment.xlsx")
}

# ----------------------------
# LOAD EVENTS (FROM XLSX)
# ----------------------------
events <- read_excel(xlsx_path, sheet = "GestureEvents") %>% as_tibble()

# Accept common column name variants for start/end (case-insensitive)
events <- events %>%
  rename_with(~"start_sec", matches("^start_?sec$|^event_start_?sec$|^event_start_?s$|^start$", ignore.case = TRUE)) %>%
  rename_with(~"end_sec",   matches("^end_?sec$|^event_end_?sec$|^event_end_?s$|^end$", ignore.case = TRUE))

if (!all(c("start_sec", "end_sec") %in% names(events))) {
  stop("GestureEvents must contain start/end columns (e.g., start_sec & end_sec).")
}

# Build event_id robustly (event_id / eventid / EventID / fallback row_number)
if ("event_id" %in% names(events)) {
  events$event_id <- as.character(events$event_id)
} else if ("eventid" %in% names(events)) {
  events$event_id <- as.character(events$eventid)
} else if ("EventID" %in% names(events)) {
  events$event_id <- as.character(events$EventID)
} else if ("EventId" %in% names(events)) {
  events$event_id <- as.character(events$EventId)
} else {
  events$event_id <- as.character(seq_len(nrow(events)))
}

events <- events %>%
  mutate(
    start_sec = suppressWarnings(as.numeric(start_sec)),
    end_sec   = suppressWarnings(as.numeric(end_sec))
  ) %>%
  # HARDEN: drop NA/invalid events (prevents while/if NA crash)
  filter(is.finite(start_sec), is.finite(end_sec), end_sec >= start_sec) %>%
  mutate(t_peak = (start_sec + end_sec) / 2) %>%
  filter(is.finite(t_peak))

if (nrow(events) == 0) stop("No valid events after cleaning (start_sec/end_sec all NA or invalid).")

T_total <- suppressWarnings(max(events$end_sec, na.rm = TRUE))
if (!is.finite(T_total) || is.na(T_total) || T_total <= 0) stop("Bad T_total inferred from events end_sec.")

message(sprintf("Loaded events: n=%d | T_total=%.3f sec", nrow(events), T_total))

# ----------------------------
# LOAD STRUCTURAL POINTS (FROM XLSX)
# ONLY use: time + Macro (no filters)
# ----------------------------
sp <- read_excel(xlsx_path, sheet = "StructuralPoints") %>% as_tibble()

# Map time column -> time_s (case-insensitive)
time_col <- names(sp)[str_detect(names(sp), regex("^time_s$|^time_sec$|^t_sec$|^time$", ignore_case = TRUE))][1]
if (is.na(time_col) || !nzchar(time_col)) {
  stop("StructuralPoints must contain a time column: time_s (or time_sec/t_sec/time).")
}
sp <- sp %>% rename(time_s = all_of(time_col))

# Map Macro column robustly (macro/Macro)
macro_col <- names(sp)[str_detect(names(sp), regex("^macro$", ignore_case = TRUE))][1]
if (is.na(macro_col) || !nzchar(macro_col)) {
  sp <- sp %>% mutate(Macro = NA_character_)
} else {
  sp <- sp %>% rename(Macro = all_of(macro_col)) %>% mutate(Macro = as.character(.data$Macro))
}

# Map PointID column robustly (pointid/PointID/point_id/id). optional.
pid_col <- names(sp)[str_detect(names(sp), regex("^pointid$|^point_id$|^id$", ignore_case = TRUE))][1]
if (!is.na(pid_col) && nzchar(pid_col)) {
  sp <- sp %>% rename(PointID = all_of(pid_col)) %>% mutate(PointID = as.character(.data$PointID))
} else {
  sp <- sp %>% mutate(PointID = as.character(row_number()))
}

sp2 <- sp %>%
  mutate(
    time_s = suppressWarnings(as.numeric(time_s)),
    Macro  = ifelse(is.na(Macro) | str_trim(Macro) == "", "UNLABELED", str_trim(Macro))
  ) %>%
  filter(is.finite(time_s))  # HARDEN

if (nrow(sp2) == 0) stop("No valid structural points after cleaning (time_s all NA?).")

message(sprintf("Loaded structural points: n=%d | unique Macro=%d",
                nrow(sp2), length(unique(sp2$Macro))))

sp_windows <- sp2 %>%
  transmute(
    PointID = .data$PointID,
    Macro   = .data$Macro,
    start   = pmax(time_s - window_s, 0),
    end     = pmin(time_s + window_s, T_total)
  ) %>%
  filter(is.finite(start), is.finite(end), end >= start)

# ----------------------------
# OVERALL UNION (all points)
# ----------------------------
overall_union <- merge_intervals(sp_windows %>% select(start, end))
T_in  <- sum(overall_union$end - overall_union$start)
T_out <- T_total - T_in

events <- events %>% mutate(struct_overall = point_in_intervals(t_peak, overall_union))

N_all <- nrow(events)
N_in  <- sum(events$struct_overall, na.rm = TRUE)
N_out <- N_all - N_in

Share <- safe_div(N_in, N_all)
r_in  <- safe_div(N_in, T_in)
r_out <- safe_div(N_out, T_out)
Lift  <- safe_div(r_in, r_out)

summary_overall <- tibble(
  tag = tag,
  window_s = window_s,
  T_total = T_total,
  N_all = N_all,
  N_in = N_in,
  N_out = N_out,
  Share = Share,
  T_in = T_in,
  T_out = T_out,
  r_in = r_in,
  r_out = r_out,
  Lift = Lift
)

write_csv(summary_overall, out_file("summary_overall"))

# ----------------------------
# MACRO SUMMARY
# ----------------------------
macros_main <- sp_windows %>% distinct(Macro) %>% pull(Macro)
baseline_r_out <- r_out

summary_by_macro <- lapply(macros_main, function(m) {
  w_m <- sp_windows %>% filter(Macro == m) %>% select(start, end)
  union_m <- merge_intervals(w_m)
  T_in_m <- sum(union_m$end - union_m$start)
  
  in_m <- point_in_intervals(events$t_peak, union_m)
  N_in_m <- sum(in_m, na.rm = TRUE)
  
  r_in_m <- safe_div(N_in_m, T_in_m)
  Lift_m <- safe_div(r_in_m, baseline_r_out)
  
  tibble(
    tag = tag,
    Macro = m,
    N_nodes = sum(sp_windows$Macro == m, na.rm = TRUE),
    T_in = T_in_m,
    N_in = N_in_m,
    r_in = r_in_m,
    r_out = baseline_r_out,
    Lift = Lift_m
  )
}) %>%
  bind_rows() %>%
  arrange(desc(Lift))

write_csv(summary_by_macro, out_file("summary_by_macro"))

macro_union_tbl <- lapply(macros_main, function(m) {
  w_m <- sp_windows %>% filter(Macro == m) %>% select(start, end)
  union_m <- merge_intervals(w_m)
  if (nrow(union_m) == 0) return(NULL)
  union_m %>% mutate(tag = tag, Macro = m, seg_len = end - start)
}) %>% bind_rows()

write_csv(macro_union_tbl, out_file("macro_windows_union"))

# ----------------------------
# EVENT FLAGS (per macro)
# ----------------------------
event_macro_flags <- lapply(macros_main, function(m) {
  w_m <- sp_windows %>% filter(Macro == m) %>% select(start, end)
  union_m <- merge_intervals(w_m)
  tibble(!!paste0("in_", make.names(m)) := point_in_intervals(events$t_peak, union_m))
})

events_out <- bind_cols(
  events %>% select(event_id, start_sec, end_sec, t_peak, struct_overall),
  bind_cols(event_macro_flags)
)

write_csv(events_out, out_file("events_with_struct_flags"))

# ============================================================
# PERMUTATION (Circular Time-Shift)
# ============================================================
compute_lift_from_points <- function(events_t, point_times, T_total, window_s) {
  events_t <- events_t[is.finite(events_t)]
  point_times <- point_times[is.finite(point_times)]
  
  if (length(point_times) == 0 || length(events_t) == 0) {
    return(tibble(
      Lift = NA_real_, r_in = NA_real_, r_out = NA_real_,
      N_in = 0, N_out = length(events_t), T_in = 0, T_out = T_total
    ))
  }
  
  w <- tibble(
    start = pmax(point_times - window_s, 0),
    end   = pmin(point_times + window_s, T_total)
  ) %>%
    arrange(start, end) %>%
    filter(is.finite(start), is.finite(end), end >= start)
  
  if (nrow(w) == 0) {
    return(tibble(
      Lift = NA_real_, r_in = NA_real_, r_out = NA_real_,
      N_in = 0, N_out = length(events_t), T_in = 0, T_out = T_total
    ))
  }
  
  # merge intervals
  out_s <- c(); out_e <- c()
  cur_s <- w$start[1]; cur_e <- w$end[1]
  if (nrow(w) >= 2) {
    for (i in 2:nrow(w)) {
      s <- w$start[i]; e <- w$end[i]
      if (!is.finite(s) || !is.finite(e) || !is.finite(cur_e)) next
      if (s <= cur_e) cur_e <- max(cur_e, e)
      else {
        out_s <- c(out_s, cur_s); out_e <- c(out_e, cur_e)
        cur_s <- s; cur_e <- e
      }
    }
  }
  out_s <- c(out_s, cur_s); out_e <- c(out_e, cur_e)
  union <- tibble(start = out_s, end = out_e)
  
  T_in  <- sum(union$end - union$start)
  T_out <- T_total - T_in
  
  # event in union
  in_win <- rep(FALSE, length(events_t))
  j <- 1L
  for (i in seq_along(events_t)) {
    t <- events_t[i]
    if (!is.finite(t)) next
    while (j <= nrow(union) && is.finite(union$end[j]) && t > union$end[j]) j <- j + 1L
    if (j <= nrow(union) &&
        is.finite(union$start[j]) &&
        is.finite(union$end[j]) &&
        t >= union$start[j] &&
        t <= union$end[j]) {
      in_win[i] <- TRUE
    }
  }
  
  N_in  <- sum(in_win, na.rm = TRUE)
  N_out <- length(events_t) - N_in
  
  r_in  <- ifelse(T_in  > 0, N_in / T_in, NA_real_)
  r_out <- ifelse(T_out > 0, N_out / T_out, NA_real_)
  Lift  <- safe_div(r_in, r_out)
  
  tibble(Lift = Lift, r_in = r_in, r_out = r_out,
         N_in = N_in, N_out = N_out, T_in = T_in, T_out = T_out)
}

perm_test_shift <- function(events_t, point_times, T_total, window_s,
                            B = 2000, seed = 42) {
  set.seed(seed)
  
  obs_tbl <- compute_lift_from_points(events_t, point_times, T_total, window_s)
  obs <- obs_tbl$Lift
  
  null <- replicate(B, {
    delta <- runif(1, 0, T_total)
    shifted <- (point_times + delta) %% T_total
    compute_lift_from_points(events_t, shifted, T_total, window_s)$Lift
  })
  
  p_one_sided <- (1 + sum(null >= obs, na.rm = TRUE)) / (B + 1)
  ci <- quantile(null, probs = c(0.025, 0.975), na.rm = TRUE)
  
  null_mean <- mean(null, na.rm = TRUE)
  null_sd   <- sd(null, na.rm = TRUE)
  z_score   <- ifelse(is.na(null_sd) | null_sd == 0, NA_real_, (obs - null_mean) / null_sd)
  
  list(
    obs = obs,
    p = p_one_sided,
    ci_low = ci[1],
    ci_high = ci[2],
    null_mean = null_mean,
    null_sd = null_sd,
    z = z_score,
    B = B,
    null = null
  )
}

# ----------------------------
# PERMUTATION: OVERALL
# ----------------------------
events_t <- events$t_peak
point_times_all <- sp2$time_s

perm_overall <- perm_test_shift(
  events_t     = events_t,
  point_times  = point_times_all,
  T_total      = T_total,
  window_s     = window_s,
  B            = B_perm,
  seed         = seed_perm
)

perm_overall_df <- tibble(
  tag        = tag,
  window_s   = window_s,
  obs_Lift   = perm_overall$obs,
  p_value    = perm_overall$p,
  CI_low     = perm_overall$ci_low,
  CI_high    = perm_overall$ci_high,
  null_mean  = perm_overall$null_mean,
  null_sd    = perm_overall$null_sd,
  z_score    = perm_overall$z,
  B          = perm_overall$B
)

write_csv(perm_overall_df, out_file("perm_overall"))

# ----------------------------
# PERMUTATION: BY MACRO
# ----------------------------
macros_perm <- sort(unique(sp2$Macro))

perm_by_macro <- lapply(macros_perm, function(m) {
  pts_m <- sp2$time_s[sp2$Macro == m]
  
  res <- perm_test_shift(
    events_t    = events_t,
    point_times = pts_m,
    T_total     = T_total,
    window_s    = window_s,
    B           = B_perm,
    seed        = seed_perm
  )
  
  tibble(
    tag        = tag,
    Macro      = m,
    N_nodes    = length(pts_m),
    obs_Lift   = res$obs,
    p_value    = res$p,
    CI_low     = res$ci_low,
    CI_high    = res$ci_high,
    null_mean  = res$null_mean,
    null_sd    = res$null_sd,
    z_score    = res$z,
    B          = res$B
  )
}) %>%
  bind_rows() %>%
  arrange(p_value)

write_csv(perm_by_macro, out_file("perm_by_macro"))

# ----------------------------
# PRINT + DONE
# ----------------------------
print(summary_overall)
cat("\nSaved to:\n",
    "  ", out_file("summary_overall"), "\n",
    "  ", out_file("summary_by_macro"), "\n",
    "  ", out_file("macro_windows_union"), "\n",
    "  ", out_file("events_with_struct_flags"), "\n",
    "  ", out_file("perm_overall"), "\n",
    "  ", out_file("perm_by_macro"), "\n", sep = "")
