# ============================================================
# Script 05: Analysis (Node ±window) + Macro breakdown + Permutation
# GitHub-ready (portable paths)
#
# INPUT : exports/<tag>/<tag>_alignment.xlsx   (ONLY THIS)
# OUTPUT: results/<tag>/  with <tag>_ prefix
#
# DEFAULT: NO KEEP FILTERS
# OPTIONAL: KEEP/DROP + Confidence gating from Alignment sheet
#
# Usage:
#   Rscript scripts/05_analysis_alignment.R
#   Rscript scripts/05_analysis_alignment.R m07 2
#   Rscript scripts/05_analysis_alignment.R m07 2 2000 42
#
# Optional env var:
#   GESTURE_PROJECT_ROOT=/path/to/repo_root
# ============================================================

suppressPackageStartupMessages({
  library(readr)
  library(readxl)
  library(dplyr)
  library(stringr)
  library(tibble)
})

# ----------------------------
# ONLY EDIT THIS (or pass via CLI)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

tag      <- if (length(args) >= 1) args[[1]] else "m07"
window_s <- if (length(args) >= 2) as.numeric(args[[2]]) else 2

# Permutation settings
B_perm    <- if (length(args) >= 3) as.integer(args[[3]]) else 2000L
seed_perm <- if (length(args) >= 4) as.integer(args[[4]]) else 42L

# ----------------------------
# OPTIONAL: KEEP/DROP + Confidence gating
# ----------------------------
use_keep_filter <- FALSE   # <- set TRUE to enable gating
min_confidence  <- 1       # <- keep only if CONF >= this (1/2/3)
default_keep    <- 1       # <- if KEEP is blank/NA, treat as keep(1) by default
default_conf    <- 3       # <- if Confidence blank/NA, treat as 3 by default

# ----------------------------
# PORTABLE PROJECT PATHS
# ----------------------------
proj_root  <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

export_dir <- file.path(proj_root, "exports", tag)
result_dir <- file.path(proj_root, "results", tag)
dir.create(result_dir, showWarnings = FALSE, recursive = TRUE)

xlsx_path  <- file.path(export_dir, paste0(tag, "_alignment.xlsx"))

# output helper
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
    if (!is.finite(t)) next
    
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
  stop(
    "Missing alignment workbook: ", xlsx_path, "\n",
    "Expected: exports/<tag>/<tag>_alignment.xlsx\n",
    "Tip: run from repo root OR set env var GESTURE_PROJECT_ROOT."
  )
}

# ----------------------------
# LOAD EVENTS (FROM XLSX)
# ----------------------------
events <- read_excel(xlsx_path, sheet = "GestureEvents") %>% as_tibble()

events <- events %>%
  rename_with(~"start_sec", matches("^start_?sec$|^event_start_?sec$|^event_start_?s$|^start$", ignore.case = TRUE)) %>%
  rename_with(~"end_sec",   matches("^end_?sec$|^event_end_?sec$|^event_end_?s$|^end$", ignore.case = TRUE))

if (!all(c("start_sec", "end_sec") %in% names(events))) {
  stop("GestureEvents must contain start/end columns (e.g., start_sec & end_sec).")
}

# Build event_id robustly
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
  )

# HARDEN: drop NA/invalid events
events <- events %>%
  filter(is.finite(start_sec), is.finite(end_sec), end_sec >= start_sec) %>%
  mutate(t_peak = (start_sec + end_sec) / 2) %>%
  filter(is.finite(t_peak))

if (nrow(events) == 0) stop("No valid events after cleaning (start_sec/end_sec all NA or invalid).")

T_total <- suppressWarnings(max(events$end_sec, na.rm = TRUE))
if (!is.finite(T_total) || is.na(T_total) || T_total <= 0) stop("Bad T_total inferred from events end_sec.")

message(sprintf("Loaded events: n=%d | T_total=%.3f sec", nrow(events), T_total))

# ============================================================
# OPTIONAL: KEEP/DROP + Confidence gating (from Alignment sheet)
# ============================================================
audit_pointlevel <- NULL
audit_eventlevel <- NULL
keep_event_ids <- NULL
drop_event_ids <- NULL

if (use_keep_filter) {
  
  aln <- read_excel(xlsx_path, sheet = "Alignment") %>% as_tibble()
  
  names(aln) <- names(aln) %>%
    str_replace_all("\\.+", "_") %>%
    str_replace_all("[^A-Za-z0-9_]", "_") %>%
    str_replace_all("_+", "_") %>%
    tolower()
  
  event_cols <- names(aln)[str_detect(names(aln), regex("^eventid($|_)|eventid_auto|eventid_manual", ignore_case = TRUE))]
  if (length(event_cols) == 0) stop("Alignment sheet: could not find an EventID column (manual/auto).")
  
  pick_event_col <- function(cols) {
    if (any(cols == "eventid_manual")) return("eventid_manual")
    if (any(str_detect(cols, "manual"))) return(cols[str_detect(cols, "manual")][1])
    if (any(cols == "eventid_auto")) return("eventid_auto")
    if (any(str_detect(cols, "auto"))) return(cols[str_detect(cols, "auto")][1])
    cols[1]
  }
  event_col <- pick_event_col(event_cols)
  
  keep_cols <- names(aln)[str_detect(names(aln), regex("^keep($|_)|keep_drop", ignore_case = TRUE))]
  if (length(keep_cols) == 0) stop("Alignment sheet: missing KEEP/DROP column.")
  keep_col <- keep_cols[1]
  
  conf_cols <- names(aln)[str_detect(names(aln), regex("^confidence($|_)", ignore_case = TRUE))]
  conf_col <- if (length(conf_cols) == 0) NA_character_ else conf_cols[1]
  
  reason_cols <- names(aln)[str_detect(names(aln), regex("^reason($|_)", ignore_case = TRUE))]
  reason_col <- if (length(reason_cols) == 0) NA_character_ else reason_cols[1]
  
  pid_cols <- names(aln)[str_detect(names(aln), regex("^pointid($|_)", ignore_case = TRUE))]
  pid_col <- if (length(pid_cols) == 0) NA_character_ else pid_cols[1]
  
  t_cols <- names(aln)[str_detect(names(aln), regex("^t_struct_sec$|^time_s$|^time_sec$|^t_sec$|^time$", ignore_case = TRUE))]
  t_col <- if (length(t_cols) == 0) NA_character_ else t_cols[1]
  
  audit_pointlevel <- aln %>%
    transmute(
      PointID = if (!is.na(pid_col)) as.character(.data[[pid_col]]) else as.character(row_number()),
      t_struct_sec = if (!is.na(t_col)) suppressWarnings(as.numeric(.data[[t_col]])) else NA_real_,
      EventID_alignment = as.character(.data[[event_col]]),
      KEEP_raw = suppressWarnings(as.numeric(.data[[keep_col]])),
      Confidence_raw = if (!is.na(conf_col)) suppressWarnings(as.numeric(.data[[conf_col]])) else NA_real_,
      Reason = if (!is.na(reason_col)) as.character(.data[[reason_col]]) else NA_character_
    ) %>%
    mutate(
      KEEP = ifelse(is.na(KEEP_raw), default_keep, KEEP_raw),
      Confidence = ifelse(is.na(Confidence_raw), default_conf, Confidence_raw),
      decision = case_when(
        is.na(EventID_alignment) | EventID_alignment == "" ~ "NO_EVENTID",
        KEEP != 1 ~ "DROP_KEEP0",
        Confidence < min_confidence ~ "DROP_LOWCONF",
        TRUE ~ "KEEP"
      )
    )
  
  audit_eventlevel <- audit_pointlevel %>%
    filter(!is.na(EventID_alignment), EventID_alignment != "") %>%
    group_by(EventID_alignment) %>%
    summarise(
      n_rows = n(),
      n_keep_rows = sum(decision == "KEEP", na.rm = TRUE),
      n_drop_keep0 = sum(decision == "DROP_KEEP0", na.rm = TRUE),
      n_drop_lowconf = sum(decision == "DROP_LOWCONF", na.rm = TRUE),
      min_conf = suppressWarnings(min(Confidence, na.rm = TRUE)),
      max_conf = suppressWarnings(max(Confidence, na.rm = TRUE)),
      reasons = paste(unique(na.omit(Reason)), collapse = " | "),
      .groups = "drop"
    ) %>%
    mutate(kept_event = n_keep_rows > 0) %>%
    arrange(desc(kept_event), desc(n_keep_rows), EventID_alignment)
  
  keep_event_ids <- audit_eventlevel %>% filter(kept_event) %>% pull(EventID_alignment) %>% unique()
  drop_event_ids <- audit_eventlevel %>% filter(!kept_event) %>% pull(EventID_alignment) %>% unique()
  
  write_csv(audit_pointlevel, out_file("keep_audit_pointlevel"))
  write_csv(audit_eventlevel, out_file("keep_audit_eventlevel"))
  write_csv(tibble(event_id = keep_event_ids), out_file("kept_event_ids"))
  write_csv(tibble(event_id = drop_event_ids), out_file("dropped_event_ids"))
  
  events_before <- nrow(events)
  events <- events %>% filter(event_id %in% keep_event_ids)
  
  if (nrow(events) == 0) {
    stop("After KEEP/CONF filtering, no events remain. Check Alignment sheet EventID values.")
  }
  
  message(sprintf("[KEEP FILTER ON] events: %d -> %d kept (min_conf=%d)",
                  events_before, nrow(events), min_confidence))
}

# ----------------------------
# LOAD STRUCTURAL POINTS (FROM XLSX)
# ----------------------------
sp <- read_excel(xlsx_path, sheet = "StructuralPoints") %>% as_tibble()

time_col <- names(sp)[str_detect(names(sp), regex("^time_s$|^time_sec$|^t_sec$|^time$", ignore_case = TRUE))][1]
if (is.na(time_col) || !nzchar(time_col)) stop("StructuralPoints must contain time_s (or time_sec/t_sec/time).")
sp <- sp %>% rename(time_s = all_of(time_col))

macro_col <- names(sp)[str_detect(names(sp), regex("^macro$", ignore_case = TRUE))][1]
if (is.na(macro_col) || !nzchar(macro_col)) {
  sp <- sp %>% mutate(Macro = NA_character_)
} else {
  sp <- sp %>% rename(Macro = all_of(macro_col)) %>% mutate(Macro = as.character(.data$Macro))
}

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
  filter(is.finite(time_s))

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

events <- events %>%
  mutate(struct_overall = point_in_intervals(t_peak, overall_union))

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
  Lift = Lift,
  keep_filter = use_keep_filter,
  min_confidence = ifelse(use_keep_filter, min_confidence, NA_integer_)
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
    Lift = Lift_m,
    keep_filter = use_keep_filter
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
  B          = perm_overall$B,
  keep_filter = use_keep_filter,
  min_confidence = ifelse(use_keep_filter, min_confidence, NA_integer_)
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
    B          = res$B,
    keep_filter = use_keep_filter,
    min_confidence = ifelse(use_keep_filter, min_confidence, NA_integer_)
  )
}) %>%
  bind_rows() %>%
  arrange(p_value)

write_csv(perm_by_macro, out_file("perm_by_macro"))

# ----------------------------
# PRINT
# ----------------------------
print(summary_overall)
cat("\nSaved to:\n",
    "  ", out_file("summary_overall"), "\n",
    "  ", out_file("summary_by_macro"), "\n",
    "  ", out_file("macro_windows_union"), "\n",
    "  ", out_file("events_with_struct_flags"), "\n",
    "  ", out_file("perm_overall"), "\n",
    "  ", out_file("perm_by_macro"), "\n", sep = "")

if (use_keep_filter) {
  cat("\nKEEP FILTER AUDIT:\n",
      "  ", out_file("keep_audit_pointlevel"), "\n",
      "  ", out_file("keep_audit_eventlevel"), "\n",
      "  ", out_file("kept_event_ids"), "\n",
      "  ", out_file("dropped_event_ids"), "\n", sep = "")
}
