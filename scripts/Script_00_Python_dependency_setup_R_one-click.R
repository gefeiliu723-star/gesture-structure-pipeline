# ============================================================
# Script 00: Python dependency setup (R one-click)
#
# You run ONLY R:
#   Rscript scripts/00_setup_python.R
#
# This script:
#   - checks that reticulate can find Python
#   - checks required Python packages
#   - installs missing ones via pip (called from R)
#
# You do NOT manually run any Python or pip commands.
# ============================================================

suppressPackageStartupMessages({
  library(reticulate)
})

# ----------------------------
# (0) Choose Python interpreter
# ----------------------------
# Recommended: user sets RETICULATE_PYTHON beforehand
# e.g. Sys.setenv(RETICULATE_PYTHON="C:/.../python.exe")

py_cfg <- tryCatch(py_config(), error = function(e) NULL)
if (is.null(py_cfg)) {
  stop(
    "reticulate cannot find a working Python.\n",
    "Please set RETICULATE_PYTHON to a valid python.exe and rerun.\n",
    call. = FALSE
  )
}

cat("[OK] Using Python:\n  ", py_cfg$python, "\n\n")

# ----------------------------
# (1) Required Python packages
# ----------------------------
required_pkgs <- c(
  "numpy",
  "opencv-python",
  "ultralytics",
  "mediapipe"
)

# ----------------------------
# (2) Check installed packages
# ----------------------------
installed <- py_list_packages()
installed_names <- installed$package

missing <- setdiff(required_pkgs, installed_names)

if (length(missing) == 0) {
  cat("[OK] All required Python packages already installed.\n")
  quit(save = "no", status = 0)
}

cat("[INFO] Missing Python packages:\n  - ",
    paste(missing, collapse = "\n  - "), "\n\n")

# ----------------------------
# (3) Install missing packages via pip (from R)
# ----------------------------
cat("[INSTALL] Installing missing packages via pip...\n\n")

# Use reticulate helper (calls pip internally)
for (pkg in missing) {
  cat("  -> installing:", pkg, "\n")
  py_install(pkg, pip = TRUE)
}

cat("\n[DONE] Python dependencies installed successfully.\n")

# ----------------------------
# (4) Sanity check imports
# ----------------------------
cat("\n[CHECK] Testing imports...\n")

py_run_string("
import numpy
import cv2
import ultralytics
import mediapipe
print('All imports successful.')
")

cat("\n[OK] Python environment is ready for Script 01.\n")
