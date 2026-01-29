# ============================================================
# Script 01 (FINAL / A VERSION: R one-click, embedded Python)
# Video -> YOLO teacher box -> MediaPipe Pose -> wrist_tracks_<TAG>.csv
#
# You run ONLY R:
#   Rscript scripts/01_make_wrist_tracks.R m01
#   Rscript scripts/01_make_wrist_tracks.R l07
#
# Python is used UNDER THE HOOD via reticulate.
# You DO NOT manually run Python scripts/commands.
#
# INPUT:
#   segments/<TAG>_*.mp4   (auto-pick newest if multiple)
#   assets/pose_landmarker_full.task
#
# OUTPUT:
#   exports/wrist_tracks/wrist_tracks_<TAG>.csv
#
# ROOT:
#   - env var GESTURE_PROJECT_ROOT (recommended)
#   - else fallback to getwd() (run from project root)
#
# PYTHON selection (optional but recommended):
#   Set env var RETICULATE_PYTHON to point to your python.exe
#   Example (Windows PowerShell):
#     $env:RETICULATE_PYTHON="C:\Users\you\miniconda3\envs\gesture\python.exe"
#
# NOTE:
#   This script does NOT interpret gesture semantics.
#   It extracts kinematic signals (wrist trajectories) only.
# ============================================================

suppressPackageStartupMessages({
  library(reticulate)
  library(stringr)
})

# ----------------------------
# (0) TAG from CLI or manual
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

TAG <- if (length(args) >= 1 && nzchar(args[[1]])) args[[1]] else "m01"  # fallback
TAG <- tolower(TAG)

if (!grepl("^[ml][0-9]{2}$", TAG)) {
  stop("TAG must look like mXX or lXX (e.g., m03, l07). Got: ", TAG, call. = FALSE)
}

# ----------------------------
# (0.1) Project root
# ----------------------------
project_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

segments_dir <- file.path(project_root, "segments")
assets_dir   <- file.path(project_root, "assets")
exports_dir  <- file.path(project_root, "exports", "wrist_tracks")

if (!dir.exists(segments_dir)) stop("Missing segments/ folder: ", segments_dir, call. = FALSE)
if (!dir.exists(assets_dir))   stop("Missing assets/ folder: ", assets_dir, call. = FALSE)
dir.create(exports_dir, showWarnings = FALSE, recursive = TRUE)

task_path <- file.path(assets_dir, "pose_landmarker_full.task")
if (!file.exists(task_path)) {
  stop("Missing MediaPipe task model: ", task_path, "\n",
       "Expected: assets/pose_landmarker_full.task", call. = FALSE)
}

# ----------------------------
# (0.2) Find exactly-one (or newest) video
# ----------------------------
cand <- list.files(
  segments_dir,
  pattern = paste0("^", TAG, "_.*\\.mp4$"),
  full.names = TRUE,
  ignore.case = TRUE
)

if (length(cand) == 0) {
  stop(
    "No video found for TAG=", TAG, " in segments/.\n",
    "Expected filename like: ", TAG, "_....mp4\n",
    "segments_dir=", segments_dir,
    call. = FALSE
  )
}

# If multiple, pick newest by mtime
if (length(cand) > 1) {
  cand <- cand[order(file.info(cand)$mtime, decreasing = TRUE)]
  message("[INFO] Multiple candidates found; picking newest:\n  ", paste(cand, collapse = "\n  "))
}
video_path <- cand[[1]]

# extra safety: infer tag from filename prefix
base_name <- basename(video_path)
tag_inferred <- sub("^(([ml][0-9]{2})).*$", "\\1", base_name)
if (!identical(TAG, tag_inferred)) {
  stop(
    "TAG mismatch:\n",
    "  TAG          = ", TAG, "\n",
    "  video_path   = ", video_path, "\n",
    "  inferred TAG = ", tag_inferred, "\n",
    "Fix: rename file to start with TAG_..., or run with the inferred TAG.",
    call. = FALSE
  )
}

out_csv <- file.path(exports_dir, paste0("wrist_tracks_", TAG, ".csv"))

cat("[OK] project_root:", project_root, "\n")
cat("[OK] TAG         :", TAG, "\n")
cat("[OK] video       :", video_path, "\n")
cat("[OK] task        :", task_path, "\n")
cat("[OUT] csv        :", out_csv, "\n")

# avoid PermissionError if open in Excel
if (file.exists(out_csv)) {
  ok_rm <- tryCatch(file.remove(out_csv), error = function(e) FALSE)
  if (!isTRUE(ok_rm)) stop("Cannot overwrite CSV (is it open?): ", out_csv, call. = FALSE)
}

# ----------------------------
# (0.3) Python interpreter setup (reticulate)
# ----------------------------
# If user set RETICULATE_PYTHON, reticulate will honor it.
# Otherwise it uses default python on PATH / configured env.
# You still only run R; this is just interpreter choice.
py_cfg <- tryCatch(py_config(), error = function(e) NULL)
if (is.null(py_cfg)) {
  stop("reticulate cannot find a working Python.\n",
       "Set RETICULATE_PYTHON to your python.exe and try again.", call. = FALSE)
} else {
  cat("[PY] python:", py_cfg$python, "\n")
}

# ----------------------------
# (1) Embedded Python code (LEGACY/A version)
# ----------------------------
# This is the "R one-click" version: Python is embedded and run via reticulate.
# You do NOT run Python manually.
py_code <- sprintf('
import os, csv
import cv2
import numpy as np
from ultralytics import YOLO

import mediapipe as mp
from mediapipe.tasks import python as mp_python
from mediapipe.tasks.python import vision as mp_vision

VIDEO_PATH = r"%s"
OUT_CSV    = r"%s"
TASK_PATH  = r"%s"

yolo_model_name  = "yolov8n.pt"
person_class_id  = 0

det_conf         = 0.25
iou_thres        = 0.45
min_box_area     = 0.01   # fraction of frame area

smooth_alpha     = 0.70
iou_keep_thres   = 0.15

idx = {"ls": 11, "rs": 12, "le": 13, "re": 14, "lw": 15, "rw": 16}

def box_iou(a, b):
    ax1, ay1, ax2, ay2 = a
    bx1, by1, bx2, by2 = b
    inter_x1 = max(ax1, bx1)
    inter_y1 = max(ay1, by1)
    inter_x2 = min(ax2, bx2)
    inter_y2 = min(ay2, by2)
    iw = max(0.0, inter_x2 - inter_x1)
    ih = max(0.0, inter_y2 - inter_y1)
    inter = iw * ih
    area_a = max(0.0, ax2-ax1) * max(0.0, ay2-ay1)
    area_b = max(0.0, bx2-bx1) * max(0.0, by2-by1)
    union = area_a + area_b - inter + 1e-9
    return inter / union

def smooth_box(prev, cur, alpha=0.7):
    if prev is None:
        return cur
    return tuple(alpha*np.array(prev) + (1-alpha)*np.array(cur))

def clamp_box(box, w, h):
    x1,y1,x2,y2 = box
    x1 = max(0, min(w-1, x1))
    y1 = max(0, min(h-1, y1))
    x2 = max(0, min(w-1, x2))
    y2 = max(0, min(h-1, y2))
    if x2 < x1: x1,x2 = x2,x1
    if y2 < y1: y1,y2 = y2,y1
    return (x1,y1,x2,y2)

def empty_landmarks():
    return [""] * (6*4)

if not os.path.exists(VIDEO_PATH):
    raise RuntimeError(f"Video not found: {VIDEO_PATH}")
if not os.path.exists(TASK_PATH):
    raise RuntimeError(f"Task model not found: {TASK_PATH}")

yolo = YOLO(yolo_model_name)

base_options = mp_python.BaseOptions(model_asset_path=TASK_PATH)
options = mp_vision.PoseLandmarkerOptions(
    base_options=base_options,
    running_mode=mp_vision.RunningMode.VIDEO,
    output_segmentation_masks=False
)
pose = mp_vision.PoseLandmarker.create_from_options(options)

cap = cv2.VideoCapture(VIDEO_PATH)
if not cap.isOpened():
    raise RuntimeError(f"Cannot open video: {VIDEO_PATH}")

fps = float(cap.get(cv2.CAP_PROP_FPS) or 0.0)
w   = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
h   = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)
n_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)

print("[PY] VIDEO:", VIDEO_PATH)
print("[PY] FPS:", fps, "W,H:", w, h, "Frames:", n_frames)
print("[PY] OUT:", OUT_CSV)

header = [
    "frame","t_sec","teacher_id",
    "conf_det","x1","y1","x2","y2",
    "frame_w","frame_h","fps","n_frames"
]
for k in ["ls","rs","le","re","lw","rw"]:
    header += [f"{k}_x", f"{k}_y", f"{k}_z", f"{k}_vis"]

os.makedirs(os.path.dirname(OUT_CSV), exist_ok=True)
with open(OUT_CSV, "w", newline="", encoding="utf-8") as fout:
    writer = csv.writer(fout)
    writer.writerow(header)

    prev_teacher_box = None
    teacher_id = 1
    frame_i = 0

    while True:
        ok, frame = cap.read()
        if not ok:
            break

        res = yolo.predict(source=frame, conf=det_conf, iou=iou_thres, verbose=False)[0]
        boxes, confs = [], []

        if res.boxes is not None and len(res.boxes) > 0:
            for b in res.boxes:
                cls = int(b.cls.item())
                if cls != person_class_id:
                    continue
                conf = float(b.conf.item())
                x1,y1,x2,y2 = b.xyxy[0].tolist()
                area = ((x2-x1)*(y2-y1)) / (w*h + 1e-9)
                if area < min_box_area:
                    continue
                boxes.append((x1,y1,x2,y2))
                confs.append(conf)

        t_sec = frame_i / (fps + 1e-9)

        if len(boxes) == 0:
            row = [frame_i, t_sec, teacher_id, 0.0, "", "", "", "", w, h, fps, n_frames] + empty_landmarks()
            writer.writerow(row)
            frame_i += 1
            continue

        if prev_teacher_box is None:
            areas = [max(0,(b[2]-b[0]))*max(0,(b[3]-b[1])) for b in boxes]
            j = int(np.argmax(areas))
        else:
            ious = [box_iou(prev_teacher_box, b) for b in boxes]
            j = int(np.argmax(ious))
            if ious[j] < iou_keep_thres:
                areas = [max(0,(b[2]-b[0]))*max(0,(b[3]-b[1])) for b in boxes]
                j = int(np.argmax(areas))

        chosen_box  = clamp_box(boxes[j], w, h)
        chosen_conf = confs[j]
        chosen_box  = smooth_box(prev_teacher_box, chosen_box, alpha=smooth_alpha)
        prev_teacher_box = chosen_box

        x1,y1,x2,y2 = map(int, chosen_box)
        crop = frame[y1:y2, x1:x2].copy()

        if crop.size == 0:
            row = [frame_i, t_sec, teacher_id, chosen_conf, x1,y1,x2,y2, w, h, fps, n_frames] + empty_landmarks()
            writer.writerow(row)
            frame_i += 1
            continue

        crop_rgb = cv2.cvtColor(crop, cv2.COLOR_BGR2RGB)
        mp_image = mp.Image(image_format=mp.ImageFormat.SRGB, data=crop_rgb)

        ts_ms = int(t_sec * 1000)
        result = pose.detect_for_video(mp_image, ts_ms)

        vals = {}
        if result.pose_landmarks and len(result.pose_landmarks) > 0:
            lms = result.pose_landmarks[0]
            for k, li in idx.items():
                lm = lms[li]
                fx = (x1 + lm.x*(x2-x1)) / (w + 1e-9)
                fy = (y1 + lm.y*(y2-y1)) / (h + 1e-9)
                fz = lm.z
                vals[k] = (fx, fy, fz, lm.visibility)
        else:
            for k in idx.keys():
                vals[k] = ("","","","")

        row = [frame_i, t_sec, teacher_id, chosen_conf, x1,y1,x2,y2, w, h, fps, n_frames]
        for k in ["ls","rs","le","re","lw","rw"]:
            row += list(vals.get(k, ("","","","")))
        writer.writerow(row)

        frame_i += 1

cap.release()
print("[PY] DONE ->", OUT_CSV)
', normalizePath(video_path, winslash = "/", mustWork = TRUE),
normalizePath(out_csv,    winslash = "/", mustWork = FALSE),
normalizePath(task_path,  winslash = "/", mustWork = TRUE)
)

# ----------------------------
# (2) Run embedded python via reticulate
# ----------------------------
py_file <- tempfile(fileext = ".py")
writeLines(py_code, py_file, useBytes = TRUE)

cat("[RUN] embedded python temp:", py_file, "\n")
reticulate::py_run_file(py_file)

if (!file.exists(out_csv)) {
  stop("Python finished but CSV missing: ", out_csv, call. = FALSE)
}

cat("[DONE] CSV written:", out_csv, "\n")
