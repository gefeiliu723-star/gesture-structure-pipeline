## ============================================================
# Script 01 (CLEAN / GitHub-ready):
# Extract wrist_tracks_<TAG>.csv from segments/<TAG>*.mp4
#
# INPUT:
#   segments/<TAG>_*.mp4  (auto pick newest)  OR a relative path you provide
#   scripts/pose_landmarker_full.task (preferred) / assets/... / root ...
#
# OUTPUT:
#   exports/wrist_tracks/wrist_tracks_<TAG>.csv
#
# CLI:
#   Rscript scripts/01_extract_wrist_tracks.R <TAG> [video_rel_path_or_empty] [fps]
#
# ROOT:
#   GESTURE_PROJECT_ROOT env var; fallback getwd() (run from repo root)
#
# PYTHON:
#   Set ONE of:
#     Sys.setenv(RETICULATE_PYTHON=".../python")
#     Sys.setenv(MP_PYTHON=".../python")

# NOTE 01-A: 
#     We do not interpret individual landmark positions semantically.
#     All downstream analyses operate on derived kinematic magnitudes (speed), not pose meaning.
## ============================================================

suppressPackageStartupMessages({
  library(stringr)
  library(reticulate)
})

# ----------------------------
# (0) SETTINGS / CLI (TAG REQUIRED)
# ----------------------------
args <- commandArgs(trailingOnly = TRUE)

if (length(args) < 1 || !nzchar(args[[1]])) {
  stop(
    "Missing TAG.\n",
    "Usage:\n  Rscript scripts/01_extract_wrist_tracks.R <TAG> [video_rel_path_or_empty] [fps]\n",
    call. = FALSE
  )
}

tag <- args[[1]]
video_path_rel <- if (length(args) >= 2 && nzchar(args[[2]])) args[[2]] else ""
fps_user <- if (length(args) >= 3 && nzchar(args[[3]])) suppressWarnings(as.numeric(args[[3]])) else NA_real_

if (!grepl("^m[0-9]{2}$", tag)) {
  stop("TAG must look like mXX (e.g., m03). Got: ", tag, call. = FALSE)
}

# ----------------------------
# (0.1) PROJECT ROOT (portable)
# ----------------------------
proj_root <- Sys.getenv("GESTURE_PROJECT_ROOT", unset = getwd())

# ----------------------------
# (0.2) Python interpreter (portable)
# ----------------------------
PY <- Sys.getenv("RETICULATE_PYTHON", unset = "")
if (!nzchar(PY)) PY <- Sys.getenv("MP_PYTHON", unset = "")

if (nzchar(PY)) {
  Sys.setenv(RETICULATE_PYTHON = PY)
  reticulate::use_python(PY, required = TRUE)
}

# ----------------------------
# (0.3) Locate video (auto newest, or user-provided relative path)
# ----------------------------
pick_newest <- function(files) {
  files <- files[file.exists(files)]
  if (length(files) == 0) return(NA_character_)
  files[order(file.info(files)$mtime, decreasing = TRUE)][1]
}

if (!nzchar(video_path_rel)) {
  seg_dir <- file.path(proj_root, "segments")
  if (!dir.exists(seg_dir)) stop("Missing segments/ folder at: ", seg_dir, call. = FALSE)
  
  cand <- list.files(
    seg_dir,
    pattern = paste0("^", tag, "_.*\\.mp4$"),
    full.names = TRUE,
    ignore.case = TRUE
  )
  video_path <- pick_newest(cand)
  
  if (is.na(video_path)) {
    stop(
      "No video found for tag=", tag, " in segments/.\n",
      "Expected filename like: ", tag, "_....mp4",
      call. = FALSE
    )
  }
} else {
  # allow relative path under root
  if (grepl("^[A-Za-z]:[/\\\\]", video_path_rel) || startsWith(video_path_rel, "/")) {
    # absolute given (allowed, but not recommended)
    video_path <- video_path_rel
  } else {
    video_path <- file.path(proj_root, video_path_rel)
  }
}

# ----------------------------
# (0.4) Locate MediaPipe task file (portable)
# ----------------------------
find_task <- function(root) {
  candidates <- c(
    file.path(root, "scripts", "pose_landmarker_full.task"),
    file.path(root, "assets",  "pose_landmarker_full.task"),
    file.path(root, "pose_landmarker_full.task")
  )
  hit <- candidates[file.exists(candidates)]
  if (length(hit) == 0) {
    stop(
      "Cannot find pose_landmarker_full.task.\nTried:\n  - ",
      paste(candidates, collapse = "\n  - "),
      call. = FALSE
    )
  }
  hit[[1]]
}
task_path <- find_task(proj_root)

# ----------------------------
# (0.5) HARD CHECKS (DO NOT EDIT)
# ----------------------------
if (!file.exists(video_path)) stop("Missing video: ", video_path, call. = FALSE)

base_name <- basename(video_path)
tag_inferred <- sub("^((m[0-9]{2})).*$", "\\1", base_name)

if (!grepl("^m[0-9]{2}$", tag_inferred)) {
  stop(
    "Cannot infer tag from video filename.\n",
    "Expected filename starting with mXX_, got: ", base_name,
    call. = FALSE
  )
}

if (tag != tag_inferred) {
  stop(
    "TAG mismatch detected.\n",
    "  tag        = ", tag, "\n",
    "  video_path = ", video_path, "\n",
    "  inferred   = ", tag_inferred, "\n\n",
    "Fix: run with tag=", tag_inferred, " OR choose the correct video.",
    call. = FALSE
  )
}

cat("[OK] project_root:", proj_root, "\n")
cat("[OK] video       :", video_path, "\n")
cat("[OK] tag         :", tag, "\n")
cat("[OK] task        :", task_path, "\n")

# ----------------------------
# (1) OUTPUT PATH (repo-local)
# ----------------------------
export_dir <- file.path(proj_root, "exports", "wrist_tracks")
dir.create(export_dir, showWarnings = FALSE, recursive = TRUE)

out_csv <- file.path(export_dir, paste0("wrist_tracks_", tag, ".csv"))
cat("[WRITE]", out_csv, "\n")

# SAFETY: avoid PermissionError (file locked by Excel)
if (file.exists(out_csv)) {
  ok_rm <- tryCatch({ file.remove(out_csv) }, error = function(e) FALSE)
  if (!isTRUE(ok_rm)) stop("Cannot overwrite CSV (maybe open in Excel?): ", out_csv, call. = FALSE)
}

# ----------------------------
# (2) PYTHON SCRIPT (as string)
# ----------------------------
# NOTE: we do NOT force fps inside python; we read actual fps from video.
# If user supplies fps, we will only use it to compute t_sec in R? (not needed).
# We'll keep python as-is but leave user fps unused (cleaner).
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
min_box_area     = 0.01

smooth_alpha     = 0.7
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

print("VIDEO:", VIDEO_PATH)
print("FPS:", fps, "W,H:", w, h, "Frames:", n_frames)
print("OUT:", OUT_CSV)

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
print("DONE ->", OUT_CSV)
', video_path, out_csv, task_path)

# ----------------------------
# (3) RUN PYTHON
# ----------------------------
py_file <- tempfile(fileext = ".py")
writeLines(py_code, py_file, useBytes = TRUE)
reticulate::py_run_file(py_file)

cat("Finished. CSV at:", out_csv, "\n")
