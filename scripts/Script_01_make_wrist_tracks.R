# ============================================================
# Script 1 (PUBLIC / FINAL):
# Segment mp4 -> YOLO teacher box -> MediaPipe Pose -> CSV
#
# Output (default):
#   <proj_root>/exports/wrist_tracks/wrist_tracks_<tag>.csv
# Note:
# - This repo does NOT include any videos/URLs. Provide your own local MP4.
# - You may override paths via environment variables (see ENV below).
# ============================================================

suppressPackageStartupMessages({
  library(reticulate)
})

# ----------------------------
# (0) USER SETTINGS (PUBLIC TEMPLATE)
# ----------------------------
# IMPORTANT:
# - Do NOT commit real local paths or YouTube IDs.
# - Configure via environment variables OR edit placeholders below.

# Option A (recommended): set env vars (Windows: System Env / .Renviron)
#   GESTURE_VIDEO : full local path to an .mp4 file (or a downloaded segment)
#   GESTURE_TAG   : a short id like "m01" / "demo01" (do not use YouTube IDs)
#   GESTURE_PY    : path to python.exe (with ultralytics + mediapipe + opencv installed)
#   GESTURE_TASK  : path to pose_landmarker_full.task
#   GESTURE_ROOT  : project root (the folder that contains exports/)

get_env_or <- function(key, default) {
  v <- Sys.getenv(key, unset = "")
  if (nzchar(v)) v else default
}

# Public-safe placeholders (OK to commit)
video_path <- get_env_or("GESTURE_VIDEO", "<PATH_TO_LOCAL_VIDEO>.mp4")
tag        <- get_env_or("GESTURE_TAG",   "mXX")

PY         <- get_env_or("GESTURE_PY",    "<PATH_TO_PYTHON_EXECUTABLE>")
task_path  <- get_env_or("GESTURE_TASK",  "<PATH_TO_MEDIAPIPE_TASK>.task")
proj_root  <- get_env_or("GESTURE_ROOT",  "<PATH_TO_PROJECT_ROOT>")

# Reticulate uses this python
Sys.setenv(RETICULATE_PYTHON = PY)


# ----------------------------
# (1) OUTPUT PATH (auto)
# ----------------------------
export_dir <- file.path(proj_root, "exports", "wrist_tracks")
dir.create(export_dir, showWarnings = FALSE, recursive = TRUE)
out_csv <- file.path(export_dir, paste0("wrist_tracks_", tag, ".csv"))

# SAFETY: avoid PermissionError (file locked by Excel)
if (file.exists(out_csv)) {
  ok_rm <- tryCatch({ file.remove(out_csv) }, error = function(e) FALSE)
  if (!isTRUE(ok_rm)) stop("[Script1] Cannot overwrite CSV (maybe open in Excel?):\n  ", out_csv, call. = FALSE)
}

# ----------------------------
# (2) RETICULATE SETUP
# ----------------------------
Sys.setenv(RETICULATE_PYTHON = PY)
reticulate::use_python(PY, required = TRUE)

# ----------------------------
# (3) PYTHON SCRIPT (as a string)
# ----------------------------
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

# ----------------------------
# CONFIG (stable defaults)
# ----------------------------
yolo_model_name  = "yolov8n.pt"
person_class_id  = 0

det_conf         = 0.25
iou_thres        = 0.45
min_box_area     = 0.01

smooth_alpha     = 0.7
iou_keep_thres   = 0.15

# joints (MediaPipe PoseLandmarker indices)
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

# ----------------------------
# INIT MODELS
# ----------------------------
yolo = YOLO(yolo_model_name)

base_options = mp_python.BaseOptions(model_asset_path=TASK_PATH)
options = mp_vision.PoseLandmarkerOptions(
    base_options=base_options,
    running_mode=mp_vision.RunningMode.VIDEO,
    output_segmentation_masks=False
)
pose = mp_vision.PoseLandmarker.create_from_options(options)

# ----------------------------
# OPEN VIDEO
# ----------------------------
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

# ----------------------------
# CSV HEADER (final stable schema)
# ----------------------------
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
# (4) RUN PYTHON
# ----------------------------
py_file <- tempfile(fileext = ".py")
writeLines(py_code, py_file, useBytes = TRUE)
reticulate::py_run_file(py_file)

cat("✅ Finished. CSV at:\n", out_csv, "\n")

