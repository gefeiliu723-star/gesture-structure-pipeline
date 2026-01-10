# Quantitative Gesture–Discourse Alignment Pipeline

Scripts:
- 01_extract_wrist_tracks.R: YOLO + MediaPipe -> wrist_tracks_<tag>.csv
- 02_define_gestures.R: event detection + robustness outputs
- 03_teacher_only_watch_clips.R: teacher-present clip selection + ffmpeg batch
- 04_build_alignment_xlsx.R: integrate outputs into alignment workbook
- 05_analysis_alignment.R: enrichment + macro breakdown + permutation test
