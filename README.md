# Quantitative Gesture–Discourse Alignment Pipeline  
**Multi-Modal Kinematic Alignment Pipeline (MKAP)**

---

## Overview

This repository provides an automated, reproducible pipeline for examining the **temporal alignment between teachers’ bodily kinematics and instructional discourse structure**. By integrating computer vision–based pose tracking with signal-processing and statistical modeling, MKAP enables the analysis of **continuous movement dynamics** without reliance on categorical gesture labeling.

The pipeline is designed as a **methods-companion codebase** supporting quantitative investigations of how embodied behavior is temporally organized around discourse-structural events in naturalistic instruction.

---

## Pipeline Summary


MKAP proceeds through four conceptual stages:

1. **Video-based kinematic extraction**
2. **Continuous signal transformation**
3. **Event detection**
4. **Statistical alignment with discourse structure**

Each stage is modular, deterministic, and independently inspectable.

---

## Core Capabilities

### Computer Vision Front-End

Instructors are automatically localized using YOLO-based person detection. Upper-limb pose landmarks (shoulders, elbows, wrists) are then extracted using the MediaPipe Pose framework. The system operates on raw classroom video and produces frame-level kinematic trajectories.

### Kinematic Signal Construction

Raw landmark positions are transformed into continuous kinematic signals (e.g., velocity and acceleration). These signals are treated as **purely physical time series**, remaining agnostic to communicative intent or gesture semantics.

### Event Detection

Salient gesture events are identified using quantile-based thresholding and temporal clustering of kinematic peaks. Event detection is fully automated and parameterized, enabling systematic sensitivity analyses.

### Alignment and Enrichment Analysis

Gesture events are aligned to independently annotated discourse-structural points. Enrichment is quantified using rate-normalized lift statistics and permutation-based null models, allowing assessment of whether gesture activity is temporally concentrated near structural boundaries beyond chance expectations.

### Analytic Orchestration

The pipeline employs an R-centric workflow, with Python-based computer vision modules integrated via `reticulate`. Intermediate outputs are serialized as standardized tables and diagnostic workbooks, enabling transparent downstream statistical analysis and visualization in R.

---

## Robustness Analysis: Annotation Confidence

Confidence-based filtering is used **only for robustness analyses**, not to define discourse structure.

Discourse-structural nodes (e.g., pedagogical transitions or conceptual shifts) are identified via manual annotation of instructional flow and assigned a confidence score (1–3).

Primary analyses include **all** annotated structural points.  
Robustness checks repeat analyses using only high-confidence points (e.g., confidence = 3).

This procedure assesses sensitivity to annotation reliability rather than modifying alignment rules or event definitions.

---

## Data Governance and Privacy

### Zero-Data Policy

This repository contains **code only**.  
No video files, identifiable human-subject data, or copyrighted instructional materials are included.

### Local Execution

All processing is performed locally on user-provided data.  
No data are transmitted to external servers.

### Schema Compatibility

The pipeline interfaces with standard annotation formats (e.g., CSV, ELAN-style time markers) while remaining agnostic to annotation semantics.

---

## Reproducibility and Portability

MKAP is designed in accordance with open-science and reproducible-research principles.

### Deterministic Execution

Given identical kinematic inputs and hyperparameters (e.g., window sizes, quantile thresholds), the pipeline yields deterministic outputs.

### Configurable Environments

Execution paths, thresholds, and dataset locations are configurable via script-level parameters or environment variables.

An empty Excel alignment template defining data structure and formulas is provided in `templates/`.  
No annotations or results are embedded in the repository.

---

## Dependencies

### Python
- Python 3.9+
- `ultralytics` (YOLOv8)
- `mediapipe`
- `opencv-python`
- `numpy`
- `pandas`

YOLOv8 pretrained weights (e.g., `yolov8n.pt`) are automatically downloaded at runtime by the Ultralytics package and are not included in this repository.

### R
- R 4.2+
- `tidyverse`
- `openxlsx`
- `reticulate`

---

## Installation and Execution

This repository is intended for research use rather than turnkey deployment.

1. Install Python and required packages (e.g., via `pip` or `conda`)
2. Install R and required packages
3. Configure local paths or environment variables for user-provided datasets
4. Execute scripts sequentially following the pipeline order documented in `scripts/`

Downstream analyses require a single alignment workbook located at:


Results are written to tag-specific subdirectories under `results/`, which are created automatically if absent.

---

## Architecture

MKAP adopts an **R-centric architecture** with Python-based computer vision modules integrated through `reticulate`.

This design enables a fully R-native analytical workflow while leveraging mature Python libraries for low-level video processing. Users do not manually execute Python code; all processing is orchestrated from R.

---

## Ethics and Compliance

Users are responsible for ensuring that application of this pipeline complies with institutional IRB protocols, GDPR / FERPA regulations, and informed-consent requirements governing source materials.

---

## License

This project is released under the MIT License.

---

## What This Pipeline Does *Not* Claim

MKAP is designed as a quantitative methods framework for analyzing the **temporal organization** of bodily kinematics relative to discourse structure. Accordingly, the pipeline makes no claims beyond what is directly supported by its measurements and analytic scope.

Specifically, MKAP does **not** claim:

- **Semantic interpretation of gesture.**  
  The pipeline does not classify gestures by type, meaning, or communicative intent. All analyses operate on continuous kinematic signals derived from pose trajectories, independent of gesture semantics.

- **Causal influence between gesture and discourse.**  
  Temporal alignment or enrichment near discourse-structural boundaries is not interpreted as evidence that gesture causes, initiates, or determines discourse structure (or vice versa). The pipeline quantifies timing relationships only.

- **Cognitive state inference.**  
  Measures produced by MKAP are not direct indicators of attention, understanding, intention, or learning. Any cognitive interpretations must be theoretically motivated and supported by independent evidence.

- **Generality beyond the analyzed context.**  
  Findings obtained using this pipeline are contingent on the instructional settings, populations, and discourse annotation schemes to which it is applied. MKAP does not assume universality across domains or interaction types.

- **Annotation objectivity.**  
  Discourse-structural annotations are treated as externally supplied analytic inputs. While robustness analyses assess sensitivity to annotation confidence, the pipeline does not claim to discover or validate discourse structure autonomously.

By explicitly separating **measurement**, **alignment**, and **interpretation**, MKAP is intended to support principled empirical investigation while leaving theoretical claims to be developed and evaluated at the level of individual studies.

---

## Citation

If you use this pipeline in academic work, please cite it as a methods-companion codebase:

**Liu, G. (2024).**  
*Multi-Modal Kinematic Alignment Pipeline: A Computational Framework for Instructional Discourse.*  
GitHub repository.
