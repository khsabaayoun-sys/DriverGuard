# DriverGuard: AI-Based Driver Distraction Detection (CMPS 262 Final Project)

**Team:** Khaled Sabehayoun, Karl Khadij · American University of Beirut · Fall 2026

**Repository:** https://github.com/khsabaayoun-sys/DriverGuard · **Final trained model:** [`models/E4_best.keras`](https://github.com/khsabaayoun-sys/DriverGuard/raw/main/models/E4_best.keras) (selected MobileNetV2) · [`models/svm_selected.joblib`](https://github.com/khsabaayoun-sys/DriverGuard/raw/main/models/svm_selected.joblib) (HOG + SVM)

```python
import tensorflow as tf, numpy as np
model = tf.keras.models.load_model('models/E4_best.keras')   # input: RGB 224x224, raw values 0-255
probs = model.predict(np.zeros((1, 224, 224, 3), np.float32))  # -> 10 class probabilities (c0..c9)
```

We classify ten driver activities (c0–c9) in the State Farm Distracted Driver Detection images and ask:
*does an ImageNet-pretrained MobileNetV2 outperform a tuned HOG + SVM baseline on previously unseen drivers, and does
fine-tuning improve on frozen features?*

## Package contents

| File | What it is |
|---|---|
| `DriverGuard_Final_Complete.ipynb` | Complete Kaggle notebook: audit → splits → HOG + SVM → MobileNetV2 E1–E6 → analyses → one-time test |
| `report/DriverGuard_Final_Report.docx` / `.pdf` (repo) — root of the ZIP | Final report (generated from the notebook's `results_summary.json`) |
| `make_report.js` | Regenerates the report from a results folder: `node make_report.js <results_dir> DriverGuard_Final_Report.docx` |
| `results/` | Outputs of the Kaggle run (tables, figures, JSON), when available |
| `requirements.txt` | Python packages (all are pre-installed on Kaggle) |
| `DriverGuard_Poster.pptx` / `.pdf` | Poster (48×36 in), editable in PowerPoint |
| `models/` | Final trained models: `E4_best.keras` (selected MobileNetV2, load with `tf.keras.models.load_model`, input = raw RGB 224×224 in [0,255]) and `svm_selected.joblib` (HOG + SVM pipeline on 128×128, 16×16-cell HOG features) |

## How to reproduce on Kaggle

1. Open the competition page and accept the rules: <https://www.kaggle.com/competitions/state-farm-distracted-driver-detection>.
2. **Code → New Notebook → File → Import Notebook** and choose `DriverGuard_Final_Complete.ipynb`.
3. **Add Input** → *State Farm Distracted Driver Detection* (competition).
4. **Settings → Accelerator → GPU T4 ×2** (or P100). Leave **Internet on**: the ImageNet weights are downloaded once.
5. **Save Version → Save & Run All (Commit)**. A committed run executes on Kaggle's servers, so a frozen or closed
   browser tab cannot kill it, and everything written to `/kaggle/working` becomes the version's **Output**.
   Expected runtime on T4 ×2 is roughly 1–1.5 h (HOG extraction ~5–10 min, SVM grid ~10–20 min, six CNN runs ~40–60 min).
6. When it finishes, open the version → **Output** and download `driverguard_outputs.zip` (tables, figures, JSON and the
   selected models). Unzip it into `results/` and run `node make_report.js results DriverGuard_Final_Report.docx`.

### Resuming after a crash
Every stage writes its artefacts to `/kaggle/working` and is skipped on re-run when they exist: HOG features are saved in
512-image chunks with a progress file, every SVM fit is appended to `tables/svm_tuning.csv`, and every CNN experiment
saves its best checkpoint, history and validation probabilities. To continue from an earlier committed version, add that
version's output as an input and set `DG_RESUME_FROM` (top of Section 1) to its path, e.g.
`os.environ['DG_RESUME_FROM'] = '/kaggle/input/run-this-driverguard-new'`.

To shorten the run, edit `RUN_EXPERIMENTS` in Section 7 (e.g. `['E1', 'E2', 'E3', 'E4']`).

## Evaluation protocol (no leakage)
* Driver-disjoint GroupShuffleSplit (seed 42): 13,904 train images / 16 drivers, 3,542 validation / 4, 4,978 test / 6.
  Code asserts zero driver overlap and the exact Milestone 1 sizes.
* Scaler, ANOVA feature ranking, PCA and all models are fitted on training drivers only.
* Every choice (HOG configuration, k, C, CNN experiment, checkpoint epoch, temperature) uses **validation macro-F1**
  (temperature: validation NLL).
* The models reported on test are fixed in `model_selection.json` before the test set is used. The test evaluation
  writes `TEST_EVALUATED.json`; later runs load it instead of re-evaluating.
* Augmentation never uses horizontal flips (they would swap the right/left-hand classes c1↔c3 and c2↔c4).

## What was wrong with the earlier notebooks, and what changed

| Problem in the earlier notebooks | Fix |
|---|---|
| Starter cell printed every file path (>100k lines) → browser freezes | Removed; the audit prints counts only |
| HOG accumulated Python lists of float64 arrays, then `np.savez_compressed` of everything; test extraction crashed at 58%; outputs lost because the session was interactive | Parallel chunked extraction written straight to memory-mapped float32 `.npy` with a progress file (resumable); committed runs keep outputs |
| `extract_hog` used `iterrows` (slow) and a single core | joblib workers over chunks |
| Scaler + `SelectKBest` refitted for every grid point | Fitted once per HOG configuration; mathematically identical, much faster |
| CNN checkpoints/early stopping on validation *loss* | Checkpoint and early stopping on validation **macro-F1** (the stated selection metric) |
| Fine-tuned model was reloaded and the frozen variant selected on stale objects; `set_weights` from a reloaded checkpoint | Each experiment saved/reloaded independently; head weights transferred explicitly |
| No augmentation; 160 px only; 5 epochs | Left/right-preserving augmentation, 224 px, label smoothing, AdamW + cosine schedule, ablations E5 (no aug) and E6 (160 px) |
| `preprocess_input` in the data pipeline only | Preprocessing is a layer inside the saved model, so robustness tests and deployment use the identical path |
| Robustness used a 3×3 average pool after resizing and did not test the SVM | Brightness ×{0.5, 0.7, 1.3, 1.6} and Gaussian blur σ ∈ {2, 4, 8} applied at full resolution, for both pipelines |
| SVM latency excluded HOG extraction | Model-only and end-to-end latency (JPEG decode + preprocessing) for both models |
| 64.91% fine-tuning result could not be verified (crash) | Not reported; replaced by E3/E4 |

## Testing performed before the Kaggle run
The notebook was executed end-to-end offline (TensorFlow 2.20 CPU) on a synthetic dataset with the exact State Farm
layout (26 drivers, 10 classes, 640×480 JPEGs): every cell ran, a second run resumed from cache in 15 s and loaded the
locked test results, and a stand-in backbone confirmed that the training loop learns and checkpoints on macro-F1. CNN
accuracy in that smoke test is meaningless (no ImageNet weights offline); **all reported numbers come from the Kaggle run.**

## Final Kaggle run (completed 20 Nov 2026, Tesla T4 ×2, 58 min)
| Model | Val. accuracy | Val. macro-F1 | **Test accuracy** | **Test macro-F1** |
|---|---|---|---|---|
| HOG (128 px, 16×16 cells) + LinearSVC, C = 0.01 | 41.25% | 0.4167 | 54.52% | 0.5343 |
| E1 frozen MobileNetV2, 160 px (best frozen) | 59.20% | 0.5853 | 55.50% | 0.5566 |
| **E4 fine-tuned from block 6, 224 px + aug (selected)** | **87.80%** | **0.8661** | **89.88%** | **0.8902** |

Test = six held-out drivers (4,978 images), evaluated once after selection. `DriverGuard_Final_Complete.ipynb` in this
package is the executed Kaggle version with all outputs; `results/` holds its tables, figures and JSON.

Known limitation in the stored results: in `TEST_EVALUATED.json`, `ece_after_val_temperature` for the *frozen* E1 model
used the temperature fitted for the selected model (E4), so that single value is not meaningful and is not reported.
The values for E4 are correct.
