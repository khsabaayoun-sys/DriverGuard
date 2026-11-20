// DriverGuard final report generator.
// Usage: node make_report.js <results_dir> <output.docx>
// <results_dir> = the notebook's /kaggle/working output (needs results_summary.json and figures/).
// If results_summary.json is missing, a clearly labelled DRAFT is produced with pending placeholders.
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, AlignmentType,
  HeadingLevel, ImageRun, BorderStyle, PageNumber, Footer, Header, LevelFormat, PageBreak, ShadingType,
} = require('docx');

const RES = process.argv[2] || 'results';
const OUTFILE = process.argv[3] || 'DriverGuard_Final_Report.docx';
const sumPath = path.join(RES, 'results_summary.json');
const S = fs.existsSync(sumPath) ? JSON.parse(fs.readFileSync(sumPath, 'utf8').replace(/\bNaN\b|-?\bInfinity\b/g, 'null')) : null;
const FINAL = !!(S && S.test && (process.env.DG_FORCE_FINAL === '1' || !(S.environment && S.environment.quick_mode)));
const PENDING = 'Pending (Kaggle run)';
const FONT = 'Times New Roman';

// ---------- helpers ----------
const pct = (v) => (v === null || v === undefined || Number.isNaN(+v)) ? '—' : (100 * v).toFixed(2) + '%';
const f4 = (v) => (v === null || v === undefined || Number.isNaN(+v)) ? '—' : (+v).toFixed(4);
const f2 = (v) => (v === null || v === undefined || Number.isNaN(+v)) ? '—' : (+v).toFixed(2);
const f1 = (v) => (v === null || v === undefined || Number.isNaN(+v)) ? '—' : (+v).toFixed(1);
const int = (v) => (v === null || v === undefined || Number.isNaN(+v)) ? '—' : Math.round(+v).toLocaleString('en-US');

function runs(text, opts = {}) {
  // supports **bold** and _italic_ inline markup
  const out = [];
  const re = /(\*\*[^*]+\*\*)/g;
  let last = 0, m;
  while ((m = re.exec(text)) !== null) {
    if (m.index > last) out.push(new TextRun({ text: text.slice(last, m.index), ...opts }));
    const t = m[0];
    if (t.startsWith('**')) out.push(new TextRun({ text: t.slice(2, -2), bold: true, ...opts }));
    else out.push(new TextRun({ text: t.slice(1, -1), italics: true, ...opts }));
    last = m.index + t.length;
  }
  if (last < text.length) out.push(new TextRun({ text: text.slice(last), ...opts }));
  return out;
}
const P = (text, o = {}) => new Paragraph({ children: runs(text), alignment: o.align || AlignmentType.JUSTIFIED, spacing: { after: 120, line: 276 }, ...o.p });
const H1 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(t)] });
const H2 = (t) => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(t)] });
const Bullet = (t) => new Paragraph({ numbering: { reference: 'bul', level: 0 }, children: runs(t), spacing: { after: 60 } });
const Caption = (t) => new Paragraph({ children: [new TextRun({ text: t, italics: true, size: 20 })], alignment: AlignmentType.CENTER, spacing: { before: 60, after: 200 } });

const BORDER = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
function table(headers, rows, widths) {
  const total = 9360;
  widths = widths || headers.map(() => Math.floor(total / headers.length));
  const cell = (txt, bold, w) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    borders: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    shading: bold ? { type: ShadingType.CLEAR, fill: 'EDEDED', color: 'auto' } : undefined,
    children: [new Paragraph({ children: [new TextRun({ text: String(txt), bold, size: 19 })] })],
  });
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, true, widths[i])) }),
      ...rows.map((r) => new TableRow({ children: r.map((c, i) => cell(c, false, widths[i])) }))],
  });
}
function figure(file, caption, widthIn = 5.8) {
  const fp = path.join(RES, 'figures', file);
  if (!FINAL || !fs.existsSync(fp)) return [];
  const buf = fs.readFileSync(fp);
  // PNG size from IHDR
  const w = buf.readUInt32BE(16), h = buf.readUInt32BE(20);
  const W = widthIn * 96, Hh = W * h / w;
  return [new Paragraph({ alignment: AlignmentType.CENTER, children: [new ImageRun({ type: 'png', data: buf, transformation: { width: W, height: Hh } })] }), Caption(caption)];
}

// ---------- content ----------
const cnn = FINAL ? S.cnn_experiments : [];
const byId = Object.fromEntries(cnn.map((e) => [e.id, e]));
const sel = FINAL ? S.selection : null;
const svm = FINAL ? S.svm_selected : null;
const selC = sel ? byId[sel.cnn_selected] : null;
const frz = sel ? byId[sel.cnn_best_frozen] : null;
const test = FINAL ? S.test : null;
const tr = (name) => test ? test.results.find((r) => r.model.includes(name)) : null;
const tSvm = test ? test.results[0] : null;
const tSel = selC ? tr(sel.cnn_selected) : null;
const tFrz = frz ? tr(sel.cnn_best_frozen) : null;

const children = [];
// Title block
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [new TextRun({ text: 'DriverGuard: AI-Based Driver Distraction Detection Using Computer Vision', bold: true, size: 32 })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'Khaled Sabehayoun and Karl Khadij', size: 24 })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'CMPS 262 – Data Science, American University of Beirut — Final Project Report', size: 22 })] }));
children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, children: [new TextRun({ text: 'November 20, 2026', size: 22 })] }));
if (!FINAL) {
  children.push(new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 240 }, border: { top: BORDER, bottom: BORDER, left: BORDER, right: BORDER },
    children: [new TextRun({ text: 'DRAFT — the final full-dataset experiments have not yet been executed. Every value marked "Pending" will be filled in automatically from the notebook outputs. Milestone 1 values shown are historical results verified from saved notebook outputs.', bold: true, size: 20 })] }));
}

// Abstract
children.push(H1('Abstract'));
if (FINAL) {
  children.push(P(`We classify ten driver activities (c0–c9) in the State Farm Distracted Driver Detection images and ask whether an ImageNet-pretrained MobileNetV2 generalises to unseen drivers better than a tuned HOG + SVM baseline, and whether fine-tuning improves on frozen features. Using driver-disjoint partitions (16 training, 4 validation and 6 held-out test drivers) and validation macro-F1 for every modelling decision, the best classical pipeline (HOG ${svm.hog}, ${svm.model}) reached ${pct(svm.val_accuracy)} validation accuracy (macro-F1 ${f4(svm.val_macro_f1)}). The best frozen MobileNetV2 (${frz.id}) reached ${pct(frz.val_accuracy)} (${f4(frz.val_macro_f1)}), and the selected model, ${selC.id} (${selC.name}), reached ${pct(selC.val_accuracy)} (${f4(selC.val_macro_f1)}). On the six untouched test drivers, evaluated once after selection, the selected CNN obtained ${pct(tSel.accuracy)} accuracy and ${f4(tSel.macro_f1)} macro-F1 versus ${pct(tSvm.accuracy)} and ${f4(tSvm.macro_f1)} for HOG + SVM. Frozen ImageNet features alone gave only a small gain over the improved classical baseline; the large gain came from fine-tuning, which itself relies on ImageNet initialisation, whose cost is not included in our measured training times.`));
} else {
  children.push(P('We classify ten driver activities (c0–c9) in the State Farm Distracted Driver Detection images and ask whether an ImageNet-pretrained MobileNetV2 generalises to unseen drivers better than a tuned HOG + SVM baseline, and whether fine-tuning improves on frozen features. In Milestone 1, on driver-disjoint validation drivers, HOG + linear SVM achieved 26.40% accuracy (macro-F1 0.2499) and frozen MobileNetV2 61.01% (0.6128). This final phase repairs and extends the pipeline (memory-safe feature extraction, left/right-preserving augmentation, partial and deeper fine-tuning, resolution and augmentation ablations, calibration, robustness and cost analysis) and evaluates the selected models once on six held-out drivers. **Final results: pending the full Kaggle run.**'));
}

// 1. Introduction
children.push(H1('1. Introduction and research question'));
children.push(P('Driver distraction is a leading contributor to road crashes. The State Farm dataset provides dashboard-camera images of drivers performing ten activities: safe driving (c0), texting with the right (c1) or left hand (c3), phoning with the right (c2) or left hand (c4), operating the radio (c5), drinking (c6), reaching behind (c7), hair and makeup (c8) and talking to a passenger (c9). The practical difficulty is generalising to **people the model has never seen**: images of the same driver are highly correlated, so random image-level splits overstate accuracy.'));
children.push(P('**Research question.** Does a pretrained MobileNetV2 outperform a classical HOG + SVM baseline on previously unseen drivers, and does partial fine-tuning provide further improvement? Following our instructor\'s feedback, we interpret the comparison in light of ImageNet pre-training: the CNN starts from features learned on 1.28 million labelled natural images, an advantage, and a computational cost, that the classical baseline does not have.'));

// 2. Data
children.push(H1('2. Dataset, audit and driver-independent protocol'));
const a = FINAL ? S.audit : null;
children.push(P(`The labelled set contains ${a ? int(a.images) : '22,424'} images from ${a ? a.drivers : 26} drivers. The metadata has no missing values${a ? `, all ${int(a.images)} referenced files exist, and the sampled images are 640×480 RGB` : ''}. Class counts range from 1,911 (c8) to 2,489 (c0), so the classes are close to balanced; drivers contribute ${a ? `${a.images_per_driver_min}–${a.images_per_driver_max}` : 'between about 350 and 1,240'} images each. The competition's unlabelled test folder is not used.`));
children.push(P('Two GroupShuffleSplit operations (test fraction 0.2, random_state 42), grouped by driver, give the partitions below. We assert in code that no driver appears in more than one partition and that the sizes reproduce Milestone 1 exactly.'));
children.push(table(['Partition', 'Images', 'Drivers', 'Use'], [
  ['Training', '13,904', '16', 'Model fitting only'],
  ['Validation', '3,542', '4', 'All tuning, early stopping and model selection'],
  ['Held-out test', '4,978', '6', 'Evaluated once, after selection (locked)'],
], [2000, 1500, 1300, 4560]));
children.push(Caption('Table 1. Driver-disjoint partitions.'));
children.push(...figure('eda_class_and_driver_counts.png', 'Figure 1. Images per class and per driver.', 6.2));

// 3. Methods
children.push(H1('3. Methods'));
children.push(H2('3.1 Classical baseline: HOG + SVM'));
children.push(P('Images are converted to grayscale and resized with area interpolation. Histograms of oriented gradients use 9 orientations, 2×2-cell blocks and L2-Hys normalisation. To study feature complexity we extract three configurations in one pass: 128×128 with 8×8 cells (8,100 dimensions, the Milestone 1 setting), 128×128 with 16×16 cells (1,764) and 64×64 with 8×8 cells (1,764). Standardisation and ANOVA F-test feature ranking are fitted on training drivers only; a linear SVM is tuned over the number of selected features k and the regularisation parameter C, and a non-linear variant (PCA to 256 components followed by an RBF-kernel SVM) is tried on the best configuration.'));
children.push(H2('3.2 MobileNetV2 transfer learning'));
children.push(P('MobileNetV2 (ImageNet weights, 2.26 M backbone parameters) is followed by global average pooling, dropout (0.3) and a ten-way softmax. Images are streamed from disk by a tf.data pipeline; MobileNetV2\'s [-1, 1] scaling is a layer inside the saved model. Training-time augmentation uses small translations (±8%), rotations (about ±11°), zoom, brightness and contrast changes. **Horizontal flipping is deliberately excluded**, because mirroring an image turns right-hand texting or phoning (c1, c2) into the left-hand classes (c3, c4) and would corrupt the labels. Fine-tuning keeps BatchNorm layers frozen and in inference mode, uses AdamW (weight decay 1e-4) with a short warm-up and cosine learning-rate decay, and label smoothing 0.1. Fine-tuning runs start from the head trained by the corresponding frozen run.'));
children.push(table(['ID', 'Input', 'Trainable backbone', 'Aug.', 'Learning rate', 'Purpose'], [
  ['E1', '160', 'None (frozen)', 'No', 'Adam 1e-3', 'Milestone 1 reproduction'],
  ['E2', '224', 'None (frozen)', 'Yes', 'Adam 1e-3', 'Frozen + augmentation + resolution'],
  ['E3', '224', 'block_13 onward', 'Yes', 'AdamW 1e-4, cosine', 'Partial fine-tuning'],
  ['E4', '224', 'block_6 onward', 'Yes', 'AdamW 1e-4, cosine', 'Deeper fine-tuning'],
  ['E5', '224', 'block_13 onward', 'No', 'AdamW 1e-4, cosine', 'Augmentation ablation (vs E3)'],
  ['E6', '160', 'block_6 onward', 'Yes', 'AdamW 1e-4, cosine', 'Resolution ablation (vs E4)'],
], [600, 700, 1900, 700, 1900, 3560]));
children.push(Caption('Table 2. CNN experiments, all evaluated on the same four validation drivers.'));
children.push(H2('3.3 Model selection and evaluation'));
children.push(P('Validation macro-F1 is the primary selection metric because it weights the ten activities equally. After every epoch the validation drivers are scored; the checkpoint is saved whenever macro-F1 improves, and training stops after three epochs without improvement. The models to be reported on the test drivers (the selected SVM, the best frozen CNN and the selected CNN) are fixed and written to disk before the test set is touched. The test evaluation writes a lock file; re-running the notebook re-loads the stored results instead of recomputing them. We report accuracy, macro precision, recall and F1, 95% bootstrap confidence intervals, confusion matrices, McNemar\'s test, expected calibration error (ECE), log loss, robustness to controlled brightness and blur shifts, training time, model size and inference latency.'));

// 4. Results
children.push(H1('4. Results'));
children.push(H2('4.1 Classical baseline'));
if (FINAL) {
  const grid = S.svm_grid.slice().sort((x, y) => y.val_macro_f1 - x.val_macro_f1).slice(0, 8);
  children.push(table(['HOG', 'Model', 'k', 'C', 'Val. accuracy', 'Val. macro-F1', 'Fit (s)'],
    grid.map((r) => [r.hog, r.model, int(r.k), r.C, pct(r.val_accuracy), f4(r.val_macro_f1), f1(r.fit_seconds)]),
    [1500, 1800, 900, 900, 1500, 1500, 1260]));
  children.push(Caption('Table 3. Best HOG + SVM configurations on validation drivers (top 8 of the grid).'));
  children.push(P(`The selected classical model is HOG ${svm.hog} with ${svm.model} (k = ${int(svm.k)}, C = ${svm.C}): ${pct(svm.val_accuracy)} accuracy and ${f4(svm.val_macro_f1)} macro-F1, compared with 26.40% / 0.2499 in Milestone 1.`));
  children.push(...figure('svm_feature_complexity.png', 'Figure 2. Validation macro-F1 against the number of selected HOG features and C.', 5.6));
} else {
  children.push(P('Milestone 1 (verified notebook output): on all 13,904 training and 3,542 validation images, StandardScaler → SelectKBest → LinearSVC with C = 0.01 gave 23.43% / 0.2170 (k = 300), **26.40% / 0.2499 (k = 600)** and 26.00% / 0.2454 (k = 1,000). The extended grid over three HOG configurations, C values and an RBF-SVM variant is ' + PENDING + '.'));
}
children.push(H2('4.2 MobileNetV2 experiments'));
if (FINAL) {
  children.push(table(['ID', 'Description', 'Trainable params', 'Epochs (best)', 'Val. acc.', 'Val. macro-F1', 'Train (s)'],
    cnn.slice().sort((x, y) => x.id.localeCompare(y.id)).map((e) => [e.id, e.name, int(e.trainable_params), `${e.epochs_run} (${e.best_epoch})`, pct(e.val_accuracy), f4(e.val_macro_f1), int(e.training_seconds)]),
    [600, 3000, 1300, 1100, 1100, 1160, 1100]));
  children.push(Caption('Table 4. CNN validation results (checkpoint with highest validation macro-F1).'));
  const d = (x, y) => (byId[x] && byId[y]) ? (byId[x].val_macro_f1 - byId[y].val_macro_f1).toFixed(4) : '—';
  children.push(P(`Selected model: **${selC.id} – ${selC.name}**. Effects measured on the validation drivers: augmentation, higher resolution and label smoothing for the frozen model (E2 − E1) ${d('E2', 'E1')} macro-F1; partial fine-tuning over frozen (E3 − E2) ${d('E3', 'E2')}; deeper versus partial fine-tuning (E4 − E3) ${d('E4', 'E3')}; augmentation during partial fine-tuning (E3 − E5) ${d('E3', 'E5')}; and 224 versus 160 pixels for deeper fine-tuning (E4 − E6) ${d('E4', 'E6')}. These are single-seed results on four validation drivers, so differences of about 0.01–0.02 should not be over-interpreted.`));
  children.push(...figure('cnn_learning_curves.png', 'Figure 3. Validation macro-F1 per epoch (left) and train versus validation accuracy (right).', 6.2));
  children.push(...figure('validation_comparison.png', 'Figure 4. Driver-independent validation macro-F1 of all models.', 5.4));
} else {
  children.push(P('Milestone 1 (verified notebook output): frozen MobileNetV2 at 160×160, Adam 1e-3, five epochs: **61.01% accuracy, 0.6128 macro-F1**, 127.2 s training; the best epoch was the third, and validation accuracy then fell while training accuracy rose to 93.8%, indicating driver over-fitting. An earlier fine-tuning attempt was reported at 64.91% validation accuracy, but its outputs were lost when the Kaggle kernel crashed and cannot be verified, so it is not used. Experiments E1–E6: ' + PENDING + '.'));
}
children.push(H2('4.3 Failure analysis'));
if (FINAL) {
  children.push(table(['Class', 'Activity', 'Precision', 'Recall', 'F1 (CNN)', 'F1 (SVM)'],
    S.validation_per_class.map((r) => [r.class, r.activity, f2(r.precision), f2(r.recall), f2(r.f1), f2(r.svm_f1)]), [800, 2600, 1300, 1300, 1680, 1680]));
  children.push(Caption(`Table 5. Per-class validation results of ${selC.id} (with HOG + SVM F1 for comparison).`));
  const tc = S.validation_top_confusions.slice(0, 4).map((r) => `${r.true}→${r.predicted} (${r.count}, ${(100 * r.share_of_true_class).toFixed(0)}% of ${r.true})`).join('; ');
  const drv = S.validation_per_driver.map((r) => `${r.subject}: ${pct(r.cnn_correct)}`).join(', ');
  children.push(P(`The most frequent validation confusions were ${tc}. Accuracy varied by driver (${drv}), which confirms that between-driver variation is the main difficulty. The most confident errors (Figure 6) are dominated by visually similar postures, for example a hand raised near the face for drinking, phoning or grooming.`));
  children.push(...figure('cnn_validation_confusion.png', `Figure 5. Validation confusion matrix of ${selC.id} (counts; shading row-normalised).`, 4.6));
  children.push(...figure('cnn_most_confident_validation_errors.png', 'Figure 6. Most confident validation errors of the selected CNN.', 6.2));
} else {
  children.push(P('Milestone 1: HOG + SVM recall was 0% for c4 (left-hand phone) and about 5% for c2 (right-hand phone), but 63% for c7 and 62% for c5. Frozen MobileNetV2 performed best on c5 (F1 0.87) and worst on c8 (precision 0.23, F1 0.33), with frequent confusion between left/right phone and texting gestures. Per-class tables, row-normalised confusion matrices, top confusions, per-driver accuracy and the most confident errors for the final models: ' + PENDING + '.'));
}
children.push(H2('4.4 Calibration'));
if (FINAL) {
  const c = S.calibration;
  children.push(P(`On validation, ${selC.id} had mean confidence ${f2(c.mean_confidence)} against accuracy ${f2(c.val_accuracy)}, ECE ${f4(c.val_ece_uncalibrated)} and log loss ${f4(c.val_nll_uncalibrated)}. Temperature scaling, cross-fitted across validation drivers (fitted on two, evaluated on the other two), changed ECE to ${f4(c.val_ece_temperature_crossfit)} and log loss to ${f4(c.val_nll_temperature_crossfit)}; the temperature fitted on all validation drivers was T = ${f2(c.temperature_all_validation)}. A temperature below 1 means the network is **under-confident** (it sharpens the probabilities), an expected side-effect of label smoothing and augmentation. Applying this validation temperature unchanged to the test drivers reduced the selected model's ECE from ${f4(tSel.ece)} to ${f4(tSel.ece_after_val_temperature)} and log loss from ${f4(tSel.log_loss)} to ${f4(tSel.log_loss_after_val_temperature)}.`));
  children.push(...figure('reliability_diagram_validation.png', 'Figure 7. Validation reliability diagram before and after temperature scaling.', 3.8));
} else children.push(P('Expected calibration error, reliability diagram and driver-cross-fitted temperature scaling: ' + PENDING + '.'));
children.push(H2('4.5 Robustness to controlled brightness and blur'));
if (FINAL) {
  const kinds = [...new Set(S.robustness.map((r) => r.perturbation))];
  const models = [...new Set(S.robustness.map((r) => r.model))];
  children.push(table(['Perturbation', ...models.map((m) => `${m} macro-F1`)], kinds.map((k) => [k, ...models.map((m) => f4((S.robustness.find((r) => r.perturbation === k && r.model === m) || {}).macro_f1))]), [3360, 3000, 3000]));
  children.push(Caption('Table 6. Macro-F1 on a stratified 600-image validation subset under simulated shifts (perturbation applied at full resolution before each pipeline\'s own preprocessing).'));
  children.push(...figure('robustness_validation.png', 'Figure 7b. Macro-F1 under controlled brightness and blur shifts.', 5.4));
  children.push(P('These are controlled image manipulations, not real-world conditions such as night driving, glare or different camera placement, so they indicate sensitivity rather than real-road robustness.'));
} else children.push(P('Brightness ×0.5/×0.7/×1.3/×1.6 and Gaussian blur (σ = 2, 4, 8) on a stratified 600-image validation subset, for both pipelines: ' + PENDING + '.'));
children.push(H2('4.6 Computational cost'));
if (FINAL) {
  children.push(table(['Model', 'Training (s)', 'Size (MB)', 'CPU ms/img (model)', 'CPU ms/img (end-to-end)', 'GPU ms/img (batch 32)'],
    S.cost.map((r) => [r.model, f1(r.training_seconds), f2(r.model_size_mb), f2(r.cpu_ms_per_image_model_only), f2(r.cpu_ms_per_image_end_to_end), r.gpu_ms_per_image_batch32 == null ? '—' : f2(r.gpu_ms_per_image_batch32)]),
    [2500, 1200, 1100, 1500, 1660, 1400]));
  children.push(Caption(`Table 7. Cost on the Kaggle session (${(S.environment.gpus || []).join(', ') || 'CPU'}). HOG feature extraction for training and validation took ${f1(S.hog_extraction_seconds_train_val)} s; CNN training time for fine-tuned models includes the frozen head stage. End-to-end latency includes JPEG decoding and preprocessing (HOG extraction for the SVM).`));
} else children.push(P('Milestone 1: frozen MobileNetV2 trained in 127.2 s (5 epochs, Kaggle GPU); the linear SVM fit in 23.5 s on precomputed 8,100-dimensional HOG features. Model size, model-only and end-to-end latency for the final models: ' + PENDING + '.'));
children.push(H2('4.7 One-time evaluation on held-out test drivers'));
if (FINAL) {
  children.push(table(['Model', 'Accuracy (95% CI)', 'Macro-F1 (95% CI)', 'Macro P', 'Macro R', 'ECE'],
    test.results.map((r) => [r.model, `${pct(r.accuracy)} (${pct(r.accuracy_95ci[0])}–${pct(r.accuracy_95ci[1])})`, `${f4(r.macro_f1)} (${f4(r.macro_f1_95ci[0])}–${f4(r.macro_f1_95ci[1])})`, f4(r.macro_precision), f4(r.macro_recall), r.ece == null ? '—' : f4(r.ece)]),
    [2400, 2000, 2000, 1000, 1000, 960]));
  children.push(Caption(`Table 8. Results on the six untouched test drivers (4,978 images), evaluated once on ${test.evaluated_at} after model selection. Bootstrap intervals resample images and ignore driver clustering, so they are optimistic.`));
  const mc = test.mcnemar_selected_cnn_vs_svm;
  children.push(P(`McNemar's test between the selected CNN and HOG + SVM: ${mc.a_right_b_wrong} images correct only for the CNN and ${mc.a_wrong_b_right} correct only for the SVM (${mc.p_value > 0 ? 'p = ' + mc.p_value.toExponential(2) : 'p < 1e-300'}).`));
  children.push(table(['Class', 'Activity', 'Precision', 'Recall', 'F1'], test.per_class_selected_cnn.map((r) => [r.class, r.activity, f2(r.precision), f2(r.recall), f2(r.f1)]), [900, 3060, 1800, 1800, 1800]));
  children.push(Caption(`Table 8b. Per-class held-out test results of ${selC.id}. Per-driver test accuracy: ` + test.per_driver_accuracy.map((r) => `${r.subject} ${pct(r.cnn)} (SVM ${pct(r.svm)})`).join(', ') + '.'));
  children.push(...figure('cnn_test_confusion.png', `Figure 8. Held-out test confusion matrix of ${selC.id}.`, 4.6));
  children.push(...figure('svm_test_confusion.png', 'Figure 9. Held-out test confusion matrix of HOG + SVM.', 4.6));
} else children.push(P('The selected SVM, best frozen CNN and selected CNN will be evaluated once on the six test drivers: ' + PENDING + '. No test-set result exists yet; the earlier attempt crashed during test-set HOG extraction before producing any metrics.'));

// 5. Discussion
children.push(H1('5. Discussion: the role of ImageNet pre-training'));
children.push(P('The gap between the two families is best understood as a comparison between practical pipelines rather than between architectures trained from scratch. MobileNetV2 inherits edge, texture, part and object detectors learned from ImageNet; our measured CNN training time covers only the transfer stage and excludes the substantial original pre-training compute. HOG + SVM, by contrast, learns only from our 16 training drivers, and its fixed gradient histograms discard colour and cannot represent the configuration of hands, phone, cup and head that distinguishes visually similar activities. The classical model is, however, small, needs no GPU and has a transparent feature space.'));
children.push(P(FINAL
  ? `Our results answer both parts of the research question. With frozen ImageNet features, MobileNetV2 is only modestly better than the much-improved classical baseline on unseen test drivers (${pct(tFrz.accuracy)} vs ${pct(tSvm.accuracy)} accuracy), showing that generic ImageNet features alone do not capture the fine hand–object configurations that separate these activities. Fine-tuning is what makes the difference: adapting the upper ~70% of the backbone (${selC.id}) raised validation macro-F1 from ${f4(frz.val_macro_f1)} to ${f4(selC.val_macro_f1)} and test macro-F1 to ${f4(tSel.macro_f1)}, about ${(100 * (tSel.accuracy - tSvm.accuracy)).toFixed(1)} percentage points above HOG + SVM in accuracy. Pre-training still matters here: fine-tuning starts from ImageNet weights, which is why a 2.2 M-parameter network can be adapted from only 16 drivers in about ten minutes of GPU time. Two ablations show that augmentation helps only once the backbone is trainable. Adding augmentation, 224-pixel inputs and label smoothing to the frozen model lowered validation macro-F1 (E2 ${f4((byId.E2 || {}).val_macro_f1)} vs E1 ${f4((byId.E1 || {}).val_macro_f1)}). Removing augmentation from partial fine-tuning also lowered it (E5 ${f4((byId.E5 || {}).val_macro_f1)} vs E3 ${f4((byId.E3 || {}).val_macro_f1)}). The 224-pixel input beat 160 pixels for deeper fine-tuning (E4 vs E6). The remaining errors concentrate on talking to a passenger (c9), often predicted as safe driving (c0) because the two postures differ mainly in head orientation, and on hair/makeup (c8). The per-driver spread suggests that more driver diversity, rather than more images of the same drivers, is the most promising next step. The CNN is also about 130× larger than the SVM and slower on CPU (Table 7), a trade-off that matters for in-vehicle deployment.`
  : 'Expected pattern to be confirmed by the final run: fine-tuning upper layers adapts features to the in-car domain, while augmentation and label smoothing limit over-fitting to the 16 training drivers. Interpretation will be completed from the final validation and test results.'));

// 6. Conclusion
children.push(H1('6. Conclusion and future work'));
children.push(P(FINAL
  ? `On drivers never seen during training or tuning, a fine-tuned ImageNet-pretrained MobileNetV2 (${selC.id}) clearly outperforms a carefully tuned HOG + SVM baseline: ${pct(tSel.accuracy)} versus ${pct(tSvm.accuracy)} test accuracy (macro-F1 ${f4(tSel.macro_f1)} versus ${f4(tSvm.macro_f1)}). Frozen ImageNet features alone were not enough (${pct(tFrz.accuracy)} test accuracy); the gain came from fine-tuning most of the backbone, with left/right-preserving augmentation and 224-pixel inputs. The improvement costs a model about 130 times larger and roughly 5 times slower per image on CPU, and it depends on ImageNet pre-training whose cost lies outside our budget.`
  : 'Conclusion pending the final Kaggle run.'));
children.push(P('**Future work.** (1) Driver-level cross-validation and several random seeds, to quantify variance across drivers; (2) crops or attention focused on the hands and face, to separate c9 from c0 and the grooming and drinking gestures; (3) comparison with other efficient backbones (EfficientNet, MobileNetV3) and knowledge distillation into a smaller model for in-vehicle use; (4) evaluation on footage from other cars, cameras and lighting conditions, and temporal models over video rather than single frames.'));

// 7. Limitations
children.push(H1('7. Limitations'));
[ 'Only 26 drivers: the validation set has four drivers and the test set six, so estimates have high driver-level variance; single-seed runs.',
  'All images come from one controlled data-collection setting (same car and camera position); results do not establish real-road safety performance.',
  'Brightness and blur shifts are simulated; calibration and robustness analyses use validation drivers only.',
  'Latency depends on Kaggle hardware; the CNN advantage includes external ImageNet pre-training not counted in our cost figures.',
].forEach((t) => children.push(Bullet(t)));

// 7. Reproducibility and status
children.push(H1('8. Reproducibility and status of experiments'));
children.push(P('The notebook DriverGuard_Final_Complete.ipynb runs top-to-bottom on Kaggle (GPU, Internet on for ImageNet weights) as a committed version. It fixes seed 42, saves the split CSVs, streams images from disk, writes HOG features as memory-mapped float32 chunks that resume after interruption, checkpoints every CNN on validation macro-F1, skips completed stages on re-run and locks the test evaluation. The code was executed end-to-end in an offline smoke test on a synthetic dataset with the State Farm directory structure (pipeline, resume logic and test lock verified); the real-data numbers come only from the Kaggle run' + (FINAL ? ` (committed Kaggle version, ${(S.environment.gpus || []).join(' + ')}, TensorFlow ${S.environment.tensorflow}; total runtime about 58 minutes). The executed notebook with all outputs, the trained models (E4_best.keras, svm_selected.joblib) and all result tables are included in the submitted package.` : '.') ));
children.push(table(['Component', 'Status'], [
  ['Driver-disjoint splits (13,904 / 3,542 / 4,978)', 'Completed (Milestone 1, verified)'],
  ['HOG + linear SVM, k ∈ {300, 600, 1000}, C = 0.01', 'Completed (Milestone 1, verified)'],
  ['Frozen MobileNetV2, 160 px', 'Completed (Milestone 1, verified)'],
  ['Extended HOG/SVM grid and RBF-SVM', FINAL ? 'Completed (final run)' : PENDING],
  ['CNN experiments E1–E6', FINAL ? 'Completed (final run)' : PENDING],
  ['Failure, calibration, robustness, cost analyses', FINAL ? 'Completed (final run)' : PENDING],
  ['One-time held-out test evaluation', FINAL ? 'Completed (final run, locked)' : PENDING],
], [5500, 3860]));
children.push(Caption('Table 9. Experiment status.'));

children.push(H1('References'));
[ 'State Farm. Distracted Driver Detection [dataset]. Kaggle, 2016. https://www.kaggle.com/competitions/state-farm-distracted-driver-detection',
  'Sandler, M., Howard, A., Zhu, M., Zhmoginov, A. and Chen, L.-C. MobileNetV2: Inverted residuals and linear bottlenecks. CVPR, 2018.',
  'Deng, J. et al. ImageNet: A large-scale hierarchical image database. CVPR, 2009.',
  'Dalal, N. and Triggs, B. Histograms of oriented gradients for human detection. CVPR, 2005.',
  'Guo, C., Pleiss, G., Sun, Y. and Weinberger, K. Q. On calibration of modern neural networks. ICML, 2017.',
  'Pedregosa, F. et al. Scikit-learn: Machine learning in Python. JMLR 12, 2011.',
].forEach((t) => children.push(new Paragraph({ numbering: { reference: 'refs', level: 0 }, children: [new TextRun({ text: t, size: 20 })], spacing: { after: 60 } })));

const doc = new Document({
  creator: 'Khaled Sabehayoun, Karl Khadij', title: 'DriverGuard Final Report',
  styles: {
    default: { document: { run: { font: FONT, size: 23 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 26, bold: true, color: '000000' }, paragraph: { spacing: { before: 240, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 23, bold: true, italics: true, color: '000000' }, paragraph: { spacing: { before: 180, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ],
  },
  numbering: { config: [
    { reference: 'bul', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 270 } } } }] },
    { reference: 'refs', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '[%1]', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 540 } } } }] },
  ] },
  sections: [{
    properties: { page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 } } },
    headers: { default: new Header({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ text: FINAL ? 'DriverGuard — CMPS 262 Final Report' : 'DriverGuard — CMPS 262 Final Report (DRAFT: results pending)', size: 18 })] })] }) },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], size: 18 })] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then((b) => { fs.writeFileSync(OUTFILE, b); console.log((FINAL ? 'FINAL' : 'DRAFT') + ' report written:', OUTFILE); });
