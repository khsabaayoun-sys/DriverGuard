from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from PIL import Image

FIG = '/home/claude/kaggle_out/figures/'
MAROON = RGBColor(0x84, 0x0B, 0x24); BLACK = RGBColor(0, 0, 0); GREY = RGBColor(0x55, 0x55, 0x55); WHITE = RGBColor(255, 255, 255)
FONT = 'Times New Roman'
prs = Presentation(); prs.slide_width = Inches(48); prs.slide_height = Inches(36)
s = prs.slides.add_slide(prs.slide_layouts[6])

def box(x, y, w, h, fill=None, line=None):
    sh = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h))
    if fill: sh.fill.solid(); sh.fill.fore_color.rgb = fill
    else: sh.fill.background()
    if line: sh.line.color.rgb = line; sh.line.width = Pt(2)
    else: sh.line.fill.background()
    sh.shadow.inherit = False
    return sh

def text(x, y, w, h, paras, size=26, color=BLACK, align=PP_ALIGN.LEFT):
    tb = s.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h)); tf = tb.text_frame; tf.word_wrap = True
    for i, p in enumerate(paras):
        if isinstance(p, str): p = {'t': p}
        para = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        para.alignment = p.get('align', align); para.space_after = Pt(p.get('after', 10))
        segs = p['t'] if isinstance(p['t'], list) else [p['t']]
        for seg in segs:
            bold = seg.startswith('**') and seg.endswith('**')
            r = para.add_run(); r.text = (('• ' if p.get('bullet') and seg is segs[0] else '') + (seg[2:-2] if bold else seg))
            r.font.size = Pt(p.get('size', size)); r.font.bold = bold or p.get('bold', False); r.font.name = FONT
            r.font.color.rgb = p.get('color', color); r.font.italic = p.get('italic', False)
    return tb

def heading(x, y, w, t):
    box(x, y, w, 0.9, fill=MAROON)
    text(x + 0.3, y + 0.08, w - 0.6, 0.8, [{'t': t, 'size': 36, 'bold': True, 'color': WHITE}])
    return y + 1.15

def pic(path, x, y, w=None, h=None):
    im = Image.open(path); ar = im.height / im.width
    if w and not h: h = w * ar
    if h and not w: w = h / ar
    s.shapes.add_picture(path, Inches(x), Inches(y), Inches(w), Inches(h)); return y + h

# ---- title band ----
box(0, 0, 48, 4.3, fill=MAROON)
text(1, 0.45, 46, 1.6, [{'t': 'DriverGuard: AI-Based Driver Distraction Detection Using Computer Vision', 'size': 80, 'bold': True, 'color': WHITE, 'align': PP_ALIGN.CENTER}])
text(1, 2.25, 46, 1.0, [{'t': 'Khaled Sabehayoun  ·  Karl Khadij', 'size': 44, 'color': WHITE, 'align': PP_ALIGN.CENTER}])
text(1, 3.2, 46, 0.8, [{'t': 'CMPS 262 – Data Science  ·  American University of Beirut  ·  Fall 2026  ·  github.com/khsabaayoun-sys/DriverGuard', 'size': 32, 'color': WHITE, 'align': PP_ALIGN.CENTER}])

M, G = 0.8, 0.8
CW = (48 - 2 * M - 2 * G) / 3
X = [M, M + CW + G, M + 2 * (CW + G)]
TOP = 5.0

# ---- column 1 ----
y = heading(X[0], TOP, CW, '1. Research question')
y = text(X[0], y, CW, 4.2, [
    {'t': ['Does an ', '**ImageNet-pretrained MobileNetV2**', ' recognise driver distraction on ', '**previously unseen drivers**', ' better than a tuned classical ', '**HOG + SVM**', ' baseline, and does ', '**fine-tuning**', ' improve on frozen features?'], 'size': 30},
    {'t': 'Why it matters: distraction is a leading cause of road crashes, and a useful detector must work on new people, not only on those it was trained on.', 'size': 26, 'color': GREY},
]).top
y = heading(X[0], 9.4, CW, '2. Introduction')
text(X[0], y, CW, 5.2, [
    {'t': 'Ten activities (c0–c9) from dashboard-camera images: safe driving, texting and phoning with the right or left hand, radio, drinking, reaching behind, hair/makeup and talking to a passenger.', 'size': 28},
    {'t': ['Images of the same driver are highly correlated, so we split ', '**by driver**', '. We compare a classical pipeline with transfer learning and measure what the CNN gains, and at what cost.'], 'size': 28},
    {'t': ['Finding: fine-tuned MobileNetV2 reaches ', '**89.9% accuracy on six held-out drivers**', ', against 54.5% for HOG + SVM.'], 'size': 28},
])
y = heading(X[0], 16.0, CW, '3. Dataset and evaluation')
text(X[0], y, CW, 4.4, [
    {'t': ['State Farm Distracted Driver Detection (Kaggle): ', '**22,424 labelled images, 26 drivers**', ', 640×480 RGB, near-balanced classes (1,911–2,489 per class); no missing data.'], 'size': 25},
    {'t': ['GroupShuffleSplit by driver (seed 42): ', '**train 13,904 / 16 drivers · validation 3,542 / 4 · test 4,978 / 6**', '. Zero driver overlap is asserted in code.'], 'size': 25},
    {'t': ['All tuning uses ', '**validation macro-F1**', '. The test drivers are evaluated ', '**once**', ', after selection, behind a lock file.'], 'size': 25},
])
pic(FIG + 'eda_example_per_class.png', X[0], 22.4, w=CW)
text(X[0], 22.4 + CW * (594 / 1540) + 0.1, CW, 0.7, [{'t': 'One example image per class.', 'size': 20, 'italic': True, 'color': GREY, 'align': PP_ALIGN.CENTER}])
y = heading(X[0], 29.0, CW, 'Preprocessing')
text(X[0], y, CW, 5.5, [
    {'t': 'HOG: grayscale, area resize; 9 orientations, 2×2 blocks; scaler and ANOVA ranking fitted on training drivers only.', 'size': 24, 'bullet': True},
    {'t': 'CNN: RGB 160/224 px streamed with tf.data; [-1, 1] scaling inside the model.', 'size': 24, 'bullet': True},
    {'t': 'Augmentation: shift, rotation, zoom, brightness, contrast. No horizontal flip, since mirroring turns right-hand classes into left-hand ones (c1↔c3, c2↔c4).', 'size': 24, 'bullet': True},
])

# ---- column 2 ----
y = heading(X[1], TOP, CW, '4. Methods: baseline vs. proposed')
text(X[1], y, CW, 6.6, [
    {'t': ['**Classical baseline.**', ' HOG features (3 configurations: 8,100 or 1,764 dims) → StandardScaler → ANOVA top-k → linear SVM; grid over k and C, plus PCA-256 + RBF-SVM.'], 'size': 25},
    {'t': ['**Proposed.**', ' MobileNetV2 pretrained on ImageNet → global pooling → dropout 0.3 → 10-way softmax. Two stages: train the head on frozen features, then unfreeze upper blocks (BatchNorm frozen), AdamW 1e-4, cosine schedule, label smoothing 0.1, checkpoint on validation macro-F1.'], 'size': 25},
    {'t': ['**Key difference:**', ' the CNN learns task-specific features of hands, objects and head from an ImageNet start; HOG uses fixed gradient histograms without colour.'], 'size': 25},
])
y = heading(X[1], 12.9, CW, '5. Results')
pic('/home/claude/build/poster_test_bars.png', X[1] + 0.6, y, w=CW - 1.2)
yy = y + (CW - 1.2) * (1040 / 1700) + 0.1
text(X[1], yy, CW, 0.7, [{'t': 'Held-out test drivers, evaluated once after model selection.', 'size': 20, 'italic': True, 'color': GREY, 'align': PP_ALIGN.CENTER}])
# results table
rows = [('Model', 'Val. acc', 'Val. F1', 'Test acc', 'Test F1'),
        ('HOG(16×16 cells) + LinearSVC', '41.2%', '0.417', '54.5%', '0.534'),
        ('E1 Frozen, 160 px (Milestone 1 setup)', '59.2%', '0.585', '55.5%', '0.557'),
        ('E2 Frozen, 224 px + aug', '53.9%', '0.523', '–', '–'),
        ('E3 Partial FT (block 13+), 224 + aug', '84.7%', '0.834', '–', '–'),
        ('E4 Deeper FT (block 6+), 224 + aug ★', '87.8%', '0.866', '89.9%', '0.890'),
        ('E5 Partial FT, no augmentation', '80.2%', '0.781', '–', '–'),
        ('E6 Deeper FT, 160 px + aug', '83.7%', '0.815', '–', '–')]
ty = yy + 0.9
tbl = s.shapes.add_table(len(rows), 5, Inches(X[1]), Inches(ty), Inches(CW), Inches(0.62 * len(rows))).table
widths = [6.4, 1.75, 1.6, 1.75, 1.6]
for j, wd in enumerate(widths): tbl.columns[j].width = Inches(wd * CW / sum(widths))
for i, r in enumerate(rows):
    for j, v in enumerate(r):
        c = tbl.cell(i, j); c.text = v; c.fill.solid()
        c.fill.fore_color.rgb = MAROON if i == 0 else (RGBColor(0xF3, 0xE6, 0xE9) if '★' in r[0] else WHITE)
        p = c.text_frame.paragraphs[0]; p.alignment = PP_ALIGN.LEFT if j == 0 else PP_ALIGN.CENTER
        for run in p.runs:
            run.font.size = Pt(21); run.font.name = FONT; run.font.bold = (i == 0 or '★' in r[0])
            run.font.color.rgb = WHITE if i == 0 else BLACK
ty2 = ty + 0.62 * len(rows) + 0.2
text(X[1], ty2, CW, 0.7, [{'t': '★ selected on validation macro-F1. Only pre-declared models were evaluated on test.', 'size': 20, 'italic': True, 'color': GREY}])
pic(FIG + 'cnn_learning_curves.png', X[1], ty2 + 0.9, w=CW)
text(X[1], ty2 + 0.9 + CW * (418 / 1320) + 0.05, CW, 0.7, [{'t': 'Validation macro-F1 per epoch; fine-tuning (E3–E6) separates clearly from frozen features (E1, E2).', 'size': 20, 'italic': True, 'color': GREY, 'align': PP_ALIGN.CENTER}])

# ---- column 3 ----
y = heading(X[2], TOP, CW, 'Failure, robustness and cost analysis')
pic(FIG + 'cnn_test_confusion.png', X[2] + 1.6, y, w=CW - 3.2)
yy = y + (CW - 3.2) * (594 / 704) + 0.05
text(X[2], yy, CW, 0.7, [{'t': 'E4 confusion matrix on test drivers. Main error: talking to passenger (c9) predicted as safe driving (c0).', 'size': 20, 'italic': True, 'color': GREY, 'align': PP_ALIGN.CENTER}])
text(X[2], yy + 1.2, CW, 6.5, [
    {'t': 'Per-class test F1: 0.69 (c9) to 0.97 (c3); per-driver test accuracy 82.7–93.4% (SVM 22.5–80.2%).', 'size': 24, 'bullet': True},
    {'t': 'McNemar: 1,943 test images correct only for the CNN vs 183 only for the SVM (p < 10⁻³⁰⁰).', 'size': 24, 'bullet': True},
    {'t': 'Robustness (600 validation images): E4 macro-F1 0.85 clean, 0.78 at brightness ×0.5, 0.78 at blur σ = 4, 0.39 at blur σ = 8; HOG + SVM 0.43 clean.', 'size': 24, 'bullet': True},
    {'t': 'Calibration: E4 is under-confident (mean confidence 0.68, accuracy 0.88); validation-fitted temperature (T = 0.52) cut test ECE from 0.226 to 0.025.', 'size': 24, 'bullet': True},
    {'t': 'Cost: SVM 0.19 MB, 7 ms per image end-to-end on CPU; E4 25.7 MB, 39 ms on CPU, 1.5 ms per image on GPU; 585 s training (E2 + E4) on T4, excluding ImageNet pre-training.', 'size': 24, 'bullet': True},
])
y = heading(X[2], 22.6, CW, '6. Conclusion')
text(X[2], y, CW, 5.0, [
    {'t': ['Fine-tuned MobileNetV2 generalises far better to unseen drivers than HOG + SVM: ', '**89.9% vs 54.5% test accuracy**', ' (macro-F1 0.890 vs 0.534). ', '**Frozen ImageNet features alone are not enough**', ' (55.5%); the gain comes from fine-tuning from an ImageNet start, with left/right-safe augmentation and 224 px inputs, at the cost of a ~130× larger model.'], 'size': 30},
    {'t': ['**Future work:**', ' driver-level cross-validation and multiple seeds, hand/face-focused crops to separate c9 from c0, lighter backbones or distillation for in-car use, and testing on other cars and lighting.'], 'size': 26},
])
y = heading(X[2], 30.4, CW, '7. References')
text(X[2], y, CW, 4.0, [
    {'t': '[1] State Farm Distracted Driver Detection, Kaggle, 2016.', 'size': 22, 'after': 4},
    {'t': '[2] Sandler et al. MobileNetV2: Inverted residuals and linear bottlenecks. CVPR 2018.', 'size': 22, 'after': 4},
    {'t': '[3] Deng et al. ImageNet: A large-scale hierarchical image database. CVPR 2009.', 'size': 22, 'after': 4},
    {'t': '[4] Dalal & Triggs. Histograms of oriented gradients for human detection. CVPR 2005.', 'size': 22, 'after': 4},
    {'t': '[5] Guo et al. On calibration of modern neural networks. ICML 2017.', 'size': 22, 'after': 4},
])
prs.save('/home/claude/build/DriverGuard_Poster.pptx'); print('saved')
