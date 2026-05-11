"""Generate comprehensive PDF research report for bus-level inertia assessment."""

import os, sys
sys.path.insert(0, os.path.dirname(__file__))

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, Image, KeepTogether
)
from reportlab.platypus.tableofcontents import TableOfContents
from reportlab.pdfgen import canvas
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics import renderPDF
import numpy as np

# ── paths ─────────────────────────────────────────────────────────────────────
OUT_DIR    = os.path.join(os.path.dirname(__file__), 'results')
PDF_PATH   = os.path.join(OUT_DIR, 'Bus_Inertia_Research_Report.pdf')
RESULTS    = OUT_DIR

# ── colour palette ─────────────────────────────────────────────────────────────
NAVY   = colors.HexColor('#0d1b4b')
BLUE   = colors.HexColor('#1a3a8c')
LTBLUE = colors.HexColor('#3a6bbf')
ACCENT = colors.HexColor('#e84040')
GOLD   = colors.HexColor('#c8a000')
GREY   = colors.HexColor('#4a4a5a')
LGREY  = colors.HexColor('#e8e8f0')
WHITE  = colors.white

W, H = A4   # 595.27, 841.89 pts


# ── page template with header/footer ──────────────────────────────────────────
def make_page(canvas, doc):
    canvas.saveState()
    # Header bar
    canvas.setFillColor(NAVY)
    canvas.rect(0, H - 28*mm, W, 28*mm, fill=1, stroke=0)
    canvas.setFillColor(WHITE)
    canvas.setFont('Helvetica-Bold', 10)
    canvas.drawString(20*mm, H - 14*mm,
                      'Bus-Level Inertia Assessment for Converter-Dominated Power Systems')
    canvas.setFont('Helvetica', 8)
    canvas.drawRightString(W - 20*mm, H - 14*mm,
                           'Ghosh, Isbeih & El Moursi | IEEE Trans. Power Del. 2023')
    # Accent stripe
    canvas.setFillColor(ACCENT)
    canvas.rect(0, H - 30*mm, W, 2*mm, fill=1, stroke=0)
    # Footer
    canvas.setFillColor(LGREY)
    canvas.rect(0, 0, W, 12*mm, fill=1, stroke=0)
    canvas.setFillColor(GREY)
    canvas.setFont('Helvetica', 7.5)
    canvas.drawString(20*mm, 4*mm, f'Shakil Haque  |  shakil.haqueee@gmail.com')
    canvas.drawRightString(W - 20*mm, 4*mm, f'Page {doc.page}')
    canvas.restoreState()


def make_cover(canvas, doc):
    """Cover page – no header/footer."""
    canvas.saveState()
    # Full dark background
    canvas.setFillColor(NAVY)
    canvas.rect(0, 0, W, H, fill=1, stroke=0)
    # Accent stripe at top
    canvas.setFillColor(ACCENT)
    canvas.rect(0, H - 8*mm, W, 8*mm, fill=1, stroke=0)
    # Accent stripe at bottom
    canvas.setFillColor(LTBLUE)
    canvas.rect(0, 0, W, 6*mm, fill=1, stroke=0)
    # Side accent bar
    canvas.setFillColor(colors.HexColor('#2244aa'))
    canvas.rect(0, 0, 8*mm, H, fill=1, stroke=0)
    canvas.restoreState()


# ── styles ─────────────────────────────────────────────────────────────────────
def build_styles():
    base = getSampleStyleSheet()

    def S(name, parent='Normal', **kw):
        s = ParagraphStyle(name, parent=base[parent], **kw)
        return s

    styles = {
        'cover_title': S('cover_title',
            fontName='Helvetica-Bold', fontSize=26, textColor=WHITE,
            leading=34, spaceAfter=10, alignment=TA_CENTER),
        'cover_sub': S('cover_sub',
            fontName='Helvetica', fontSize=14, textColor=colors.HexColor('#aabbdd'),
            leading=20, spaceAfter=6, alignment=TA_CENTER),
        'cover_author': S('cover_author',
            fontName='Helvetica-Bold', fontSize=11, textColor=GOLD,
            leading=16, spaceAfter=4, alignment=TA_CENTER),
        'cover_info': S('cover_info',
            fontName='Helvetica', fontSize=9, textColor=colors.HexColor('#8899bb'),
            leading=13, spaceAfter=4, alignment=TA_CENTER),

        'h1': S('h1', 'Heading1',
            fontName='Helvetica-Bold', fontSize=16, textColor=NAVY,
            spaceBefore=14, spaceAfter=6, leading=20,
            borderPad=(0,0,2,0)),
        'h2': S('h2', 'Heading2',
            fontName='Helvetica-Bold', fontSize=13, textColor=BLUE,
            spaceBefore=10, spaceAfter=4, leading=17),
        'h3': S('h3', 'Heading3',
            fontName='Helvetica-Bold', fontSize=11, textColor=LTBLUE,
            spaceBefore=7, spaceAfter=3, leading=14),
        'body': S('body',
            fontName='Helvetica', fontSize=10, textColor=colors.black,
            leading=15, spaceAfter=6, alignment=TA_JUSTIFY),
        'bullet': S('bullet',
            fontName='Helvetica', fontSize=10, textColor=colors.black,
            leading=15, spaceAfter=4, leftIndent=14, bulletIndent=4,
            alignment=TA_JUSTIFY),
        'eq': S('eq',
            fontName='Helvetica-BoldOblique', fontSize=10,
            textColor=NAVY, leading=18, spaceAfter=4, leftIndent=20,
            backColor=LGREY, borderPad=6),
        'caption': S('caption',
            fontName='Helvetica-Oblique', fontSize=8.5, textColor=GREY,
            leading=12, spaceAfter=8, alignment=TA_CENTER),
        'ref': S('ref',
            fontName='Helvetica', fontSize=8.5, textColor=GREY,
            leading=13, spaceAfter=3, leftIndent=14),
        'note': S('note',
            fontName='Helvetica-Oblique', fontSize=9, textColor=LTBLUE,
            leading=13, spaceAfter=4, leftIndent=10, borderPad=4,
            backColor=colors.HexColor('#eef2ff')),
    }
    return styles


# ── helper flowables ───────────────────────────────────────────────────────────
def rule(color=LTBLUE, thickness=1):
    return HRFlowable(width='100%', thickness=thickness, color=color,
                      spaceAfter=4, spaceBefore=4)


def section_box(title, styles):
    """Shaded section header box."""
    data = [[Paragraph(title, ParagraphStyle('sb',
             fontName='Helvetica-Bold', fontSize=12,
             textColor=WHITE, leading=15))]]
    t = Table(data, colWidths=[16.5*cm])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), NAVY),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('RIGHTPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [NAVY]),
    ]))
    return t


def result_table(headers, rows, col_widths=None):
    data = [headers] + rows
    if col_widths is None:
        col_widths = [16.5 * cm / len(headers)] * len(headers)
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), NAVY),
        ('TEXTCOLOR',  (0,0), (-1,0), WHITE),
        ('FONTNAME',   (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE',   (0,0), (-1,-1), 9),
        ('FONTNAME',   (0,1), (-1,-1), 'Helvetica'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [WHITE, colors.HexColor('#f0f4ff')]),
        ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor('#ccccdd')),
        ('ALIGN',      (0,0), (-1,-1), 'CENTER'),
        ('VALIGN',     (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING',  (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    return t


def img(fname, width=14*cm, caption=None, styles=None):
    path = os.path.join(RESULTS, fname)
    elems = []
    if os.path.exists(path):
        from PIL import Image as PILImage
        with PILImage.open(path) as im:
            iw, ih = im.size
        aspect = ih / iw
        img_h = width * aspect
        # Cap height to avoid overflow
        if img_h > 14*cm:
            img_h = 14*cm
            width = img_h / aspect
        elems.append(Image(path, width=width, height=img_h))
    if caption and styles:
        elems.append(Paragraph(caption, styles['caption']))
    return elems


# ── content sections ───────────────────────────────────────────────────────────

def cover_page(styles):
    elems = []
    elems.append(Spacer(1, 5.5*cm))
    elems.append(Paragraph(
        'Bus-Level Inertia Assessment<br/>for Converter-Dominated<br/>Power Systems',
        styles['cover_title']))
    elems.append(Spacer(1, 0.5*cm))
    elems.append(Paragraph(
        'A Comprehensive Research Report', styles['cover_sub']))
    elems.append(Spacer(1, 0.8*cm))
    elems.append(Paragraph('───────────────────────────────', styles['cover_info']))
    elems.append(Spacer(1, 0.6*cm))
    elems.append(Paragraph('Based on:', styles['cover_info']))
    elems.append(Paragraph(
        'Ghosh, S., Isbeih, Y. J., &amp; El Moursi, M. S. (2023)',
        styles['cover_author']))
    elems.append(Paragraph(
        '<i>IEEE Transactions on Power Delivery</i>, Vol. 38, No. 4, pp. 2635–2648',
        styles['cover_info']))
    elems.append(Spacer(1, 1.2*cm))
    elems.append(Paragraph('───────────────────────────────', styles['cover_info']))
    elems.append(Spacer(1, 0.6*cm))
    elems.append(Paragraph('Prepared by:', styles['cover_info']))
    elems.append(Paragraph('Shakil Haque', styles['cover_author']))
    elems.append(Paragraph('shakil.haqueee@gmail.com', styles['cover_info']))
    elems.append(Spacer(1, 0.8*cm))
    elems.append(Paragraph('IEEE Test Systems: 14-Bus  |  39-Bus  |  68-Bus  |  118-Bus',
                             styles['cover_info']))
    elems.append(Paragraph('Python Simulation + Y-bus Algorithm + Frequency Analysis',
                             styles['cover_info']))
    elems.append(Spacer(1, 0.6*cm))
    elems.append(Paragraph('May 2026', styles['cover_info']))
    elems.append(PageBreak())
    return elems


def sec_background(styles):
    elems = []
    elems.append(section_box('1. Background and Motivation', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph('1.1 The Global Energy Transition', styles['h2']))
    elems.append(Paragraph(
        'Modern power systems are undergoing a profound transformation driven by climate change '
        'mitigation policies and the rapid deployment of Renewable Energy Sources (RES) such as '
        'wind turbines and solar photovoltaic (PV) systems. Globally, RES capacity additions '
        'reached record levels exceeding 300 GW per year, with projections indicating that '
        'converter-interfaced generation will constitute the majority of installed capacity by 2035. '
        'Countries like the United Kingdom, Denmark, and Australia already experience periods where '
        'RES provides over 70% of instantaneous electricity generation.',
        styles['body']))

    elems.append(Paragraph('1.2 The Inertia Problem', styles['h2']))
    elems.append(Paragraph(
        'Synchronous generators — conventional steam turbines, hydro turbines, and gas turbines — '
        'contain massive rotating masses that store kinetic energy proportional to their <b>inertia '
        'constant H</b> (measured in seconds). This stored energy acts as a natural buffer against '
        'sudden power imbalances: when a large generator trips offline, the kinetic energy of all '
        'rotating machines is instantly released to arrest the frequency fall. This mechanism is '
        'entirely automatic and requires no communication or control action.',
        styles['body']))

    elems.append(Paragraph(
        'Converter-interfaced RES (wind, solar, battery storage) do NOT inherently provide this '
        'inertial response unless specifically programmed to do so via <b>Virtual Inertia (VI)</b> '
        'control. As synchronous generators are displaced by RES, system inertia falls, leading to:',
        styles['body']))

    bullets = [
        'Higher Rate of Change of Frequency (ROCOF) following generation loss events',
        'Deeper frequency nadirs (minimum frequency reached after a disturbance)',
        'Higher risk of Under-Frequency Load Shedding (UFLS) at 59.5 Hz (60 Hz systems)',
        'Potential cascade failures and blackouts if frequency falls below relay thresholds',
        'Faster response time requirements for protective relay settings',
    ]
    for b in bullets:
        elems.append(Paragraph(f'• {b}', styles['bullet']))

    elems.append(Paragraph('1.3 The Swing Equation', styles['h2']))
    elems.append(Paragraph(
        'The fundamental relationship between system inertia and frequency dynamics is '
        'captured by the <b>swing equation</b> of a synchronous generator:', styles['body']))
    elems.append(Paragraph(
        '2H/ω₀ · dω/dt  =  Pm − Pe − D·Δω', styles['eq']))
    elems.append(Paragraph(
        'The system-level ROCOF immediately after a generation loss ΔP is:', styles['body']))
    elems.append(Paragraph(
        'ROCOF = df/dt = −f₀·ΔP / (2·H_sys)    [Hz/s]', styles['eq']))
    elems.append(Paragraph(
        'where H_sys is the system inertia constant [s], f₀ = 60 Hz (or 50 Hz), '
        'and ΔP is the per-unit generation loss. This formula is used by grid operators '
        'to estimate system inertia from online measurements.', styles['body']))

    elems.append(Paragraph('1.4 Why Bus-Level Inertia?', styles['h2']))
    elems.append(Paragraph(
        'Traditional inertia assessment treats the entire power system as a single lumped '
        'inertia constant H_sys — the "Centre of Inertia" (COI) model. While simple, this '
        'approach has critical limitations in modern large-scale networks:', styles['body']))
    bullets2 = [
        'Frequency is NOT uniform across all buses during transients — spatial variation exists',
        'A bus electrically close to a large generator experiences slower frequency decline than one far away',
        'Optimal Virtual Inertia placement requires knowing which buses have the LEAST inertia support',
        'UFLS relay settings should be bus-specific, not based on system-wide average inertia',
        'Converter control (grid-forming inverters) needs localized inertia information for optimal tuning',
    ]
    for b in bullets2:
        elems.append(Paragraph(f'• {b}', styles['bullet']))

    elems.append(Spacer(1, 4))
    elems.append(Paragraph(
        'The Ghosh et al. 2023 paper addresses this fundamental gap by providing a '
        '<b>bus-level inertia assessment framework</b> that assigns a unique inertia '
        'constant H_B(bus i) to every bus in the network, computed from the network\'s '
        'admittance matrix and generator characteristics.', styles['note']))
    elems.append(PageBreak())
    return elems


def sec_existing(styles):
    elems = []
    elems.append(section_box('2. Existing Approaches and Their Limitations', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph('2.1 Review of Existing Inertia Estimation Methods', styles['h2']))
    elems.append(Paragraph(
        'Prior to Ghosh et al. 2023, six categories of methods existed for inertia '
        'assessment in power systems. Each has significant shortcomings when applied '
        'to modern converter-dominated networks:', styles['body']))

    methods = [
        ('L1', 'PMU-Based ROCOF Estimation',
         'Uses Phasor Measurement Unit (PMU) data to estimate df/dt after a disturbance event, '
         'then back-calculates H_sys = −f₀·ΔP / (2·ROCOF). Simple and widely deployed.',
         ['Provides only system-wide (lumped) inertia — cannot identify bus-level spatial variation',
          'Requires a measurable disturbance event (not continuously available)',
          'Inaccurate under high noise levels and converter-dominated conditions',
          'Cannot distinguish between synchronous inertia and virtual inertia contributions']),
        ('L2', 'Ambient Data Methods (Stochastic)',
         'Exploits small ambient load fluctuations to continuously estimate inertia using '
         'Kalman filters, Bayesian estimation, or spectral analysis of frequency signals.',
         ['Estimation variance is very high — inaccurate during RES-dominated periods',
          'Cannot provide spatial resolution at bus level',
          'Requires long data windows (minutes) — not suitable for real-time applications',
          'Sensitive to measurement noise and model assumptions']),
        ('L3', 'Machine Learning / Data-Driven',
         'Trains neural networks or regression models on historical frequency events '
         'to predict system inertia from current operating conditions.',
         ['Black-box models — no physical interpretability',
          'Require extensive labelled training datasets from disturbance events',
          'Cannot generalize to network topology changes or new generator additions',
          'System-level output only — no bus-level granularity']),
        ('L4', 'Centre of Inertia (COI) Model',
         'Classical formulation treating all generators as a single equivalent machine '
         'with H_sys = Σ(H_Gi × S_Gi) / S_base at the electrical centre of the network.',
         ['Fundamentally ignores spatial distribution of inertia across the network',
          'All buses assigned same inertia value — physically incorrect',
          'Cannot support bus-specific virtual inertia placement or UFLS settings',
          'Invalid for networks with significant electrical distances between generators']),
        ('L5', 'Modal Analysis / Small-Signal Methods',
         'Computes inter-area and local oscillation modes from linearized system model. '
         'Uses participation factors to associate modes with specific generators.',
         ['Requires full linearized system model — computationally intensive',
          'Provides oscillation mode information, not directly H_B per bus',
          'Does not account for time-varying operating conditions (load, RES output)',
          'Cannot handle fault conditions where voltages deviate from nominal']),
        ('L6', 'Distributed Inertia Estimation (Regional)',
         'Divides the network into zones/areas and assigns a lumped inertia to each zone '
         'based on generators within that zone.',
         ['Zone boundaries are arbitrary — not based on electrical coupling',
          'Does not capture intra-zone spatial variation',
          'Cannot handle generators that straddle zone boundaries',
          'Zone inertia changes when generators trip — requires manual recalculation']),
    ]

    for code, name, desc, lims in methods:
        box_data = [[
            Paragraph(f'<b>{code}</b>', ParagraphStyle('lc', fontName='Helvetica-Bold',
                      fontSize=10, textColor=WHITE)),
            Paragraph(f'<b>{name}</b>', ParagraphStyle('ln', fontName='Helvetica-Bold',
                      fontSize=10, textColor=WHITE)),
        ]]
        bt = Table(box_data, colWidths=[1.2*cm, 15.3*cm])
        bt.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), BLUE),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
            ('TOPPADDING',  (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ]))
        elems.append(bt)
        elems.append(Paragraph(desc, styles['body']))
        elems.append(Paragraph('<b>Key Limitations:</b>', styles['h3']))
        for l in lims:
            elems.append(Paragraph(f'✗  {l}', styles['bullet']))
        elems.append(Spacer(1, 6))

    elems.append(Paragraph('2.2 Critical Gaps Summary', styles['h2']))
    gap_headers = ['Capability', 'L1', 'L2', 'L3', 'L4', 'L5', 'L6', 'Ghosh 2023']
    gap_rows = [
        ['Bus-level resolution',    '✗','✗','✗','✗','~','✗','✓'],
        ['Real-time computation',   '~','✗','✓','✓','✗','~','✓'],
        ['No disturbance required', '✗','✓','✓','✓','✗','✓','✓'],
        ['Handles fault conditions','✗','✗','~','~','✗','✗','✓'],
        ['Physical interpretability','~','✗','✗','✓','~','~','✓'],
        ['RES penetration tracking', '~','~','~','~','✗','~','✓'],
    ]
    elems.append(result_table(gap_headers, gap_rows,
                               col_widths=[5.5*cm]+[1.43*cm]*7))
    elems.append(Paragraph('Table 1: Comparison of inertia assessment methods. '
                            '✓ = fully supported, ~ = partially, ✗ = not supported.',
                            styles['caption']))
    elems.append(PageBreak())
    return elems


def sec_novelty(styles):
    elems = []
    elems.append(section_box('3. Research Novelty — Ghosh et al. 2023', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph(
        'The Ghosh et al. 2023 paper introduces six novel contributions (N1–N6) that '
        'collectively enable, for the first time, a rigorous bus-level inertia assessment '
        'framework applicable to real-time power system monitoring:', styles['body']))

    novelties = [
        ('N1', 'Augmented Admittance Matrix Partitioning',
         'The full Y-bus admittance matrix is partitioned into four submatrices separating '
         'generator buses (g) and non-generator load buses (b): Y_gg, Y_gb, Y_bg, Y_bb. '
         'This partitioning captures the complete electrical topology of the network including '
         'transformers, parallel lines, and shunt elements.',
         'Y_bus = [Y_gg  Y_gb;  Y_bg  Y_bb]   (Eq. 1)',
         'Enables systematic computation of electrical coupling between every load bus '
         'and every generator bus through the inverse admittance relationship.'),
        ('N2', 'Inertia Weighting Matrix W_c',
         'A novel (nb × ng) matrix W_c that quantifies the electrical coupling from each '
         'load bus to each generator. The (i,j) element represents how much of generator j\'s '
         'inertia "reaches" load bus i through the network impedances.',
         'W_c(t) = { -Y_bb⁻¹(t) × Y_bg(t) .÷ v_b(t) } .× e_g(t)   (Eq. 15)',
         'Physical meaning: buses electrically close to large generators get high W_c '
         'values → high H_B. Buses far from generators get low W_c → low H_B and high ROCOF.'),
        ('N3', 'Bus-Level Inertia Formula',
         'The bus-level inertia H_B for each load bus is computed as a weighted sum of '
         'all online generator inertia constants, where the weights come from W_c.',
         'H_B(t) = W_c(t) × H_G   →   H_B(bus i) = Σⱼ W_c,ij(t) · H_G^j   (Eq. 14)',
         'This gives every bus a unique, physically meaningful inertia value that changes '
         'with network topology, generator dispatch, and voltage profile.'),
        ('N4', 'System Strength Metric',
         'Aggregates all bus inertia values into a single system-level metric that replaces '
         'the traditional lumped H_sys, while retaining spatial information.',
         'Sys_str(t) = Σᵢ H_B(t)   (Eq. 16)',
         'Tracks how total system inertia evolves under varying RES penetration, '
         'generator outages, and load changes — all in real time.'),
        ('N5', 'Fault-Condition Adaptation',
         'During a fault, bus voltages collapse and cannot be measured reliably. '
         'The algorithm adapts by assuming |v_b| = 1 pu (flat voltage profile), '
         'maintaining functionality through fault events.',
         'W_c_fault = -Y_bb⁻¹ × Y_bg   (Eq. — fault mode)',
         'This makes the framework applicable across the full operational envelope '
         'including severe network disturbances and N-1 contingencies.'),
        ('N6', 'Virtual Inertia Placement Guidance',
         'By identifying buses with the lowest H_B values, the framework directly indicates '
         'where converter-based virtual inertia should be deployed for maximum impact on '
         'system frequency stability.',
         'Optimal VI buses = argmin{ H_B(bus i) }   for load buses only',
         'Provides actionable grid-planning guidance: place grid-forming inverters or '
         'BESS at lowest-H_B buses to most efficiently compensate for lost synchronous inertia.'),
    ]

    for code, title, desc, eq, impact in novelties:
        elems.append(KeepTogether([
            Paragraph(f'<b>{code}: {title}</b>', styles['h2']),
            Paragraph(desc, styles['body']),
            Paragraph(f'<i>Formula:</i>  {eq}', styles['eq']),
            Paragraph(f'<i>Impact:</i>  {impact}', styles['note']),
            Spacer(1, 8),
        ]))

    elems.append(rule())
    elems.append(Paragraph('3.1 Algorithm Flowchart (Conceptual)', styles['h2']))
    flow_steps = [
        ['Step', 'Operation', 'Output'],
        ['1', 'Collect branch data (R, X, B) + generator H_G', 'Input parameters'],
        ['2', 'Build Y-bus from branch admittances (pi model)', 'Y ∈ ℂⁿˣⁿ'],
        ['3', 'Partition: Y_gg, Y_gb, Y_bg, Y_bb', '4 sub-matrices'],
        ['4', 'Invert: Y_bb⁻¹  (load bus admittance)', 'Y_bb⁻¹ ∈ ℂⁿᵇˣⁿᵇ'],
        ['5', 'Compute M = -Y_bb⁻¹ × Y_bg', 'M ∈ ℂⁿᵇˣⁿᵍ'],
        ['6', 'Normalize by voltage: W_raw = |M| ./ |v_b|', 'W_raw ∈ ℝⁿᵇˣⁿᵍ'],
        ['7', 'Apply generator status: W_c = W_norm .× e_g', 'W_c ∈ ℝⁿᵇˣⁿᵍ'],
        ['8', 'Bus inertia: H_B = W_c × H_G', 'H_B ∈ ℝⁿᵇ (seconds)'],
        ['9', 'System strength: Sys_str = Σ H_B', 'Scalar [seconds]'],
    ]
    elems.append(result_table(flow_steps[0], flow_steps[1:],
                               col_widths=[1.5*cm, 8*cm, 7*cm]))
    elems.append(Paragraph('Table 2: Step-by-step algorithm execution.',
                            styles['caption']))
    elems.append(PageBreak())
    return elems


def sec_topology(styles):
    elems = []
    elems.append(section_box('4. IEEE Test System Network Topologies', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph(
        'The algorithm was validated on four standard IEEE test systems, ranging from '
        'a small 14-bus network to a large 118-bus transmission system. Each system '
        'represents different scales of complexity, inertia distribution, and network topology.',
        styles['body']))

    systems = [
        ('4.1', 'IEEE 14-Bus System', 14, 5, 20, '100 MVA',
         'GEN at buses 1, 2, 3, 6, 8',
         'The smallest standard IEEE test case. Used to validate basic algorithm correctness. '
         'Five generators with relatively uniform inertia constants (H = 4.0–10.3 s). '
         'The generator at Bus 1 (slack bus) has the highest inertia (H=10.3 s), making Bus 1 '
         'and adjacent buses (4, 5) the highest-inertia load buses.'),
        ('4.2', 'IEEE 39-Bus New England System', 39, 10, 46, '100 MVA',
         'GEN at buses 30–39 (Bus 39 = New England equivalent)',
         'The classic New England 10-machine test system representing a large regional grid. '
         'Generator at Bus 39 is the "infinite bus" equivalent with H=500 s, representing '
         'the aggregate New England system. This dominates the inertia distribution: '
         'Bus 1 (directly connected to Bus 39) achieves H_B=374 s, while electrically distant '
         'Bus 20 has only H_B=26 s — a 14× spatial variation across the network.'),
        ('4.3', 'IEEE 68-Bus NETS-NYPS System', 68, 16, 85, '100 MVA',
         'GEN at buses 1, 9, 22, 31, 46, 54–62, 65, 68',
         'Combined New England Test System (NETS) and New York Power System (NYPS) with '
         '16 generators. Three large equivalent generators (H=500 s each) at buses 59, 65, 68 '
         'represent external grid equivalents. Complex inter-area tie lines create distinct '
         'inertia zones. The NETS-NYPS interface is a classical stability study benchmark.'),
        ('4.4', 'IEEE 118-Bus System', 118, 54, 186, '100 MVA',
         'GEN at 54 buses distributed across the network',
         'The largest standard test system, representing a realistic transmission network. '
         '54 generators with relatively uniform inertia constants (H = 4.17–7.3 s). '
         'Dense connectivity means inertia is more spatially uniform than the 39-bus case. '
         'System strength of 653 s across 118 buses gives an average H_B of ~5.5 s/bus.'),
    ]

    topo_figs = [
        ('topology_ieee14bus.png', 'Fig. 1: IEEE 14-Bus network. Red squares = generators, blue circles = load buses. Colour intensity = H_B value.'),
        ('topology_ieee39bus.png', 'Fig. 2: IEEE 39-Bus New England network. Large H_B gradient visible from Bus 39 (H=500s) outward.'),
        ('topology_ieee68bus.png', 'Fig. 3: IEEE 68-Bus NETS-NYPS network. Three external equivalents at bottom-right create high-H_B zone.'),
        ('topology_ieee118bus.png', 'Fig. 4: IEEE 118-Bus network. Dense topology with more spatially uniform H_B distribution.'),
    ]

    for (sec, name, n_bus, n_gen, n_br, base, gens, desc), (fig, cap) in zip(systems, topo_figs):
        elems.append(Paragraph(f'{sec} {name}', styles['h2']))
        meta = [['Parameter', 'Value'],
                ['Total Buses', str(n_bus)],
                ['Generator Buses', f'{n_gen}  ({gens})'],
                ['Transmission Branches', str(n_br)],
                ['System Base MVA', base]]
        mt = Table(meta, colWidths=[6*cm, 10.5*cm])
        mt.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (0,-1), LGREY),
            ('FONTNAME',   (0,0), (-1,-1), 'Helvetica'),
            ('FONTNAME',   (0,0), (0,-1), 'Helvetica-Bold'),
            ('FONTSIZE',   (0,0), (-1,-1), 9),
            ('GRID',       (0,0), (-1,-1), 0.5, colors.HexColor('#ccccdd')),
            ('TOPPADDING',    (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('LEFTPADDING',   (0,0), (-1,-1), 6),
        ]))
        elems.append(mt)
        elems.append(Spacer(1, 4))
        elems.append(Paragraph(desc, styles['body']))
        elems += img(fig, width=13*cm, caption=cap, styles=styles)
        elems.append(Spacer(1, 8))

    elems.append(PageBreak())
    return elems


def sec_results(styles):
    elems = []
    elems.append(section_box('5. Simulation Results', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph('5.1 Bus-Level Inertia Distribution', styles['h2']))
    elems.append(Paragraph(
        'The bus-level inertia H_B was computed for all buses across all four IEEE test '
        'systems using the Ghosh et al. algorithm. Results are presented below:', styles['body']))

    # Summary table
    summary_headers = ['System', 'Buses', 'Gens', 'Sys. Strength', 'Max H_B (Bus)', 'Min H_B (Bus)', 'Avg H_B']
    summary_rows = [
        ['IEEE 14-Bus',  '14',  '5',  '75.3 s',   '10.3 s (Bus 1)',   '4.0 s (Bus 8)',    '5.4 s'],
        ['IEEE 39-Bus',  '39',  '10', '2860 s',   '374.5 s (Bus 1)',  '26.0 s (Bus 20)',  '73.3 s'],
        ['IEEE 68-Bus',  '68',  '16', '7201 s',   '500.0 s (Bus 59)', '24.3 s (Bus 56)',  '105.9 s'],
        ['IEEE 118-Bus', '118', '54', '652.5 s',  '7.3 s (Bus 4)',    '4.2 s (Bus 10)',   '5.5 s'],
    ]
    elems.append(result_table(summary_headers, summary_rows,
                               col_widths=[2.8*cm, 1.4*cm, 1.2*cm, 2.4*cm, 3.3*cm, 3.3*cm, 2.1*cm]))
    elems.append(Paragraph('Table 3: Summary of bus-level inertia results across all four IEEE test systems.',
                            styles['caption']))

    bar_figs = [
        ('inertia_bar_ieee14bus.png',  'Fig. 5: IEEE 14-Bus inertia distribution. Bar chart with CDF.'),
        ('inertia_bar_ieee39bus.png',  'Fig. 6: IEEE 39-Bus inertia. Large spatial variation (26–374 s for load buses).'),
        ('inertia_bar_ieee68bus.png',  'Fig. 7: IEEE 68-Bus inertia. Three high-H_B external equivalent buses dominate.'),
        ('inertia_bar_ieee118bus.png', 'Fig. 8: IEEE 118-Bus inertia. More uniform distribution due to dense connectivity.'),
    ]
    for fig, cap in bar_figs:
        elems += img(fig, width=13*cm, caption=cap, styles=styles)
        elems.append(Spacer(1, 4))

    elems.append(Paragraph('5.2 Frequency Response Analysis', styles['h2']))
    elems.append(Paragraph(
        'Frequency response simulations were performed for a generation loss of ΔP = 0.1 pu '
        '(10% of system base) for each test system. The swing equation was solved using a '
        'fourth-order Runge-Kutta integrator with governor droop response (R = 5%) '
        'activating after a 0.5 s delay. Results confirm the direct correlation between '
        'H_B and frequency nadir:', styles['body']))

    freq_rows = [
        ['IEEE 14-Bus',  'Bus 5 (H_B=6.3s)',   'Bus 12 (H_B=4.2s)', '59.709 Hz', '59.652 Hz', '0.057 Hz'],
        ['IEEE 39-Bus',  'Bus 1 (H_B=374.5s)', 'Bus 20 (H_B=26.0s)','59.930 Hz', '59.714 Hz', '0.216 Hz'],
        ['IEEE 68-Bus',  'Bus 66 (H_B=500s)',  'Bus 45 (H_B=28.7s)','59.946 Hz', '59.717 Hz', '0.229 Hz'],
        ['IEEE 118-Bus', 'Bus 3 (H_B=6.8s)',   'Bus 57 (H_B=4.8s)', '59.709 Hz', '59.690 Hz', '0.019 Hz'],
    ]
    freq_headers = ['System', 'High-H_B Bus', 'Low-H_B Bus', 'Nadir (High-H_B)', 'Nadir (Low-H_B)', 'Nadir Diff.']
    elems.append(result_table(freq_headers, freq_rows,
                               col_widths=[2.4*cm, 3.6*cm, 3.2*cm, 2.8*cm, 2.8*cm, 2.2*cm]))
    elems.append(Paragraph('Table 4: Frequency nadir comparison between highest and lowest H_B load buses.',
                            styles['caption']))

    freq_figs = [
        ('freq_response_ieee14bus.png',  'Fig. 9: IEEE 14-Bus frequency response. Upper: freq. curves. Lower: ROCOF.'),
        ('freq_response_ieee39bus.png',  'Fig. 10: IEEE 39-Bus response. 14× H_B difference → 0.22 Hz nadir improvement.'),
        ('freq_response_ieee68bus.png',  'Fig. 11: IEEE 68-Bus response. External equivalents provide strong frequency support.'),
        ('freq_response_ieee118bus.png', 'Fig. 12: IEEE 118-Bus. Uniform H_B → smaller nadir differences between buses.'),
    ]
    for fig, cap in freq_figs:
        elems += img(fig, width=13*cm, caption=cap, styles=styles)
        elems.append(Spacer(1, 4))

    elems.append(Paragraph('5.3 Weighting Matrix Analysis', styles['h2']))
    elems.append(Paragraph(
        'The inertia weighting matrix W_c reveals the electrical coupling structure '
        'between load buses and generators. High W_c(i,j) means load bus i is strongly '
        'coupled to generator j — it receives most of its inertia support from generator j.',
        styles['body']))

    wc_figs = [
        ('weighting_matrix_ieee14bus.png', 'Fig. 13: W_c heatmap for IEEE 14-Bus. Warm colours = strong coupling.'),
        ('weighting_matrix_ieee39bus.png', 'Fig. 14: W_c heatmap for IEEE 39-Bus. Generator at Bus 39 dominates most load buses.'),
    ]
    for fig, cap in wc_figs:
        elems += img(fig, width=12*cm, caption=cap, styles=styles)
        elems.append(Spacer(1, 4))

    elems.append(Paragraph('5.4 RES Penetration Impact', styles['h2']))
    elems.append(Paragraph(
        'The impact of increasing RES penetration was studied by progressively displacing '
        'synchronous generators (starting with lowest H_G first) and recomputing system strength:',
        styles['body']))

    res_rows = [
        ['0%',  '75.3 s',  '2860 s', '7201 s', '652.5 s'],
        ['20%', '64.8 s',  '2680 s', '6667 s', '548.5 s'],
        ['40%', '41.0 s',  '2547 s', '6377 s', '428.7 s'],
        ['60%', '30.3 s',  '2252 s', '5987 s', '293.2 s'],
        ['80%', '15.5 s',  '1796 s', '5505 s', '152.8 s'],
    ]
    res_headers = ['RES %', '14-Bus Str.', '39-Bus Str.', '68-Bus Str.', '118-Bus Str.']
    elems.append(result_table(res_headers, res_rows,
                               col_widths=[2*cm, 3.4*cm, 3.4*cm, 3.4*cm, 3.3*cm]))
    elems.append(Paragraph('Table 5: System strength degradation with increasing RES penetration.',
                            styles['caption']))

    res_figs = [
        ('res_impact_ieee14bus.png',  'Fig. 15: IEEE 14-Bus RES impact. Left: system strength. Right: online generator count.'),
        ('res_impact_ieee39bus.png',  'Fig. 16: IEEE 39-Bus RES impact. 80% RES → 37% loss of system strength.'),
    ]
    for fig, cap in res_figs:
        elems += img(fig, width=13*cm, caption=cap, styles=styles)
        elems.append(Spacer(1, 4))

    elems += img('system_comparison_all.png', width=14*cm,
                  caption='Fig. 17: Cross-system inertia comparison. Left: total system strength. Right: avg H_B per bus.',
                  styles=styles)
    elems.append(PageBreak())
    return elems


def sec_analysis(styles):
    elems = []
    elems.append(section_box('6. Analysis and Discussion', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph('6.1 Spatial Inertia Variation and Network Topology', styles['h2']))
    elems.append(Paragraph(
        'The most striking finding across all test systems is the significant spatial variation '
        'in bus-level inertia. For the IEEE 39-bus system, H_B ranges from 26 s (Bus 20) to '
        '374 s (Bus 1 — load bus) — a 14× ratio. This variation is entirely determined by the '
        'electrical distance from each bus to the generator pool:', styles['body']))
    obs = [
        'Buses directly connected to large generators via low-impedance branches have high H_B',
        'Buses at the "electrical periphery" of the network — far from generators — have low H_B',
        'The Y_bb⁻¹ × Y_bg product captures this electrical distance mathematically',
        'Network topology changes (line outages, topology switching) instantly change H_B distribution',
        'The IEEE 118-bus system shows more uniform H_B (4.2–7.3 s) due to dense multi-path connectivity',
    ]
    for o in obs: elems.append(Paragraph(f'• {o}', styles['bullet']))

    elems.append(Paragraph('6.2 Frequency Response Correlation', styles['h2']))
    elems.append(Paragraph(
        'The frequency simulation results validate the physical interpretation of H_B. '
        'For the IEEE 39-bus system, the 14× H_B difference between Bus 1 (374 s) and '
        'Bus 20 (26 s) translates to a 0.22 Hz difference in frequency nadir after a '
        '0.1 pu generation loss. Key observations:', styles['body']))
    obs2 = [
        'ROCOF at Bus 20 is approximately 14× higher than at Bus 1 — matches H_B ratio',
        'Bus 20 approaches the 59.5 Hz UFLS threshold (nadir = 59.71 Hz) while Bus 1 stays at 59.93 Hz',
        'Governor response (droop) partially compensates, but cannot fully overcome inertia deficiency',
        'The 0.5 s governor delay means inertia alone determines the first 0.5 s of frequency trajectory',
        'IEEE 68-bus shows largest nadir difference (0.23 Hz) due to 20× H_B range between buses',
    ]
    for o in obs2: elems.append(Paragraph(f'• {o}', styles['bullet']))

    elems.append(Paragraph('6.3 Weighting Matrix Insights', styles['h2']))
    elems.append(Paragraph(
        'The W_c heatmaps reveal the network\'s "inertia routing" — which generators provide '
        'inertia support to which load buses. For the IEEE 39-bus case:', styles['body']))
    obs3 = [
        'Generator at Bus 39 (H=500 s) dominates W_c for nearly all load buses in buses 1–15',
        'Generators at buses 30–38 (H=24–42 s) primarily serve geographically nearby buses',
        'Load buses 19–29 show more distributed W_c — electrically equidistant from multiple generators',
        'This information directly identifies which generators are "most important" for each bus',
    ]
    for o in obs3: elems.append(Paragraph(f'• {o}', styles['bullet']))

    elems.append(Paragraph('6.4 RES Penetration Critical Thresholds', styles['h2']))
    elems.append(Paragraph(
        'The RES penetration study reveals that system strength degrades monotonically '
        'but non-linearly. For the IEEE 14-bus system:', styles['body']))
    obs4 = [
        '20% RES → 14% strength loss (4 generators still online)',
        '60% RES → 60% strength loss — frequency becomes very sensitive to disturbances',
        '80% RES → 79% strength loss — system strength barely above ROCOF stability limit',
        'IEEE 68-bus degrades more slowly due to large external equivalents (H=500 s) remaining online',
        'The non-linear drop is because small-H generators are tripped first, then large-H at end',
    ]
    for o in obs4: elems.append(Paragraph(f'• {o}', styles['bullet']))

    elems.append(Paragraph('6.5 Virtual Inertia Placement Strategy', styles['h2']))
    elems.append(Paragraph(
        'The most actionable output of this framework is identification of buses where '
        'Virtual Inertia (VI) from grid-forming inverters or BESS should be deployed:', styles['body']))

    vi_rows = [
        ['IEEE 14-Bus',  'Bus 12, 13, 11', '4.2–4.5 s',  'Place VI here first for maximum ROCOF reduction'],
        ['IEEE 39-Bus',  'Bus 20, 21, 19', '26–31 s',    'Bus 20 is >14× weaker than Bus 1 — highest priority'],
        ['IEEE 68-Bus',  'Bus 56, 22, 46', '24–26 s',    'Remote buses from external equivalents most vulnerable'],
        ['IEEE 118-Bus', 'Bus 10, 15, 56', '4.2–4.5 s',  'Uniform network — VI at any of these gives similar benefit'],
    ]
    vi_headers = ['System', 'Priority VI Buses', 'H_B Range', 'Strategy']
    elems.append(result_table(vi_headers, vi_rows,
                               col_widths=[2.6*cm, 3.4*cm, 2.4*cm, 8.1*cm]))
    elems.append(Paragraph('Table 6: Recommended Virtual Inertia placement priority based on H_B analysis.',
                            styles['caption']))
    elems.append(PageBreak())
    return elems


def sec_future(styles):
    elems = []
    elems.append(section_box('7. Future Work', styles))
    elems.append(Spacer(1, 6))

    future_items = [
        ('F1', 'Real-Time PMU Integration',
         'The current algorithm uses offline power flow data. Future work should integrate '
         'real-time PMU measurements to update Y_bus, v_b, and e_g continuously. This would '
         'enable H_B computation at the PMU sampling rate (30–120 samples/second), providing '
         'dynamic inertia monitoring during normal operations and transients.'),
        ('F2', 'Optimal Virtual Inertia Control',
         'Using H_B as a control signal, develop a closed-loop Virtual Inertia Controller '
         'for grid-forming inverters: when H_B at a bus falls below a threshold (due to '
         'generator trip or RES increase), the nearest inverter automatically increases its '
         'emulated inertia. Target: maintain H_B > H_threshold at all buses.'),
        ('F3', 'Probabilistic Inertia Assessment',
         'Model uncertainty in RES output (wind speed, solar irradiance) as stochastic '
         'processes and compute probabilistic H_B distributions: E[H_B], Var[H_B], and '
         'probability(H_B < H_critical). This supports risk-based grid planning under '
         'uncertainty rather than deterministic worst-case analysis.'),
        ('F4', '3-Phase Unbalanced Systems',
         'Extend the algorithm to 3-phase unbalanced distribution networks where single-phase '
         'RES connections create asymmetric inertia distribution. Requires unbalanced '
         'admittance matrix formulation (Y_abc or Y_012 sequence components).'),
        ('F5', 'Converter Inertia Modelling',
         'Grid-forming inverters with Virtual Synchronous Machine (VSM) control provide '
         'programmable inertia. Incorporate these as "pseudo-generators" with time-varying '
         'H_G(t) that can be scheduled or optimized. This enables joint optimization of '
         'synchronous inertia + virtual inertia for minimum cost of frequency stability.'),
        ('F6', 'Hardware-in-the-Loop Validation',
         'Validate the bus-level inertia algorithm against hardware-in-the-loop (HIL) '
         'simulation using real-time digital simulator (RTDS or OPAL-RT). Deploy the '
         'algorithm as an EMS (Energy Management System) module and test with physical '
         'frequency relay hardware to verify H_B-based ROCOF predictions.'),
        ('F7', 'Multi-Area Interconnected Systems',
         'Extend the framework to multi-area systems with DC tie lines (HVDC) and weak '
         'AC interconnections. HVDC links do not transmit inertia directly — this requires '
         'modified Y_bus formulation incorporating converter control characteristics.'),
        ('F8', 'Machine Learning Acceleration',
         'Train a Graph Neural Network (GNN) on the Y-bus structure to predict H_B '
         'without matrix inversion — enabling sub-millisecond computation for very large '
         'networks (>1000 buses) where Y_bb^{-1} becomes computationally prohibitive.'),
    ]

    for code, title, desc in future_items:
        elems.append(KeepTogether([
            Paragraph(f'<b>{code}: {title}</b>', styles['h2']),
            Paragraph(desc, styles['body']),
            Spacer(1, 4),
        ]))

    elems.append(rule())
    elems.append(Paragraph('7.1 Research Roadmap Timeline', styles['h2']))
    timeline_rows = [
        ['Phase', 'Timeline', 'Focus', 'Key Deliverable'],
        ['1 (Near-term)', '0–12 months', 'PMU integration + HIL validation', 'Real-time H_B monitor'],
        ['2 (Mid-term)',  '12–24 months', 'Optimal VI control + probabilistic', 'Adaptive VI controller'],
        ['3 (Long-term)', '24–36 months', 'HVDC + 3-phase + GNN acceleration', 'Production-grade EMS module'],
    ]
    elems.append(result_table(timeline_rows[0], timeline_rows[1:],
                               col_widths=[2.8*cm, 2.8*cm, 5.4*cm, 5.5*cm]))
    elems.append(Paragraph('Table 7: Proposed research roadmap for extending bus-level inertia assessment.',
                            styles['caption']))
    elems.append(PageBreak())
    return elems


def sec_conclusion(styles):
    elems = []
    elems.append(section_box('8. Conclusion', styles))
    elems.append(Spacer(1, 6))

    elems.append(Paragraph(
        'This report presented a comprehensive implementation and analysis of the '
        '<b>Bus-Level Inertia Assessment</b> framework proposed by Ghosh, Isbeih, and '
        'El Moursi in IEEE Transactions on Power Delivery, Vol. 38, No. 4, 2023. '
        'The work was validated across four standard IEEE test systems '
        '(14-bus, 39-bus, 68-bus, 118-bus) using a complete Python simulation framework.',
        styles['body']))

    elems.append(Paragraph('Key conclusions drawn from this work:', styles['h2']))
    conclusions = [
        '<b>Spatial inertia is non-uniform.</b> Bus-level inertia H_B varies significantly '
        'across the network — by up to 14× in the IEEE 39-bus case — making the lumped COI '
        'model fundamentally insufficient for modern grid analysis.',
        '<b>Y-bus partitioning is the key insight.</b> The augmented admittance matrix '
        'partitioned into Y_gg, Y_gb, Y_bg, Y_bb provides a clean mathematical basis for '
        'computing bus-level inertia without requiring time-domain simulation.',
        '<b>ROCOF correlates precisely with H_B.</b> Buses with low H_B experience higher '
        'ROCOF and deeper frequency nadirs — validated by the swing equation simulations '
        'across all four test systems. This correlation enables bus-specific relay settings.',
        '<b>RES integration systematically reduces H_B.</b> Each synchronous generator '
        'displaced by RES reduces system strength proportionally, with the rate of decline '
        'accelerating as the last large-inertia generators are removed.',
        '<b>Virtual Inertia placement is now guided.</b> The framework directly identifies '
        'the most vulnerable buses (lowest H_B) where grid-forming inverters or BESS should '
        'be deployed for maximum frequency stability benefit.',
        '<b>The algorithm is computationally efficient.</b> For all four test systems, '
        'H_B computation requires only one matrix inversion (Y_bb⁻¹) — O(nb³) complexity '
        '— enabling near-real-time updates when integrated with PMU data streams.',
        '<b>The framework is extensible.</b> Future work on probabilistic H_B, HVDC systems, '
        '3-phase networks, and GNN acceleration will make this a production-ready tool '
        'for Energy Management Systems in RES-dominated grids.',
    ]
    for i, c in enumerate(conclusions, 1):
        elems.append(Paragraph(f'{i}.  {c}', styles['bullet']))
        elems.append(Spacer(1, 3))

    elems.append(Spacer(1, 8))
    elems.append(rule(color=GOLD, thickness=2))
    elems.append(Spacer(1, 6))
    elems.append(Paragraph(
        'The increasing penetration of converter-interfaced renewable energy is a defining '
        'challenge of 21st-century power engineering. As synchronous generators retire and '
        'wind turbines and solar arrays take their place, the traditional inertia buffers of '
        'the power grid disappear. The bus-level inertia framework presented here provides '
        'grid operators with the spatial resolution needed to understand, monitor, and '
        '<b>proactively manage</b> inertia distribution — ensuring frequency stability in '
        'the converter-dominated grids of the future.',
        styles['note']))

    elems.append(Spacer(1, 12))
    elems.append(section_box('References', styles))
    refs = [
        '[1]  S. Ghosh, Y. J. Isbeih, and M. S. El Moursi, "Bus-Level Inertia Assessment for '
        'Converter-Dominated Power Systems," IEEE Transactions on Power Delivery, vol. 38, '
        'no. 4, pp. 2635–2648, Aug. 2023. DOI: 10.1109/TPWRD.2023.3237879',
        '[2]  ENTSO-E, "Frequency Stability Evaluation Criteria for the Synchronous Zone of '
        'Continental Europe," Technical Report, Mar. 2016.',
        '[3]  P. Kundur, Power System Stability and Control. New York: McGraw-Hill, 1994.',
        '[4]  F. Milano et al., "Foundations and Challenges of Low-Inertia Systems," '
        'IEEE Power & Energy Society General Meeting, 2018.',
        '[5]  IEEE PES Task Force on Low-Inertia Systems, "Stability and Control of Low-Inertia '
        'Power Systems: A Review," IEEE Trans. Power Systems, 2021.',
        '[6]  Z. Guo et al., "Assessment of Power System Inertia Using PMU Measurements," '
        'IEEE Trans. Power Delivery, vol. 35, pp. 3079–3090, 2020.',
        '[7]  IEEE Standard 39-bus New England Test System Data, Power Systems Test Case '
        'Archive, University of Washington, 2020. [Online]. Available: labs.ece.uw.edu/pstca/',
    ]
    for r in refs:
        elems.append(Paragraph(r, styles['ref']))
        elems.append(Spacer(1, 2))

    return elems


# ── build PDF ─────────────────────────────────────────────────────────────────
def main():
    try:
        from PIL import Image
    except ImportError:
        print("Installing Pillow for image size detection...")
        import subprocess
        subprocess.run(['pip', 'install', 'Pillow', '--quiet'])
        from PIL import Image

    styles = build_styles()
    story = []

    # Cover (special page callback)
    story += cover_page(styles)
    story += sec_background(styles)
    story += sec_existing(styles)
    story += sec_novelty(styles)
    story += sec_topology(styles)
    story += sec_results(styles)
    story += sec_analysis(styles)
    story += sec_future(styles)
    story += sec_conclusion(styles)

    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=3.5*cm, bottomMargin=2*cm,
        title='Bus-Level Inertia Assessment Research Report',
        author='Shakil Haque',
        subject='Power Systems Inertia',
    )

    print(f'Building PDF...')
    doc.build(story, onFirstPage=make_cover, onLaterPages=make_page)
    print(f'PDF saved: {PDF_PATH}')


if __name__ == '__main__':
    main()
