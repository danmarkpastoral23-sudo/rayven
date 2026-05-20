#!/usr/bin/env node
/**
 * REAM-R-KABLE FIBERS — PPT Generator v2 (Redesigned)
 * Generates a valid .pptx file using Open XML format.
 * Requirements: Node.js + Python3 (built-in zipfile)
 * Run: node generate-pptx.js
 * Output: REAM-R-KABLE-FIBERS.pptx
 */

'use strict';
const fs   = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const os   = require('os');

// ─── SLIDE DATA ───────────────────────────────────────────────────────────────
const slides = [

  {
    title: 'REAM-R-KABLE FIBERS',
    subtitle: 'Paper and Pulp Industry — Industrial Plant Engineering (PROFME 105)',
    bullets: [
      'Prepared by:  DELA CRUZ, JHAN RAYVEN D.  |  BSME 4A',
      'Submitted to: ENGR. MARFEL D. ROSARIO, Instructor',
      'Pangasinan State University — Urdaneta City Campus',
      'College of Engineering and Architecture',
      'A.Y. 2025–2026  |  May 2026',
    ],
    isTitleSlide: true,
  },
  {
    title: 'Company Profile',
    subtitle: 'About REAM-R-KABLE FIBERS',
    bullets: [
      'Location: Barangay Ciudad Real, San Jose del Monte, Bulacan',
      '"REAM" (unit of paper) + "FIBER" (foundation of product) = REMARKABLE',
      '',
      'VISION',
      'To be the premier, most technologically advanced and environmentally',
      'remarkable pulp and paper mill in Southeast Asia.',
      '',
      'MISSION',
      '• Deliver Excellence — high-grade paper exceeding global standards',
      '• Champion Sustainability — water recycling; wastepaper as raw material',
      '• Empower Community — high-value technical jobs in Bulacan',
      '• Drive Innovation — modern engineering in pulping and refining',
    ],
  },
  {
    title: 'Introduction',
    subtitle: 'Paper, Pulp & the Philippine Industry',
    bullets: [
      'PAPER — wood fibers reduced to pulp, formed into a mat, compressed, dried',
      'PULP  — fibrous material from wood/non-wood via chemical or mechanical process',
      '',
      'Paper Grades:',
      '• Printing & Writing  (bond paper, notebooks)',
      '• Newsprint            (newspapers)',
      '• Corrugating / Containerboard  (linerboard, fluting)',
      '• Tissue & Sanitary Papers',
      '• Specialty           (art paper, construction board)',
      '',
      'Philippine Context:',
      '• PICOP Resources Inc. (only integrated mill) ceased operations in 2010',
      '• Industry now relies PRIMARILY on recycled wastepaper',
      '• Wastepaper = 19% of total municipal solid waste in PH',
    ],
  },
  {
    title: 'Plant Location',
    subtitle: 'San Jose Del Monte (SJDM), Bulacan',
    bullets: [
      '~42 km from Metro Manila  |  Landlocked component city',
      '3 national roads  |  4 provincial roads  |  402 city roads  |  15 bridges',
      '',
      'Why SJDM?',
      '• Proximity to Metro Manila — highest wastepaper source density',
      '• Near Region 3 (Central Luzon) — rice straw, bagasse, bamboo supply',
      '• Access to NAIA, Clark, Subic airports and major seaports',
      '• Inside established industrial cluster of Central Luzon',
      '• Bulacan Province: 279,610 ha  |  Cities: Malolos, Meycauayan, SJDM',
    ],
  },
  {
    title: 'Hazard Assessment',
    subtitle: 'Why SJDM is Safe for Industrial Investment',
    bullets: [
      'FLOODING  ➜  LOW RISK',
      '  Eastern location; away from coastal lowlands and river floodplains',
      '',
      'STORM SURGE  ➜  ZERO RISK',
      '  Inland position; surge risk confined to coastal municipalities only',
      '',
      'LIQUEFACTION  ➜  ZERO RISK',
      '  Stable bedrock / thick soil — WHITE ZONE on hazard map',
      '',
      'LANDSLIDE (Earthquake)  ➜  VERY LOW RISK',
      '  Urbanized flatland; away from Sierra Madre unstable slopes',
      '',
      'SJDM falls in white/yellow zones across ALL FOUR hazard maps',
    ],
  },
  {
    title: 'Raw Materials',
    subtitle: 'Industry Inputs',
    bullets: [
      'A.  VIRGIN / WOOD PULP',
      '  Softwood (Pine, Spruce)      — long fibers — strength for printing/writing paper',
      '  Hardwood (Eucalyptus, Birch) — short fibers — improves bulk and opacity',
      '',
      'B.  NON-WOOD PULP',
      '  Sources: Abaca, banana, rice straw, wheat straw, bagasse, bamboo',
      '  Abaca pulp — exceptional strength; key Philippine export commodity',
      '',
      'C.  RECYCLED WASTEPAPER  (Primary Source in Philippines)',
      '  Processed via deinking + fiber separation',
      '  Reduces waste, conserves resources, lowers GHG emissions',
      '  Supplemented by imported virgin pulp when quality demands it',
    ],
  },
  {
    title: 'Chemicals & Additives',
    subtitle: 'Key Chemicals Used in Production',
    bullets: [
      'PULPING CHEMICALS',
      '  Caustic Soda NaOH  — breaks down wood chips into pulp',
      '  Sodium Sulfide Na2S — Kraft cooking liquor; dissolves lignin',
      '',
      'BLEACHING AGENTS',
      '  Chlorine Dioxide ClO2         — high-level whiteness bleaching',
      '  Hydrogen Peroxide H2O2        — eco-friendly brightening agent',
      '  Sodium Hypochlorite           — wastepaper pulp brightness improvement',
      '  Sodium Hydrosulfite           — mechanical pulp brightening',
      '',
      'PAPERMAKING ADDITIVES',
      '  Fillers (clay, CaCO3)         — opacity, brightness, printability',
      '  Sizing agents (starch, rosin) — ink/water resistance',
      '  Retention aids, strength agents, slime control, antifoam, dyes',
    ],
  },
  {
    title: '5 Design Constraints',
    subtitle: 'Engineering Boundaries of the Plant',
    bullets: [
      '1  ENVIRONMENTAL & REGULATORY',
      '   Effluent: BOD, COD, TSS, pH strictly controlled before discharge',
      '   Air: scrubbers for total reduced sulfur (TRS) and particulate matter',
      '',
      '2  RAW MATERIAL & FIBER',
      '   De-inking and contaminant-removal systems for recycled fiber',
      '   Large chip/bale storage areas for uninterrupted supply',
      '',
      '3  ENERGY & UTILITY',
      '   Co-generation (CHP) target — plant aims to be energy self-sufficient',
      '',
      '4  TECHNICAL & PROCESS',
      '   Kraft pulping + ECF/TCF bleaching (no elemental chlorine)',
      '   Duplex stainless steel / titanium for corrosion resistance',
      '',
      '5  SITE & GEOGRAPHICAL',
      '   Near water source for intake/discharge; proximity to ports & highways',
    ],
  },
  {
    title: 'Process Flow Diagram',
    subtitle: 'End-to-End Manufacturing Process',
    bullets: [
      'RAW MATERIAL PREPARATION',
      '  Debarking  ➜  Log Cutting  ➜  Chipping  ➜  Screening  ➜  Storage',
      '  ↓',
      'PULPING  (choose method based on raw material)',
      '  Chemical Pulping (Kraft)  |  Mechanical Pulping  |  Wastepaper Pulping',
      '  ↓',
      'WASHING  ➜  SCREENING  ➜  CLEANING',
      '  ↓',
      'BLEACHING',
      '  ↓',
      'STOCK / PULP PREPARATION',
      '  Beating  ➜  Blending  ➜  Fillers + Sizing + Chemicals',
      '  ↓',
      'PAPERMAKING STAGE',
      '  Headbox ➜ Wire ➜ Press ➜ Dryer ➜ Size Press ➜ Calendar ➜ Reel',
      '  ↓',
      'FINISHED PRODUCT  (Jumbo rolls, trimmed, shipped)',
    ],
  },
  {
    title: 'Pulping Methods',
    subtitle: 'Three Methods Used in the Plant',
    bullets: [
      'METHOD 1 — CHEMICAL PULPING (KRAFT PROCESS)',
      '  Wood chips + White Liquor (NaOH + Na2S) in digester at ~160°C / 10 atm',
      '  Cellulose fibers separated from lignin; black liquor → recovery boiler',
      '',
      'METHOD 2 — MECHANICAL PULPING',
      '  Twin rotating disk refiner physically grinds chips into fibers',
      '  Higher yield, lower strength; retains most lignin',
      '  Bleached with sodium hydrosulfite or hydrogen peroxide',
      '',
      'METHOD 3 — WASTEPAPER PULPING',
      '  Hydrapulper: 3–5% concentration, 30–60 min mechanical agitation',
      '  Foreign objects removed by screens and cleaners (cleaners, reject separator)',
      '  Deinking: ink removed using alkali + surface-active agents',
      '  Bleaching to achieve 55–75% target brightness',
    ],
  },
  {
    title: 'Stock Preparation',
    subtitle: 'Preparing Pulp Before Papermaking',
    bullets: [
      '1.  BEATING',
      '    Disk refiner swells fibers; improves flexibility and surface area',
      '    More inter-fiber bonding points = stronger sheet formation',
      '',
      '2.  BLENDING',
      '    Recycles broke (trim and defective paper from production)',
      '    Blends raw materials + recovered white water from wire section',
      '',
      '3.  ADDITION OF FILLERS, SIZING AGENTS & CHEMICALS',
      '    Fillers (clay, CaCO3)     — opacity, brightness, smoothness',
      '    Sizing agents (starch)    — ink and water resistance',
      '    Retention aids            — reduce fiber loss in wire section',
      '    Strengthening agents      — improve wet and dry sheet strength',
      '    Slime control / antifoaming / dyes — added as process requires',
    ],
  },
  {
    title: 'Papermaking Stage',
    subtitle: 'From Pulp Slurry to Finished Paper',
    bullets: [
      '1.  HEADBOX       — distributes fiber-water slurry uniformly across wire width',
      '2.  WIRE SECTION  — fibers settle into mat; water drains away; sheet forms',
      '3.  PRESS SECTION — rollers + felt cloth; mechanical dewatering;',
      '                    increases strength and improves surface quality',
      '4.  DRYER SECTION — steam-heated cylinders; reduces moisture to 6–10%',
      '5.  SIZE PRESS    — applies starch for water resistance + surface strength',
      '6.  CALENDERING   — steel roller stack; smooths surface, adds gloss,',
      '                    ensures uniform sheet thickness',
      '7.  POPE REEL     — winds paper into large jumbo rolls',
      '8.  REWINDER      — trims and cuts jumbo rolls to sellable widths',
      '                    removes defective sections; final product shipped',
    ],
  },
  {
    title: 'Key Equipment',
    subtitle: 'Major Machinery & Specifications',
    bullets: [
      'RAW MATERIAL HANDLING',
      '  Roller Wood Debarker     5–18 T/H  |  95% peeling rate',
      '  Bale Breaker             100–900 TPD  |  Dual motor drive',
      '  Round Silo Chip Storage  Up to 42 m diameter',
      '',
      'PULPING',
      '  Hydrapulper              1–102 m³  |  Concentration 2–15%',
      '  Displacement Pulp Digester  110–400 m³  |  0.9–1.2 MPa',
      '  Double Disc Refiner (PM20)  Max motor 315–1,800 kW',
      '  Flotation Deinking Cell  Ink removal rate 0.25–0.3; 15–150 T/D',
      '',
      'PAPERMAKING MACHINE',
      '  Headbox / Wire / Press / Dryer / Size Press / Calendar / Reel',
      '  Trim width: up to 6,600 mm  |  Speed: up to 2,200 m/min',
      '  Capacity: up to 1,000 TPD  |  GSM range: 13–500 g/m²',
    ],
  },
  {
    title: 'Manpower',
    subtitle: '80 Total Personnel — Priority Hiring from SJDM & Bulacan',
    bullets: [
      'Top Management ................  4',
      'Supply Chain & Logistics ......  11',
      'Operating Personnel (Mill) ....  19',
      'Quality Control & Assurance ...  4',
      'Maintenance & Engineering .....  8',
      'Packaging .....................  9',
      'EHS & Security ................  6',
      'Sales & Marketing .............  7',
      'Human Resource Administration .  7',
      'Finance & Accounting ..........  5',
      '──────────────────────────────────',
      'TOTAL .........................  80',
    ],
  },
  {
    title: 'Industry Context',
    subtitle: 'Existing Major Paper Mills — Central Luzon',
    bullets: [
      'United Pulp & Paper Corp.          Calumpit, Bulacan',
      '  Products: Corrugating medium, test liner',
      '  Capacity: 230,000 MT / year',
      '',
      'Trust International Paper Corp.    Mabalacat, Pampanga',
      '  Products: Newsprint, printing & writing paper',
      '  Capacity: 230,000 MT / year',
      '',
      'Bataan 2020 Inc.                   Samal, Bataan',
      '  Products: Printing/writing, newsprint, tissue',
      '  Capacity: 73,000 MT / year',
      '',
      'Container Corp. of the Philippines  Quezon City',
      '  Products: Corrugating medium, chipboard',
      '  Capacity: 89,000 MT / year',
      '',
      'SJDM sits at the center of this cluster — shared labor, suppliers & logistics',
    ],
  },
  {
    title: 'Design Computations',
    subtitle: 'Key Calculations Summary',
    bullets: [
      'MASS BALANCE  (Target: 1,200 kg finished paper / hour)',
      '  Dry fiber required       ~1,333 kg / hr',
      '  Process water required   ~25,333 kg / hr',
      '  Total pulp slurry input  ~26,667 kg / hr',
      '  Water evaporated at dryer  1,800 kg / hr',
      '',
      'PAPERMAKING MACHINE PARAMETERS',
      '  Machine speed: 200 m/min  |  Deckle: 3.2 m  |  GSM: 50',
      '  Final dryness: 95%  |  Dryer inlet dryness: 38%',
      '  Wire length: ~27.09 m  |  Size press: max 250 rpm',
      '',
      'PIPING MATERIAL SELECTION',
      '  Kraft / Sulfite pulp         ➜  Copper pipe',
      '  Mechanical / Groundwood      ➜  PVC pipe',
      '  Long-fiber kraft (never dried) ➜  Stainless Steel / PVC',
      '  Soda pulp                    ➜  Steel pipe',
    ],
  },
  {
    title: 'Conclusion',
    subtitle: 'Summary of Key Findings',
    bullets: [
      '✔  Location: SJDM, Bulacan — strategic, low-hazard, logistically superior',
      '✔  Raw Material: Recycled wastepaper (primary) + imported virgin pulp',
      '✔  Three Pulping Methods: Kraft Chemical, Mechanical, Wastepaper',
      '✔  Five Design Constraints fully addressed:',
      '     Environmental, Raw Material, Energy, Technical, Site',
      '✔  Total Workforce: 80 personnel',
      '     Priority hiring: local community of SJDM and Bulacan Province',
      '✔  Natural hazard safe: flood, storm surge, liquefaction, landslide',
      '',
      'OUR VISION',
      '  To be the premier, most technologically advanced, and environmentally',
      '  REMARKABLE pulp and paper mill in Southeast Asia.',
    ],
  },
  {
    title: 'References',
    subtitle: 'Sources & Citations',
    bullets: [
      'AGICO Paper. (n.d.). Pulp mill machinery & paper making equipment.',
      '  https://agicopaper.com/',
      'Andritz. (n.d.). KPR paper capabilities brochure.',
      'Bureau of Energy Efficiency. (n.d.). Best practice manual: Fluid piping systems.',
      'Golden Paper Group. (n.d.). Raw materials for paper making guide.',
      'Paper Mill Machine. (n.d.). Pulp mill section equipment catalogue.',
      '  https://papermillmachine.com/',
      'Provincial Government of Bulacan. (2010). PDPFP 2010–2030.',
      'Sarkar, S. (n.d.). Headbox calculations. Scribd.',
      'DAV University. (n.d.). MEC250-POM: Plant location and layout.',
      'JSS College of Arts. (2023). Department file — Paper Industry Overview.',
      'Landmark University. (n.d.). GEC524: Plant Layout.',
    ],
  },
];



// ─── DESIGN CONSTANTS ────────────────────────────────────────────────────────
// Slide size: 13.33 x 7.5 inches (standard 16:9)
const W  = 8229600;   // 13.33 in × 914400 EMU/in  — wait, use exact:
// Actually: 12192000 EMU wide × 6858000 EMU tall (standard 16:9 in pptx)
const SW = 12192000;
const SH = 6858000;

function pt(n) { return Math.round(n * 12700); }  // points → EMU

// Color palette
const C = {
  navy:    '1A3A5C',   // dark navy — title bar bg
  teal:    '0D7377',   // teal accent bar
  gold:    'F5A623',   // gold accent stripe
  white:   'FFFFFF',
  offwhite:'F0F4F8',   // slide background
  darkgray:'2C3E50',   // body text
  midgray: '5D6D7E',   // subtitle / secondary text
  lightbg: 'D6E8F5',   // left panel bg on title slide
  accent:  '14A0A8',   // bright teal for category labels
  footerbg:'1A3A5C',
};

function xe(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function run(text, { bold=false, sz=1600, color=C.darkgray, font='Calibri' } = {}) {
  const b = bold ? '<a:b/>' : '';
  return `<a:r><a:rPr lang="en-US" dirty="0" b="${bold?1:0}">
    <a:solidFill><a:srgbClr val="${color}"/></a:solidFill>
    <a:latin typeface="${font}"/>
    <a:sz val="${sz}"/>
    ${bold ? '<a:b/>' : ''}
  </a:rPr><a:t>${xe(text)}</a:t></a:r>`;
}

// Simplified run builder
function R(text, sz, color, bold=false) {
  return `<a:r><a:rPr lang="en-US" dirty="0"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:latin typeface="Calibri"/><a:sz val="${sz}"/>${bold?'<a:b/>':''}</a:rPr><a:t>${xe(text)}</a:t></a:r>`;
}

function P(runs, { algn='l', marL=0, indent=0, spcBef=0, spcAft=0, buNone=false } = {}) {
  const bu = buNone ? '<a:buNone/>' : '';
  const sp = spcBef ? `<a:spcBef><a:spcPts val="${spcBef}"/></a:spcBef>` : '';
  const sa = spcAft ? `<a:spcAft><a:spcPts val="${spcAft}"/></a:spcAft>` : '';
  return `<a:p><a:pPr algn="${algn}" marL="${marL}" indent="${indent}">${bu}${sp}${sa}</a:pPr>${runs}</a:p>`;
}

function rect(x, y, cx, cy, fillColor, lnNone=true) {
  const ln = lnNone ? '<a:ln><a:noFill/></a:ln>' : '';
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name=""/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val="${fillColor}"/></a:solidFill>${ln}</p:spPr>
  <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
</p:sp>`;
}

function textbox(x, y, cx, cy, paragraphs, anchor='t', lIns=0, tIns=0, wrap=true) {
  const w = wrap ? 'wrap="square"' : 'wrap="none"';
  return `<p:sp>
  <p:nvSpPr><p:cNvPr id="0" name=""/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>
  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
  <p:txBody>
    <a:bodyPr ${w} anchor="${anchor}" lIns="${lIns}" rIns="0" tIns="${tIns}" bIns="0"/>
    <a:lstStyle/>
    ${paragraphs}
  </p:txBody>
</p:sp>`;
}



// ─── TITLE SLIDE ─────────────────────────────────────────────────────────────
function buildTitleSlide(slide, idx) {
  const leftW  = Math.round(SW * 0.38);
  const rightX = leftW + pt(4);
  const rightW = SW - rightX;

  const shapes = [
    // Full background — off-white
    rect(0, 0, SW, SH, C.offwhite),
    // Left panel — navy
    rect(0, 0, leftW, SH, C.navy),
    // Gold accent stripe between panels
    rect(leftW, 0, pt(6), SH, C.gold),
    // Teal top bar on right panel
    rect(rightX, 0, rightW, pt(8), C.teal),
    // Teal bottom bar
    rect(0, SH - pt(14), SW, pt(14), C.teal),

    // Left panel: big company name stacked
    textbox(pt(24), pt(60), leftW - pt(48), pt(100),
      P(R('REAM-R-KABLE', 4800, C.gold, true), { algn:'l', buNone:true }),
      'b', pt(16), pt(8)
    ),
    textbox(pt(24), pt(160), leftW - pt(48), pt(60),
      P(R('FIBERS', 3600, C.white, true), { algn:'l', buNone:true }),
      't', pt(16), 0
    ),
    // Teal divider line inside left panel
    rect(pt(24), pt(228), leftW - pt(48), pt(3), C.teal),
    // Left panel: subtitle lines
    textbox(pt(24), pt(240), leftW - pt(48), pt(80),
      P(R('Paper and Pulp Industry', 1600, C.lightbg), { algn:'l', buNone:true }),
      't', pt(8), pt(4)
    ),
    textbox(pt(24), pt(272), leftW - pt(48), pt(50),
      P(R('Industrial Plant Engineering', 1400, C.midgray), { algn:'l', buNone:true }),
      't', pt(8), 0
    ),
    textbox(pt(24), pt(308), leftW - pt(48), pt(40),
      P(R('PROFME 105', 1400, C.midgray), { algn:'l', buNone:true }),
      't', pt(8), 0
    ),

    // Right panel: university / details
    textbox(rightX + pt(30), pt(28), rightW - pt(60), pt(50),
      P(R('Pangasinan State University', 1800, C.navy, true), { algn:'l', buNone:true }),
      't', pt(8), pt(4)
    ),
    textbox(rightX + pt(30), pt(68), rightW - pt(60), pt(40),
      P(R('Urdaneta City Campus  ·  College of Engineering and Architecture', 1300, C.midgray), { algn:'l', buNone:true }),
      't', pt(8), 0
    ),
    rect(rightX + pt(30), pt(114), rightW - pt(60), pt(2), C.teal),

    // Presenter info box
    textbox(rightX + pt(30), pt(130), rightW - pt(60), pt(180),
      [
        P(R('Prepared by:', 1300, C.teal, true), { algn:'l', buNone:true, spcBef:200 }),
        P(R('DELA CRUZ, JHAN RAYVEN D.', 2000, C.navy, true), { algn:'l', buNone:true }),
        P(R('BSME 4A', 1500, C.midgray), { algn:'l', buNone:true, spcBef:100 }),
        P(R(''), { buNone:true }),
        P(R('Submitted to:', 1300, C.teal, true), { algn:'l', buNone:true }),
        P(R('ENGR. MARFEL D. ROSARIO', 1700, C.navy, true), { algn:'l', buNone:true }),
        P(R('Instructor', 1400, C.midgray), { algn:'l', buNone:true }),
      ].join(''),
      't', pt(12), pt(8)
    ),

    // Bottom badge
    textbox(rightX + pt(30), SH - pt(44), rightW - pt(60), pt(30),
      P(R('A.Y. 2025–2026  |  May 2026', 1300, C.white), { algn:'l', buNone:true }),
      'ctr', pt(8), 0
    ),

    // Slide number in footer
    textbox(SW - pt(80), SH - pt(12), pt(75), pt(12),
      P(R(`${idx + 1} of ${slides.length}`, 900, C.white), { algn:'r', buNone:true }),
      'ctr', 0, 0
    ),
  ];

  return wrapSlide(shapes);
}



// ─── CONTENT SLIDE ───────────────────────────────────────────────────────────
function buildContentSlide(slide, idx) {
  const headerH  = pt(68);
  const accentH  = pt(6);
  const subtitleH = pt(28);
  const footerH  = pt(24);
  const sideW    = pt(10);  // left accent stripe width
  const padX     = pt(32);
  const bodyY    = headerH + accentH + subtitleH + pt(12);
  const bodyH    = SH - bodyY - footerH - pt(8);

  // Build body paragraphs
  const bodyParas = slide.bullets.map(line => {
    if (line === '') {
      return P(R(' ', 700, C.darkgray), { buNone:true });
    }
    // Section header (ALL CAPS words ending with colon or all-caps line)
    if (/^[A-Z0-9][A-Z0-9 &./–-]{3,}[:]/.test(line) || /^\d\.\s+[A-Z]/.test(line)) {
      return P(R(line, 1500, C.teal, true), { buNone:true, marL: pt(8), spcBef:150 });
    }
    // Bullet lines starting with • ✔ ✓ ➜
    if (/^[•✔✓➜]/.test(line)) {
      return P(R(line, 1400, C.darkgray), { buNone:true, marL: pt(24), spcBef:50 });
    }
    // Indented sub-lines (2+ spaces)
    if (/^ {2,}/.test(line)) {
      return P(R(line.trim(), 1300, C.midgray), { buNone:true, marL: pt(40), spcBef:30 });
    }
    // Numbered steps
    if (/^\d+\./.test(line)) {
      return P(R(line, 1400, C.navy, true), { buNone:true, marL: pt(8), spcBef:120 });
    }
    // Dashed / separator lines
    if (/^[─—─]+$/.test(line.trim())) {
      return P(R(line, 900, C.midgray), { buNone:true });
    }
    // Normal bullet
    return P(R(line, 1400, C.darkgray), { buNone:true, marL: pt(8), spcBef:60 });
  }).join('');

  const shapes = [
    // Background
    rect(0, 0, SW, SH, C.offwhite),
    // Header bar — navy
    rect(0, 0, SW, headerH, C.navy),
    // Gold accent stripe below header
    rect(0, headerH, SW, accentH, C.gold),
    // Subtitle bar — light teal strip
    rect(0, headerH + accentH, SW, subtitleH, C.teal),
    // Left accent stripe on body
    rect(0, bodyY, sideW, bodyH, C.teal),
    // Footer bar
    rect(0, SH - footerH, SW, footerH, C.navy),

    // Header title text
    textbox(padX, pt(10), SW - padX - pt(90), headerH - pt(12),
      P(R(slide.title, 2800, C.white, true), { algn:'l', buNone:true }),
      'ctr', pt(8), 0
    ),

    // Slide number badge (top-right corner)
    textbox(SW - pt(88), pt(10), pt(80), headerH - pt(12),
      P(R(`${idx + 1} / ${slides.length}`, 1300, C.gold, true), { algn:'r', buNone:true }),
      'ctr', 0, 0
    ),

    // Subtitle text
    textbox(padX, headerH + accentH + pt(4), SW - padX * 2, subtitleH - pt(4),
      P(R(slide.subtitle || '', 1300, C.white), { algn:'l', buNone:true }),
      'ctr', pt(8), 0
    ),

    // Body content
    textbox(sideW + pt(16), bodyY + pt(10), SW - sideW - pt(28), bodyH - pt(10),
      bodyParas,
      't', pt(6), pt(4)
    ),

    // Footer text
    textbox(padX, SH - footerH + pt(4), SW - padX * 2, footerH - pt(4),
      P(R('REAM-R-KABLE FIBERS  |  Paper and Pulp Industry  |  BSME 4A  |  A.Y. 2025–2026', 1000, C.lightbg), { algn:'l', buNone:true }),
      'ctr', pt(8), 0
    ),
  ];

  return wrapSlide(shapes);
}



// ─── WRAP SLIDE ──────────────────────────────────────────────────────────────
function wrapSlide(shapes) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr>
      <p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>
    </p:nvGrpSpPr>
    <p:grpSpPr>
      <a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
      <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm>
    </p:grpSpPr>
    ${shapes.join('\n    ')}
  </p:spTree></p:cSld>
</p:sld>`;
}

function buildSlide(slide, idx) {
  return slide.isTitleSlide
    ? buildTitleSlide(slide, idx)
    : buildContentSlide(slide, idx);
}

// ─── OPEN XML PACKAGE ─────────────────────────────────────────────────────────
const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml"
    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
${slides.map((_, i) => `  <Override PartName="/ppt/slides/slide${i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('\n')}
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const PPT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${slides.map((_, i) => `  <Relationship Id="rId${i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i+1}.xml"/>`).join('\n')}
</Relationships>`;

const SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
    <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"
   accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

const MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${SW}" cy="${SH}"/>
    <a:chOff x="0" y="0"/><a:chExt cx="${SW}" cy="${SH}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
</p:sldLayout>`;

const LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

function buildPresentation() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId0"/></p:sldMasterIdLst>
  <p:sldIdLst>
${slides.map((_, i) => `    <p:sldId id="${256+i}" r:id="rId${i+1}"/>`).join('\n')}
  </p:sldIdLst>
  <p:sldSz cx="${SW}" cy="${SH}" type="screen16x9"/>
  <p:notesSz cx="${SH}" cy="${SW}"/>
</p:presentation>`;
}

const SLIDE_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

// ─── BUILD .pptx (via Python3 zipfile) ───────────────────────────────────────
function buildPptx(outFile) {
  const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-'));
  const w   = (rel, txt) => {
    const f = path.join(tmp, rel);
    fs.mkdirSync(path.dirname(f), { recursive: true });
    fs.writeFileSync(f, txt, 'utf8');
  };

  w('[Content_Types].xml',                      CONTENT_TYPES);
  w('_rels/.rels',                               ROOT_RELS);
  w('ppt/presentation.xml',                     buildPresentation());
  w('ppt/_rels/presentation.xml.rels',          PPT_RELS);
  w('ppt/slideMasters/slideMaster1.xml',        SLIDE_MASTER);
  w('ppt/slideMasters/_rels/slideMaster1.xml.rels', MASTER_RELS);
  w('ppt/slideLayouts/slideLayout1.xml',        LAYOUT);
  w('ppt/slideLayouts/_rels/slideLayout1.xml.rels', LAYOUT_RELS);

  slides.forEach((slide, i) => {
    w(`ppt/slides/slide${i+1}.xml`,              buildSlide(slide, i));
    w(`ppt/slides/_rels/slide${i+1}.xml.rels`,   SLIDE_RELS);
  });

  const abs = path.resolve(outFile);
  const py  = path.join(os.tmpdir(), '_mk_pptx.py');
  fs.writeFileSync(py, `
import zipfile, os, sys
d, o = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(o,'w',zipfile.ZIP_DEFLATED) as z:
    for root,dirs,files in os.walk(d):
        for f in files:
            fp=os.path.join(root,f); z.write(fp,os.path.relpath(fp,d))
print("OK:", o)
`);
  try {
    console.log(execSync(`python3 "${py}" "${tmp}" "${abs}"`, { encoding:'utf8' }).trim());
  } finally {
    fs.rmSync(tmp, { recursive:true, force:true });
    try { fs.unlinkSync(py); } catch(_) {}
  }
}

// ─── MAIN ─────────────────────────────────────────────────────────────────────
const out = process.argv[2] || 'REAM-R-KABLE-FIBERS.pptx';
console.log(`Building ${slides.length} slides…`);
buildPptx(out);
console.log(`\nDone! → ${path.resolve(out)}`);
console.log('Open with: PowerPoint · Google Slides · LibreOffice Impress');
