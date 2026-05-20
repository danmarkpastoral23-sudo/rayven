#!/usr/bin/env node
/**
 * REAM-R-KABLE FIBERS — PPT Generator
 * Generates a valid .pptx file from scratch using Open XML format.
 * Requirements: Node.js (built-in only) + Python3 (for ZIP creation)
 * Run: node generate-pptx.js
 * Output: REAM-R-KABLE-FIBERS.pptx
 */

const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const os = require('os');

// ─────────────────────────────────────────────
// SLIDE DATA — 18 slides
// ─────────────────────────────────────────────
const slides = [
  {
    title: 'REAM-R-KABLE FIBERS',
    bullets: [
      'PAPER AND PULP INDUSTRY',
      'Industrial Plant Engineering (PROFME 105)',
      '',
      'Prepared by: DELA CRUZ, JHAN RAYVEN D. | BSME 4A',
      'Submitted to: ENGR. MARFEL D. ROSARIO',
      'Pangasinan State University — Urdaneta City Campus',
      'College of Engineering and Architecture',
      'A.Y. 2025–2026 | May 2026',
    ],
  },
  {
    title: 'Company Profile',
    bullets: [
      'Company: REAM-R-KABLE FIBERS',
      'Location: Barangay Ciudad Real, San Jose del Monte, Bulacan',
      '"REAM" = unit of paper + "FIBER" = foundation of our product',
      '',
      'VISION: Premier, most technologically advanced and environmentally',
      'remarkable pulp and paper mill in Southeast Asia',
      '',
      'MISSION:',
      '• Deliver Excellence — high-grade paper exceeding global standards',
      '• Champion Sustainability — water recycling; wastepaper as raw material',
      '• Empower Community — high-value jobs in Bulacan',
      '• Drive Innovation — modern engineering in pulping and refining',
    ],
  },
  {
    title: 'Introduction',
    bullets: [
      'PAPER — from wood fibers reduced to pulp, formed, compressed, dried',
      'PULP — fibrous material from wood/non-wood (cotton, bagasse, rice straw)',
      '',
      'Paper Grades:',
      '• Printing & Writing (bond paper, notebooks)',
      '• Newsprint (newspapers)',
      '• Corrugating / Containerboard (linerboard, fluting)',
      '• Tissue & Sanitary Papers',
      '• Specialty (art paper, construction board)',
      '',
      'Philippine Context:',
      '• Only integrated mill (PICOP) ceased operations in 2010',
      '• Industry now relies primarily on RECYCLED WASTEPAPER',
      '• Wastepaper = 19% of total municipal solid waste in PH',
    ],
  },
  {
    title: 'Plant Location — San Jose Del Monte, Bulacan',
    bullets: [
      'SJDM: Landlocked component city, ~42 km from Metro Manila',
      '3 national roads | 4 provincial roads | 402 city roads | 15 bridges',
      '',
      'Why SJDM?',
      '• Near Metro Manila — highest density of wastepaper sources',
      '• Near Region 3 (Central Luzon) — rice straw, bagasse, bamboo',
      '• Access to NAIA, Clark, Subic airports and seaports',
      '• Part of established industrial cluster in Central Luzon',
      '• Bulacan Province total land area: 279,610 hectares',
      '• Cities: Malolos (capital), Meycauayan, SJDM',
    ],
  },
  {
    title: 'Hazard Assessment',
    bullets: [
      'Why SJDM is Safe for Industrial Use:',
      '',
      'FLOODING       → LOW RISK',
      '  Eastern location; away from coastal lowlands & river systems',
      '',
      'STORM SURGE    → NO RISK',
      '  Inland position; surge confined to coastal municipalities only',
      '',
      'LIQUEFACTION   → NO RISK',
      '  Stable bedrock/thick soil — white zone on hazard map',
      '',
      'LANDSLIDE      → VERY LOW RISK',
      '  Urbanized area away from Sierra Madre unstable slopes',
      '',
      'SJDM = white/yellow zones across ALL four hazard maps',
    ],
  },
  {
    title: 'Raw Materials',
    bullets: [
      'A. VIRGIN / WOOD PULP',
      '  Softwood (Pine, Spruce) — long fibers — best for printing & writing paper',
      '  Hardwood (Eucalyptus, Birch) — short fibers — improves bulk & opacity',
      '',
      'B. NON-WOOD PULP',
      '  Abaca, banana, rice straw, wheat straw, bagasse, bamboo',
      '  Abaca pulp — exceptional strength; key PH export commodity',
      '  Lower tensile strength but cost-effective and sustainable',
      '',
      'C. RECYCLED WASTEPAPER (Primary Source in Philippines)',
      '  Processed via deinking + fiber separation',
      '  Reduces waste, conserves resources, lowers GHG emissions',
      '  Supplemented by imported virgin pulp for quality requirements',
    ],
  },
  {
    title: 'Chemicals & Additives',
    bullets: [
      'PULPING CHEMICALS:',
      '  Caustic Soda (NaOH) — breaks down wood chips into pulp',
      '  Sodium Sulfide (Na2S) — Kraft cooking liquor; dissolves lignin',
      '',
      'BLEACHING AGENTS:',
      '  Chlorine Dioxide (ClO2) — high-level whiteness bleaching',
      '  Hydrogen Peroxide (H2O2) — eco-friendly brightening',
      '  Sodium Hypochlorite — wastepaper pulp brightness improvement',
      '  Sodium Hydrosulfite — mechanical pulp brightening',
      '',
      'PAPERMAKING ADDITIVES:',
      '  Fillers (clay, CaCO3) — opacity, brightness, printability',
      '  Sizing agents — ink resistance, water resistance',
      '  Retention aids, strength agents, slime control, antifoam, dyes',
    ],
  },
  {
    title: '5 Design Constraints',
    bullets: [
      '1. ENVIRONMENTAL & REGULATORY',
      '   Effluent: control BOD, COD, TSS, pH before discharge',
      '   Air: scrubbers for sulfur compounds & particulate matter',
      '',
      '2. RAW MATERIAL & FIBER',
      '   De-inking & contaminant-removal for recycled fiber',
      '   Large chip storage for continuous supply',
      '',
      '3. ENERGY & UTILITY',
      '   Co-generation (CHP) target — energy self-sufficient plant',
      '',
      '4. TECHNICAL & PROCESS',
      '   Kraft pulping + ECF/TCF bleaching (no elemental chlorine)',
      '   Duplex stainless steel / titanium for corrosion resistance',
      '',
      '5. SITE & GEOGRAPHICAL',
      '   Near water source; proximity to ports and highways',
    ],
  },
  {
    title: 'Process Flow Diagram',
    bullets: [
      'RAW MATERIAL PREPARATION',
      '  Debarking → Log Cutting → Chipping → Screening → Storage',
      '        ↓',
      'PULPING',
      '  Chemical Pulping (Kraft) | Mechanical Pulping | Wastepaper Pulping',
      '        ↓',
      'WASHING, SCREENING & CLEANING',
      '        ↓',
      'BLEACHING',
      '        ↓',
      'STOCK / PULP PREPARATION',
      '  Beating → Blending → Fillers + Sizing + Chemicals',
      '        ↓',
      'PAPERMAKING STAGE',
      '  Headbox→Wire→Press→Dryer→Size Press→Calendar→Reel',
      '        ↓',
      'FINISHED PRODUCT',
    ],
  },
  {
    title: 'Pulping Methods',
    bullets: [
      'METHOD 1 — CHEMICAL PULPING (KRAFT PROCESS)',
      '  Wood chips + White Liquor (NaOH + Na2S) in digester',
      '  ~160°C, 10 atm → cellulose fibers separated from lignin',
      '  Black liquor → recovery boiler → energy + chemical recycling',
      '',
      'METHOD 2 — MECHANICAL PULPING',
      '  Twin rotating disk refiner physically grinds chips into fibers',
      '  Higher yield, lower strength; retains most lignin',
      '  Bleached with sodium hydrosulfite or hydrogen peroxide',
      '',
      'METHOD 3 — WASTEPAPER PULPING',
      '  Hydrapulper: 3–5% concentration, 30–60 min agitation',
      '  Foreign objects removed by screens and cleaners',
      '  Deinking: ink removed with alkali + surface-active agents',
      '  Bleaching to achieve 55–75% brightness',
    ],
  },
  {
    title: 'Stock Preparation',
    bullets: [
      '1. BEATING',
      '   Disk refiner swells fibers; improves flexibility & surface area',
      '   More bonding points → stronger sheet formation',
      '',
      '2. BLENDING',
      '   Recycles broke (trim/defective paper from production)',
      '   Blends raw materials + recovered white water',
      '',
      '3. ADDITION OF FILLERS, SIZING AGENTS & CHEMICALS',
      '   Fillers (clay, CaCO3) → opacity, brightness, smoothness',
      '   Sizing agents (starch, rosin) → ink/water resistance',
      '   Retention aids → reduce fiber loss in wire section',
      '   Strengthening agents → improve sheet strength',
      '   Slime control agents, antifoaming agents, dyes → as needed',
    ],
  },
  {
    title: 'Papermaking Stage',
    bullets: [
      '1. HEADBOX — distributes fiber-water slurry across wire width',
      '',
      '2. WIRE SECTION — fibers settle into mat; water drains away',
      '',
      '3. PRESS SECTION — rollers + felt cloth; mechanical dewatering',
      '   increases strength, improves surface quality',
      '',
      '4. DRYER SECTION — steam-heated cylinders; dries to 6–10% moisture',
      '',
      '5. SIZE PRESS — applies starch; water resistance + surface strength',
      '',
      '6. CALENDERING — steel roller stack; smooths, glosses,',
      '   ensures uniform thickness',
      '',
      '7. POPE REEL + REWINDER — winds into jumbo rolls;',
      '   trimmed to sellable widths; shipped as finished product',
    ],
  },
  {
    title: 'Key Equipment',
    bullets: [
      'RAW MATERIAL HANDLING:',
      '  Roller Wood Debarker — 5–18 T/H, 95% peeling rate',
      '  Wood Chipper — uniform chip production',
      '  Bale Breaker — 100–900 TPD, dual motor drive',
      '  Round Silo Storage — up to 42m diameter',
      '',
      'PULPING:',
      '  Hydrapulper — 1–102 m3, concentration 2–15%',
      '  Displacement Pulp Digester — 110–400 m3, 0.9–1.2 MPa',
      '  Double Disc Refiner — max motor 315–1,800 kW',
      '',
      'PAPERMAKING MACHINE:',
      '  Headbox / Wire / Press / Dryer / Size Press / Calendar',
      '  Trim width: up to 6,600 mm',
      '  Speed: up to 2,200 m/min | Capacity: up to 1,000 TPD',
      '  GSM range: 13–500 g/m2',
    ],
  },
  {
    title: 'Manpower — 80 Total Personnel',
    bullets: [
      'Top Management .............. 4',
      'Supply Chain & Logistics ..... 11',
      'Operating Personnel (Mill) ... 19',
      'Quality Control & Assurance .. 4',
      'Maintenance & Engineering .... 8',
      'Packaging ................... 9',
      'EHS & Security .............. 6',
      'Sales & Marketing ........... 7',
      'Human Resource Admin ........ 7',
      'Finance & Accounting ........ 5',
      '─────────────────────────────',
      'TOTAL ....................... 80',
      '',
      'Priority hiring: Local residents of SJDM and Bulacan Province',
    ],
  },
  {
    title: 'Industry Context — Existing Paper Mills',
    bullets: [
      'Major paper mills in Central Luzon:',
      '',
      'United Pulp & Paper Corp. — Calumpit, Bulacan',
      '  Products: Corrugating medium, test liner',
      '  Capacity: 230,000 MT/year',
      '',
      'Trust International Paper Corp. — Mabalacat, Pampanga',
      '  Products: Newsprint, printing & writing',
      '  Capacity: 230,000 MT/year',
      '',
      'Bataan 2020 Inc. — Samal, Bataan',
      '  Products: Printing/writing, newsprint, tissue',
      '  Capacity: 73,000 MT/year',
      '',
      'Container Corp. of the Philippines — Quezon City',
      '  Products: Corrugating medium, chipboard',
      '  Capacity: 89,000 MT/year',
    ],
  },
  {
    title: 'Design Computations',
    bullets: [
      'MASS BALANCE (Target: 1,200 kg paper/hr):',
      '  Dry fiber required .......... ~1,333 kg/hr',
      '  Water required .............. ~25,333 kg/hr',
      '  Total pulp slurry ........... ~26,667 kg/hr',
      '  Water evaporated at dryer ... 1,800 kg/hr',
      '',
      'PAPERMAKING CALCULATIONS:',
      '  Machine speed: 200 m/min | Deckle: 3.2 m | GSM: 50',
      '  Ream weight: 40cm x 60cm, 500 sheets at 50 GSM',
      '  Final dryness: 95% | Entering dryer dryness: 38%',
      '',
      'PIPING MATERIAL SELECTION:',
      '  Kraft / Sulfite pulp → Copper pipe',
      '  Mechanical / Groundwood → PVC pipe',
      '  Long-fiber kraft (never dried) → Stainless Steel / PVC',
      '  Soda pulp → Steel pipe',
    ],
  },
  {
    title: 'Conclusion',
    bullets: [
      'REAM-R-KABLE FIBERS — key takeaways:',
      '',
      '✓ Location: SJDM, Bulacan — strategic, safe, accessible',
      '✓ Raw Material: Recycled wastepaper (primary) + imported pulp',
      '✓ Three Pulping Methods: Kraft, Mechanical, Wastepaper',
      '✓ Five Design Constraints addressed: Environmental, Raw Material,',
      '  Energy, Technical, and Site',
      '✓ Total Workforce: 80 personnel',
      '  Priority hiring: local community of SJDM and Bulacan',
      '✓ Hazard-safe: low risk for flood, surge, liquefaction, landslide',
      '',
      'VISION:',
      'Premier, most technologically advanced, and environmentally',
      'remarkable pulp and paper mill in Southeast Asia.',
    ],
  },
  {
    title: 'References',
    bullets: [
      'AGICO Paper. (n.d.). Pulp mill machinery & paper making equipment.',
      '  https://agicopaper.com/',
      '',
      'Andritz. (n.d.). KPR paper capabilities brochure.',
      '',
      'Bureau of Energy Efficiency. (n.d.).',
      '  Best practice manual: Fluid piping systems.',
      '',
      'Golden Paper Group. (n.d.). Raw materials for paper making guide.',
      '',
      'Paper Mill Machine. (n.d.). Pulp mill section equipment catalogue.',
      '  https://papermillmachine.com/',
      '',
      'Provincial Government of Bulacan. (2010). PDPFP 2010–2030.',
      '',
      'Sarkar, S. (n.d.). Headbox calculations. Scribd.',
      '',
      'DAV University. (n.d.). MEC250-POM: Plant location and layout.',
    ],
  },
];



// ─────────────────────────────────────────────
// OPEN XML HELPERS
// ─────────────────────────────────────────────

function emu(pt) { return Math.round(pt * 12700); } // points → EMU
const W = 9144000;   // slide width  (10 inches) in EMU
const H = 5143500;   // slide height (5.625 inches) in EMU

// Escape XML special chars
function xe(s) {
  return String(s)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;');
}

// Build a text run
function txRun(text, opts = {}) {
  const bold  = opts.bold  ? '<a:b/>' : '';
  const sz    = opts.sz    ? `<a:sz val="${opts.sz}"/>` : '';
  const color = opts.color ? `<a:solidFill><a:srgbClr val="${opts.color}"/></a:solidFill>` : '';
  return `<a:r><a:rPr lang="en-US" dirty="0">${bold}${sz}${color}</a:rPr><a:t>${xe(text)}</a:t></a:r>`;
}

// Build a paragraph with optional bullet
function para(text, opts = {}) {
  const indent = opts.indent || 0;
  const buNone = opts.noBullet ? '<a:buNone/>' : '';
  const algn   = opts.center  ? ' algn="ctr"' : '';
  const spcBef = opts.spcBef  ? `<a:spcBef><a:spcPts val="${opts.spcBef}"/></a:spcBef>` : '';
  return `<a:p>
      <a:pPr marL="${emu(indent * 18)}" indent="0"${algn}>
        ${buNone}${spcBef}
      </a:pPr>
      ${txRun(text, opts)}
    </a:p>`;
}

// Build the slide XML
function buildSlide(slideIdx, titleText, bulletLines) {
  const titleParas = para(titleText, { bold: true, sz: 2800, color: '1F3864', noBullet: true, center: true });

  const contentParas = bulletLines.map(line => {
    if (line === '') return para(' ', { noBullet: true, sz: 1000 });
    if (line.startsWith('  ') || line.startsWith('   ')) {
      return para(line.trim(), { sz: 1400, indent: 2 });
    }
    if (line.startsWith('•') || line.startsWith('✓')) {
      return para(line, { sz: 1500, indent: 1 });
    }
    if (line.endsWith(':') || line.match(/^[A-Z][A-Z &]+:/)) {
      return para(line, { bold: true, sz: 1600, noBullet: true, color: '2E75B6' });
    }
    return para(line, { sz: 1500 });
  }).join('\n    ');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
       xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
       xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm>
      </p:grpSpPr>

      <!-- BACKGROUND -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="bg"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="EBF3FB"/></a:solidFill>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>

      <!-- TITLE BAR -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="titlebar"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${emu(72)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="1F3864"/></a:solidFill>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>
      </p:sp>

      <!-- TITLE TEXT -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${emu(18)}" y="${emu(4)}"/><a:ext cx="${W - emu(36)}" cy="${emu(62)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          ${para(titleText, { bold: true, sz: 2200, color: 'FFFFFF', noBullet: true, center: true })}
        </p:txBody>
      </p:sp>

      <!-- SLIDE NUMBER -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="slidenum"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${W - emu(60)}" y="${emu(4)}"/><a:ext cx="${emu(55)}" cy="${emu(62)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          ${para(`${slideIdx + 1} / ${slides.length}`, { sz: 1200, color: 'AAAAAA', noBullet: true, center: true })}
        </p:txBody>
      </p:sp>

      <!-- CONTENT -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="6" name="content"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="${emu(30)}" y="${emu(84)}"/><a:ext cx="${W - emu(60)}" cy="${H - emu(104)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square" lIns="${emu(4)}" rIns="${emu(4)}" tIns="${emu(4)}" bIns="${emu(4)}" anchor="t"/>
          <a:lstStyle/>
          ${contentParas}
        </p:txBody>
      </p:sp>

      <!-- FOOTER BAR -->
      <p:sp>
        <p:nvSpPr><p:cNvPr id="7" name="footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="0" y="${H - emu(22)}"/><a:ext cx="${W}" cy="${emu(22)}"/></a:xfrm>
          <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
          <a:solidFill><a:srgbClr val="2E75B6"/></a:solidFill>
          <a:ln><a:noFill/></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:lstStyle/>
          ${para('REAM-R-KABLE FIBERS  |  Paper and Pulp Industry  |  BSME 4A  |  A.Y. 2025–2026', { sz: 1000, color: 'FFFFFF', noBullet: true, center: true })}
        </p:txBody>
      </p:sp>

    </p:spTree>
  </p:cSld>
</p:sld>`;
}



// ─────────────────────────────────────────────
// OPEN XML PACKAGE FILES
// ─────────────────────────────────────────────

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
${slides.map((_, i) => `  <Override PartName="/ppt/slides/slide${i + 1}.xml"\n    ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`).join('\n')}
</Types>`;

const ROOT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`;

const PPT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
${slides.map((_, i) => `  <Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${i + 1}.xml"/>`).join('\n')}
</Relationships>`;

const SLIDE_MASTER = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;

const SLIDE_MASTER_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;

const SLIDE_LAYOUT = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
    <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${W}" cy="${H}"/><a:chOff x="0" y="0"/><a:chExt cx="${W}" cy="${H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
</p:sldLayout>`;

const SLIDE_LAYOUT_RELS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;

function buildPresentation() {
  const slideIdList = slides.map((_, i) =>
    `    <p:sldId id="${256 + i}" r:id="rId${i + 1}"/>`
  ).join('\n');

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId0"/></p:sldMasterIdLst>
  <p:sldIdLst>
${slideIdList}
  </p:sldIdLst>
  <p:sldSz cx="${W}" cy="${H}" type="screen16x9"/>
  <p:notesSz cx="${H}" cy="${W}"/>
</p:presentation>`;
}

function buildSlideRels(slideIdx) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
}



// ─────────────────────────────────────────────
// BUILD PPTX (ZIP via Python3)
// ─────────────────────────────────────────────

function buildPptx(outputPath) {
  const tmpDir = fs.mkdtempSync(path.join(os.tmpdir(), 'pptx-'));

  const write = (rel, content) => {
    const full = path.join(tmpDir, rel);
    fs.mkdirSync(path.dirname(full), { recursive: true });
    fs.writeFileSync(full, content, 'utf8');
  };

  // Package structure
  write('[Content_Types].xml', CONTENT_TYPES);
  write('_rels/.rels', ROOT_RELS);
  write('ppt/presentation.xml', buildPresentation());
  write('ppt/_rels/presentation.xml.rels', PPT_RELS);
  write('ppt/slideMasters/slideMaster1.xml', SLIDE_MASTER);
  write('ppt/slideMasters/_rels/slideMaster1.xml.rels', SLIDE_MASTER_RELS);
  write('ppt/slideLayouts/slideLayout1.xml', SLIDE_LAYOUT);
  write('ppt/slideLayouts/_rels/slideLayout1.xml.rels', SLIDE_LAYOUT_RELS);

  // All slides
  slides.forEach((slide, i) => {
    write(`ppt/slides/slide${i + 1}.xml`, buildSlide(i, slide.title, slide.bullets));
    write(`ppt/slides/_rels/slide${i + 1}.xml.rels`, buildSlideRels(i));
  });

  // Use Python3 to zip into .pptx
  const absOut = path.resolve(outputPath);
  const pyScript = `
import zipfile, os, sys

tmpDir = sys.argv[1]
outFile = sys.argv[2]

with zipfile.ZipFile(outFile, 'w', zipfile.ZIP_DEFLATED) as zf:
    for root, dirs, files in os.walk(tmpDir):
        for file in files:
            fullPath = os.path.join(root, file)
            arcName = os.path.relpath(fullPath, tmpDir)
            zf.write(fullPath, arcName)
print("Created:", outFile)
`;

  const pyFile = path.join(os.tmpdir(), 'make_pptx.py');
  fs.writeFileSync(pyFile, pyScript);

  try {
    const result = execSync(`python3 "${pyFile}" "${tmpDir}" "${absOut}"`, { encoding: 'utf8' });
    console.log(result.trim());
  } finally {
    // Cleanup temp files
    fs.rmSync(tmpDir, { recursive: true, force: true });
    fs.unlinkSync(pyFile);
  }
}

// ─────────────────────────────────────────────
// MAIN
// ─────────────────────────────────────────────

const outputFile = process.argv[2] || 'REAM-R-KABLE-FIBERS.pptx';
console.log(`Generating ${slides.length} slides...`);
buildPptx(outputFile);
console.log(`\nDone! Open: ${path.resolve(outputFile)}`);
console.log('Compatible with: Microsoft PowerPoint, Google Slides, LibreOffice Impress');
