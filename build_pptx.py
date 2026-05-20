#!/usr/bin/env python3
"""
REAM-R-KABLE FIBERS - Clean PPTX Builder
Uses minimal valid Open XML structure that PowerPoint can reliably read.
Run: python3 build_pptx.py
Output: REAM-R-KABLE-FIBERS.pptx
"""
import zipfile, os
from zipfile import ZIP_DEFLATED, ZIP_STORED

OUTPUT = 'REAM-R-KABLE-FIBERS.pptx'

# Slide dimensions: 10 x 7.5 inches in EMU (914400 EMU per inch)
W = 9144000
H = 6858000

# Colors
NAVY   = '1A3A5C'
TEAL   = '0D7377'
GOLD   = 'F5A623'
WHITE  = 'FFFFFF'
BGCLR  = 'F0F4F8'
DARK   = '2C3E50'
MID    = '5D6D7E'
LIGHT  = 'D6E8F5'

def p(n): return n * 12700  # points to EMU

def xe(s):
    """Escape XML, strip non-XML-safe chars"""
    s = str(s)
    # Remove chars not valid in XML 1.0
    s = ''.join(c for c in s if ord(c) >= 0x20 or c in '\t\n\r')
    s = s.replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')
    # Replace special unicode that may cause issues with safe ASCII equivalents
    s = s.replace('\u2192','->').replace('\u2794','->').replace('\u2714','v')
    s = s.replace('\u2192','->').replace('\u2503','|').replace('\u2022','*')
    s = s.replace('\u2013','-').replace('\u2014','--').replace('\u2500','-')
    s = s.replace('\u2192','->').replace('\u00b3','3').replace('\u00b2','2')
    s = s.replace('\u2192','->').replace('\u25bc','v').replace('\u25ba','>')
    return s



def run(text, sz=1400, color=DARK, bold=False):
    b = '<a:b/>' if bold else ''
    return (
        f'<a:r>'
        f'<a:rPr lang="en-US" dirty="0" b="{1 if bold else 0}">'
        f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>'
        f'<a:latin typeface="Calibri"/>'
        f'<a:cs typeface="Calibri"/>'
        f'<a:sz val="{sz}"/>'
        f'{b}'
        f'</a:rPr>'
        f'<a:t>{xe(text)}</a:t>'
        f'</a:r>'
    )

def para(runs_xml, algn='l', marL=0, buNone=True, spcBef=0):
    bu = '<a:buNone/>' if buNone else ''
    sp = f'<a:spcBef><a:spcPts val="{spcBef}"/></a:spcBef>' if spcBef else ''
    return (
        f'<a:p>'
        f'<a:pPr algn="{algn}" marL="{marL}" indent="0">{bu}{sp}</a:pPr>'
        f'{runs_xml}'
        f'</a:p>'
    )

def rect_shape(x, y, cx, cy, fill):
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="0" name="r"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        f'<a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>'
        f'</p:sp>'
    )

def text_shape(x, y, cx, cy, paragraphs_xml, anchor='t'):
    return (
        f'<p:sp><p:nvSpPr>'
        f'<p:cNvPr id="0" name="t"/>'
        f'<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
        f'<p:nvPr/></p:nvSpPr>'
        f'<p:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:noFill/><a:ln><a:noFill/></a:ln>'
        f'</p:spPr>'
        f'<p:txBody>'
        f'<a:bodyPr wrap="square" anchor="{anchor}" lIns="{p(6)}" rIns="{p(4)}" tIns="{p(4)}" bIns="{p(4)}"/>'
        f'<a:lstStyle/>'
        f'{paragraphs_xml}'
        f'</p:txBody>'
        f'</p:sp>'
    )



# ── SLIDE DATA ────────────────────────────────────────────────────────────────
SLIDES = [
  {'title':'REAM-R-KABLE FIBERS','sub':'Paper and Pulp Industry | Industrial Plant Engineering PROFME 105','is_title':True,'bullets':[]},
  {'title':'Company Profile','sub':'About REAM-R-KABLE FIBERS','bullets':[
    'Location: Barangay Ciudad Real, San Jose del Monte, Bulacan',
    '"REAM" (unit of paper) + "FIBER" (foundation of product) = REMARKABLE',
    '','VISION:','To be the premier, most technologically advanced and environmentally',
    'remarkable pulp and paper mill in Southeast Asia.','','MISSION:',
    '* Deliver Excellence -- high-grade paper exceeding global standards',
    '* Champion Sustainability -- water recycling; wastepaper as raw material',
    '* Empower Community -- high-value technical jobs in Bulacan',
    '* Drive Innovation -- modern engineering in pulping and refining']},
  {'title':'Introduction','sub':'Paper, Pulp & the Philippine Industry','bullets':[
    'PAPER -- wood fibers reduced to pulp, formed into a mat, compressed, dried',
    'PULP  -- fibrous material from wood/non-wood via chemical or mechanical process',
    '','Paper Grades:',
    '* Printing & Writing  (bond paper, notebooks)',
    '* Newsprint  (newspapers)',
    '* Corrugating / Containerboard  (linerboard, fluting)',
    '* Tissue & Sanitary Papers','* Specialty (art paper, construction board)',
    '','Philippine Context:',
    '* PICOP Resources Inc. (only integrated mill) ceased operations in 2010',
    '* Industry now relies PRIMARILY on recycled wastepaper',
    '* Wastepaper = 19% of total municipal solid waste in PH']},
  {'title':'Plant Location','sub':'San Jose Del Monte (SJDM), Bulacan','bullets':[
    '~42 km from Metro Manila  |  Landlocked component city',
    '3 national roads | 4 provincial roads | 402 city roads | 15 bridges',
    '','Why SJDM?',
    '* Proximity to Metro Manila -- highest wastepaper source density',
    '* Near Region 3 (Central Luzon) -- rice straw, bagasse, bamboo supply',
    '* Access to NAIA, Clark, Subic airports and major seaports',
    '* Inside established industrial cluster of Central Luzon',
    '* Bulacan Province: 279,610 ha  |  Cities: Malolos, Meycauayan, SJDM']},
  {'title':'Hazard Assessment','sub':'Why SJDM is Safe for Industrial Investment','bullets':[
    'FLOODING  ->  LOW RISK',
    '  Eastern location; away from coastal lowlands and river floodplains',
    '','STORM SURGE  ->  ZERO RISK',
    '  Inland position; surge risk confined to coastal municipalities only',
    '','LIQUEFACTION  ->  ZERO RISK',
    '  Stable bedrock / thick soil -- WHITE ZONE on hazard map',
    '','LANDSLIDE (Earthquake)  ->  VERY LOW RISK',
    '  Urbanized flatland; away from Sierra Madre unstable slopes',
    '','SJDM falls in white/yellow zones across ALL FOUR hazard maps']},
  {'title':'Raw Materials','sub':'Industry Inputs','bullets':[
    'A.  VIRGIN / WOOD PULP',
    '  Softwood (Pine, Spruce) -- long fibers -- strength for printing/writing paper',
    '  Hardwood (Eucalyptus, Birch) -- short fibers -- improves bulk and opacity',
    '','B.  NON-WOOD PULP',
    '  Sources: Abaca, banana, rice straw, wheat straw, bagasse, bamboo',
    '  Abaca pulp -- exceptional strength; key Philippine export commodity',
    '','C.  RECYCLED WASTEPAPER  (Primary Source in Philippines)',
    '  Processed via deinking + fiber separation',
    '  Reduces waste, conserves resources, lowers GHG emissions',
    '  Supplemented by imported virgin pulp when quality demands it']},
  {'title':'Chemicals & Additives','sub':'Key Chemicals Used in Production','bullets':[
    'PULPING CHEMICALS:',
    '  Caustic Soda (NaOH) -- breaks down wood chips into pulp',
    '  Sodium Sulfide (Na2S) -- Kraft cooking liquor; dissolves lignin',
    '','BLEACHING AGENTS:',
    '  Chlorine Dioxide (ClO2) -- high-level whiteness bleaching',
    '  Hydrogen Peroxide (H2O2) -- eco-friendly brightening agent',
    '  Sodium Hypochlorite -- wastepaper pulp brightness improvement',
    '  Sodium Hydrosulfite -- mechanical pulp brightening',
    '','PAPERMAKING ADDITIVES:',
    '  Fillers (clay, CaCO3) -- opacity, brightness, printability',
    '  Sizing agents (starch, rosin) -- ink/water resistance',
    '  Retention aids, strength agents, slime control, antifoam, dyes']},
  {'title':'5 Design Constraints','sub':'Engineering Boundaries of the Plant','bullets':[
    '1.  ENVIRONMENTAL & REGULATORY',
    '    Effluent: BOD, COD, TSS, pH strictly controlled before discharge',
    '    Air: scrubbers for total reduced sulfur and particulate matter',
    '','2.  RAW MATERIAL & FIBER',
    '    De-inking and contaminant-removal systems for recycled fiber',
    '    Large chip/bale storage areas for uninterrupted supply',
    '','3.  ENERGY & UTILITY',
    '    Co-generation (CHP) target -- plant aims to be energy self-sufficient',
    '','4.  TECHNICAL & PROCESS',
    '    Kraft pulping + ECF/TCF bleaching (no elemental chlorine)',
    '    Duplex stainless steel / titanium for corrosion resistance',
    '','5.  SITE & GEOGRAPHICAL',
    '    Near water source for intake/discharge; proximity to ports & highways']},
  {'title':'Process Flow Diagram','sub':'End-to-End Manufacturing Process','bullets':[
    'RAW MATERIAL PREPARATION:',
    '  Debarking -> Log Cutting -> Chipping -> Screening -> Storage',
    '','PULPING (choose method based on raw material):',
    '  Chemical Pulping (Kraft) | Mechanical Pulping | Wastepaper Pulping',
    '','WASHING -> SCREENING -> CLEANING  ->  BLEACHING',
    '','STOCK / PULP PREPARATION:',
    '  Beating -> Blending -> Fillers + Sizing + Chemicals',
    '','PAPERMAKING STAGE:',
    '  Headbox -> Wire -> Press -> Dryer -> Size Press -> Calendar -> Reel',
    '','FINISHED PRODUCT  (Jumbo rolls, trimmed, shipped)']},
  {'title':'Pulping Methods','sub':'Three Methods Used in the Plant','bullets':[
    'METHOD 1 -- CHEMICAL PULPING (KRAFT PROCESS):',
    '  Wood chips + White Liquor (NaOH + Na2S) in digester at ~160 deg C / 10 atm',
    '  Cellulose fibers separated from lignin; black liquor -> recovery boiler',
    '','METHOD 2 -- MECHANICAL PULPING:',
    '  Twin rotating disk refiner physically grinds chips into fibers',
    '  Higher yield, lower strength; retains most lignin',
    '  Bleached with sodium hydrosulfite or hydrogen peroxide',
    '','METHOD 3 -- WASTEPAPER PULPING:',
    '  Hydrapulper: 3-5% concentration, 30-60 min mechanical agitation',
    '  Foreign objects removed by screens and cleaners',
    '  Deinking: ink removed using alkali + surface-active agents',
    '  Bleaching to achieve 55-75% target brightness']},
  {'title':'Stock Preparation','sub':'Preparing Pulp Before Papermaking','bullets':[
    '1.  BEATING:',
    '    Disk refiner swells fibers; improves flexibility and surface area',
    '    More inter-fiber bonding points = stronger sheet formation',
    '','2.  BLENDING:',
    '    Recycles broke (trim and defective paper from production)',
    '    Blends raw materials + recovered white water from wire section',
    '','3.  ADDITION OF FILLERS, SIZING AGENTS & CHEMICALS:',
    '    Fillers (clay, CaCO3) -- opacity, brightness, smoothness',
    '    Sizing agents (starch) -- ink and water resistance',
    '    Retention aids -- reduce fiber loss in wire section',
    '    Strengthening agents -- improve wet and dry sheet strength',
    '    Slime control / antifoaming / dyes -- added as process requires']},
  {'title':'Papermaking Stage','sub':'From Pulp Slurry to Finished Paper','bullets':[
    '1.  HEADBOX -- distributes fiber-water slurry uniformly across wire width',
    '2.  WIRE SECTION -- fibers settle into mat; water drains away; sheet forms',
    '3.  PRESS SECTION -- rollers + felt cloth; mechanical dewatering;',
    '    increases strength and improves surface quality',
    '4.  DRYER SECTION -- steam-heated cylinders; reduces moisture to 6-10%',
    '5.  SIZE PRESS -- applies starch for water resistance + surface strength',
    '6.  CALENDERING -- steel roller stack; smooths surface, adds gloss,',
    '    ensures uniform sheet thickness',
    '7.  POPE REEL -- winds paper into large jumbo rolls',
    '8.  REWINDER -- trims and cuts jumbo rolls to sellable widths',
    '    removes defective sections; final product shipped']},
  {'title':'Key Equipment','sub':'Major Machinery & Specifications','bullets':[
    'RAW MATERIAL HANDLING:',
    '  Roller Wood Debarker  5-18 T/H  |  95% peeling rate',
    '  Bale Breaker  100-900 TPD  |  Dual motor drive',
    '  Round Silo Chip Storage  Up to 42 m diameter',
    '','PULPING:',
    '  Hydrapulper  1-102 m3  |  Concentration 2-15%',
    '  Displacement Pulp Digester  110-400 m3  |  0.9-1.2 MPa',
    '  Double Disc Refiner (PM20)  Max motor 315-1,800 kW',
    '  Flotation Deinking Cell  Ink removal rate 0.25-0.3; 15-150 T/D',
    '','PAPERMAKING MACHINE:',
    '  Headbox / Wire / Press / Dryer / Size Press / Calendar / Reel',
    '  Trim width: up to 6,600 mm  |  Speed: up to 2,200 m/min',
    '  Capacity: up to 1,000 TPD  |  GSM range: 13-500 g/m2']},
  {'title':'Manpower','sub':'80 Total Personnel -- Priority Hiring from SJDM & Bulacan','bullets':[
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
    '------------------------------',
    'TOTAL .........................  80',
    '','Priority hiring: Local residents of SJDM and Bulacan Province']},
  {'title':'Industry Context','sub':'Existing Major Paper Mills -- Central Luzon','bullets':[
    'United Pulp & Paper Corp. -- Calumpit, Bulacan',
    '  Products: Corrugating medium, test liner  |  Capacity: 230,000 MT/year',
    '','Trust International Paper Corp. -- Mabalacat, Pampanga',
    '  Products: Newsprint, printing & writing  |  Capacity: 230,000 MT/year',
    '','Bataan 2020 Inc. -- Samal, Bataan',
    '  Products: Printing/writing, newsprint, tissue  |  Capacity: 73,000 MT/year',
    '','Container Corp. of the Philippines -- Quezon City',
    '  Products: Corrugating medium, chipboard  |  Capacity: 89,000 MT/year',
    '','SJDM sits at the center of this cluster -- shared labor, suppliers & logistics']},
  {'title':'Design Computations','sub':'Key Calculations Summary','bullets':[
    'MASS BALANCE (Target: 1,200 kg finished paper / hour):',
    '  Dry fiber required       ~1,333 kg/hr',
    '  Process water required   ~25,333 kg/hr',
    '  Total pulp slurry input  ~26,667 kg/hr',
    '  Water evaporated at dryer  1,800 kg/hr',
    '','PAPERMAKING MACHINE PARAMETERS:',
    '  Machine speed: 200 m/min  |  Deckle: 3.2 m  |  GSM: 50',
    '  Final dryness: 95%  |  Dryer inlet dryness: 38%',
    '  Wire length: ~27.09 m  |  Size press: max 250 rpm',
    '','PIPING MATERIAL SELECTION:',
    '  Kraft / Sulfite pulp         ->  Copper pipe',
    '  Mechanical / Groundwood      ->  PVC pipe',
    '  Long-fiber kraft (never dried) ->  Stainless Steel / PVC',
    '  Soda pulp                    ->  Steel pipe']},
  {'title':'Conclusion','sub':'Summary of Key Findings','bullets':[
    'v  Location: SJDM, Bulacan -- strategic, low-hazard, logistically superior',
    'v  Raw Material: Recycled wastepaper (primary) + imported virgin pulp',
    'v  Three Pulping Methods: Kraft Chemical, Mechanical, Wastepaper',
    'v  Five Design Constraints fully addressed:',
    '     Environmental, Raw Material, Energy, Technical, Site',
    'v  Total Workforce: 80 personnel',
    '     Priority hiring: local community of SJDM and Bulacan Province',
    'v  Natural hazard safe: flood, storm surge, liquefaction, landslide',
    '','OUR VISION:',
    '  To be the premier, most technologically advanced, and environmentally',
    '  REMARKABLE pulp and paper mill in Southeast Asia.']},
  {'title':'References','sub':'Sources & Citations','bullets':[
    'AGICO Paper. (n.d.). Pulp mill machinery & paper making equipment.',
    '  https://agicopaper.com/',
    'Andritz. (n.d.). KPR paper capabilities brochure.',
    'Bureau of Energy Efficiency. (n.d.). Best practice manual: Fluid piping systems.',
    'Golden Paper Group. (n.d.). Raw materials for paper making guide.',
    'Paper Mill Machine. (n.d.). Pulp mill section equipment catalogue.',
    '  https://papermillmachine.com/',
    'Provincial Government of Bulacan. (2010). PDPFP 2010-2030.',
    'Sarkar, S. (n.d.). Headbox calculations. Scribd.',
    'DAV University. (n.d.). MEC250-POM: Plant location and layout.',
    'JSS College of Arts. (2023). Department file -- Paper Industry Overview.',
    'Landmark University. (n.d.). GEC524: Plant Layout.']},
]



# ── BUILD SLIDES ──────────────────────────────────────────────────────────────

NS = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'

def build_title_slide(slide, idx):
    lw = int(W * 0.40)
    rx = lw + p(5)
    rw = W - rx
    shapes = []
    shapes.append(rect_shape(0,0,W,H,BGCLR))
    shapes.append(rect_shape(0,0,lw,H,NAVY))
    shapes.append(rect_shape(lw,0,p(7),H,GOLD))
    shapes.append(rect_shape(0,H-p(16),W,p(16),TEAL))
    # Big title on left
    shapes.append(text_shape(p(22),p(55),lw-p(44),p(90),
        para(run('REAM-R-KABLE',4800,GOLD,True),'l',0,True),'ctr'))
    shapes.append(text_shape(p(22),p(148),lw-p(44),p(55),
        para(run('FIBERS',3600,WHITE,True),'l',0,True),'ctr'))
    shapes.append(rect_shape(p(22),p(220),lw-p(44),p(3),TEAL))
    shapes.append(text_shape(p(22),p(230),lw-p(44),p(30),
        para(run('Paper and Pulp Industry',1500,LIGHT),'l',0,True),'ctr'))
    shapes.append(text_shape(p(22),p(262),lw-p(44),p(25),
        para(run('Industrial Plant Engineering',1300,MID),'l',0,True),'ctr'))
    shapes.append(text_shape(p(22),p(288),lw-p(44),p(22),
        para(run('PROFME 105',1300,MID),'l',0,True),'ctr'))
    # Right panel
    shapes.append(rect_shape(rx,0,rw,p(7),TEAL))
    shapes.append(text_shape(rx+p(28),p(22),rw-p(56),p(40),
        para(run('Pangasinan State University',1700,NAVY,True),'l',0,True),'ctr'))
    shapes.append(text_shape(rx+p(28),p(64),rw-p(56),p(25),
        para(run('Urdaneta City Campus  |  College of Engineering and Architecture',1200,MID),'l',0,True),'ctr'))
    shapes.append(rect_shape(rx+p(28),p(94),rw-p(56),p(2),TEAL))
    pinfo = (
        para(run('Prepared by:',1200,TEAL,True),'l',0,True,200) +
        para(run('DELA CRUZ, JHAN RAYVEN D.',1900,NAVY,True),'l',0,True) +
        para(run('BSME 4A',1400,MID),'l',0,True,100) +
        para(run(''),'l',0,True) +
        para(run('Submitted to:',1200,TEAL,True),'l',0,True,200) +
        para(run('ENGR. MARFEL D. ROSARIO',1600,NAVY,True),'l',0,True) +
        para(run('Instructor',1300,MID),'l',0,True)
    )
    shapes.append(text_shape(rx+p(28),p(102),rw-p(56),p(170),pinfo,'t'))
    shapes.append(text_shape(rx+p(28),H-p(46),rw-p(56),p(30),
        para(run(f'A.Y. 2025-2026  |  May 2026  |  Slide {idx+1}/{len(SLIDES)}',1200,WHITE),'l',0,True),'ctr'))
    return make_slide_xml(shapes)

def build_content_slide(slide, idx):
    hh = p(65)
    ah = p(6)
    sh = p(26)
    fh = p(22)
    sw = p(10)
    px = p(30)
    by = hh + ah + sh + p(10)
    bh = H - by - fh - p(8)
    shapes = []
    shapes.append(rect_shape(0,0,W,H,BGCLR))
    shapes.append(rect_shape(0,0,W,hh,NAVY))
    shapes.append(rect_shape(0,hh,W,ah,GOLD))
    shapes.append(rect_shape(0,hh+ah,W,sh,TEAL))
    shapes.append(rect_shape(0,by,sw,bh,TEAL))
    shapes.append(rect_shape(0,H-fh,W,fh,NAVY))
    # Title
    shapes.append(text_shape(px,p(8),W-px-p(85),hh-p(10),
        para(run(slide['title'],2700,WHITE,True),'l',0,True),'ctr'))
    # Slide num
    shapes.append(text_shape(W-p(82),p(8),p(76),hh-p(10),
        para(run(f'{idx+1}/{len(SLIDES)}',1300,GOLD,True),'r',0,True),'ctr'))
    # Subtitle
    shapes.append(text_shape(px,hh+ah+p(3),W-px*2,sh-p(3),
        para(run(slide.get('sub',''),1200,WHITE),'l',0,True),'ctr'))
    # Body
    body_paras = []
    for line in slide['bullets']:
        if line == '':
            body_paras.append(para(run(' ',700,DARK),'l',0,True))
        elif line.endswith(':') and line == line.upper().replace(' ','').replace(':','').upper()+':':
            body_paras.append(para(run(line,1500,TEAL,True),'l',p(6),True,150))
        elif line.endswith(':') and len(line) < 50:
            body_paras.append(para(run(line,1500,TEAL,True),'l',p(6),True,150))
        elif line.startswith('  ') or line.startswith('    '):
            body_paras.append(para(run(line.strip(),1250,MID),'l',p(36),True,30))
        elif line.startswith('* ') or line.startswith('v '):
            body_paras.append(para(run(line,1350,DARK),'l',p(20),True,50))
        elif line[:2].strip() and line[0].isdigit() and '.' in line[:3]:
            body_paras.append(para(run(line,1400,NAVY,True),'l',p(6),True,120))
        elif line.startswith('-----'):
            body_paras.append(para(run(line,900,MID),'l',0,True))
        else:
            body_paras.append(para(run(line,1350,DARK),'l',p(6),True,60))
    shapes.append(text_shape(sw+p(14),by+p(8),W-sw-p(24),bh-p(8),''.join(body_paras),'t'))
    # Footer
    shapes.append(text_shape(px,H-fh+p(3),W-px*2,fh-p(3),
        para(run('REAM-R-KABLE FIBERS  |  Paper and Pulp Industry  |  BSME 4A  |  A.Y. 2025-2026',950,LIGHT),'l',0,True),'ctr'))
    return make_slide_xml(shapes)

def make_slide_xml(shapes):
    grp = f'<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>'
    nvg = f'<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
    tree = f'<p:spTree>{nvg}{grp}{"".join(shapes)}</p:spTree>'
    return f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld {NS}><p:cSld>{tree}</p:cSld></p:sld>'

def build_slide(slide, idx):
    if slide.get('is_title'):
        return build_title_slide(slide, idx)
    return build_content_slide(slide, idx)



# ── OPEN XML PACKAGE PARTS ────────────────────────────────────────────────────

CONTENT_TYPES = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
  <Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
''' + '\n'.join(f'  <Override PartName="/ppt/slides/slide{i+1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>' for i in range(len(SLIDES))) + '''
</Types>'''

ROOT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>'''

PPT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
''' + '\n'.join(f'  <Relationship Id="rId{i+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{i+1}.xml"/>' for i in range(len(SLIDES))) + '''
</Relationships>'''

SLIDE_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

SLIDE_MASTER = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
  <p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>'''

MASTER_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>'''

SLIDE_LAYOUT = f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" type="blank" preserve="1">
  <p:cSld name="Blank"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{W}" cy="{H}"/><a:chOff x="0" y="0"/><a:chExt cx="{W}" cy="{H}"/></a:xfrm></p:grpSpPr>
  </p:spTree></p:cSld>
</p:sldLayout>'''

LAYOUT_RELS = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>'''

def build_presentation():
    slide_ids = '\n'.join(f'    <p:sldId id="{256+i}" r:id="rId{i+1}"/>' for i in range(len(SLIDES)))
    return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" saveSubsetFonts="1">
  <p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId0"/></p:sldMasterIdLst>
  <p:sldIdLst>
{slide_ids}
  </p:sldIdLst>
  <p:sldSz cx="{W}" cy="{H}" type="screen16x9"/>
  <p:notesSz cx="{H}" cy="{W}"/>
</p:presentation>'''

# ── WRITE PPTX ────────────────────────────────────────────────────────────────

def write_pptx(filename):
    files = {}
    files['[Content_Types].xml']                          = CONTENT_TYPES
    files['_rels/.rels']                                   = ROOT_RELS
    files['ppt/presentation.xml']                         = build_presentation()
    files['ppt/_rels/presentation.xml.rels']              = PPT_RELS
    files['ppt/slideMasters/slideMaster1.xml']            = SLIDE_MASTER
    files['ppt/slideMasters/_rels/slideMaster1.xml.rels'] = MASTER_RELS
    files['ppt/slideLayouts/slideLayout1.xml']            = SLIDE_LAYOUT
    files['ppt/slideLayouts/_rels/slideLayout1.xml.rels'] = LAYOUT_RELS
    for i, slide in enumerate(SLIDES):
        files[f'ppt/slides/slide{i+1}.xml']              = build_slide(slide, i)
        files[f'ppt/slides/_rels/slide{i+1}.xml.rels']  = SLIDE_RELS

    with zipfile.ZipFile(filename, 'w') as zf:
        # [Content_Types].xml MUST be first and STORED (OOXML spec)
        ct_info = zipfile.ZipInfo('[Content_Types].xml')
        ct_info.compress_type = ZIP_STORED
        zf.writestr(ct_info, files['[Content_Types].xml'].encode('utf-8'))
        # Everything else DEFLATED
        for name, content in files.items():
            if name == '[Content_Types].xml':
                continue
            info = zipfile.ZipInfo(name)
            info.compress_type = ZIP_DEFLATED
            zf.writestr(info, content.encode('utf-8'))

    size = os.path.getsize(filename)
    print(f'Done! {filename}  ({size:,} bytes, {len(SLIDES)} slides)')

if __name__ == '__main__':
    write_pptx(OUTPUT)
