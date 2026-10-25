import json, os, sys, copy
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
from PIL import Image

SC='/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/'
VIDDIR=os.environ.get('VIDDIR','vid'); OUT=os.environ.get('OUT','J-FARP_Seizure_HQ.pptx')
FR=SC+'frames/'; VD=SC+VIDDIR+'/'; AS=os.path.dirname(os.path.abspath(__file__))+'/assets/'
S=json.load(open(SC+'slides.json')); cap=lambda n:S[n-1]['cap']
BLACK=RGBColor(0,0,0); GREY=RGBColor(0xED,0xED,0xED); NAVY=RGBColor(0x1F,0x4E,0x79); MUTED=RGBColor(0x59,0x59,0x59); PURPLE=RGBColor(0x70,0x30,0xA0); RED=RGBColor(0xC0,0,0)

prs=Presentation(AS+'template_storm_surge.pptx')
# drop the template's own slides and its section list
sldIdLst=prs.slides._sldIdLst
for sldId in list(sldIdLst):
    prs.part.drop_rel(sldId.rId); sldIdLst.remove(sldId)
pres_el=prs.part._element
for ext in pres_el.iter():
    if ext.tag==qn('p:extLst'):
        pres_el.remove(ext); break
LAYOUT=[l for l in prs.slide_masters[0].slide_layouts if l.name=='2_Blank'][0]

def rgb(h): return RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))
def shadow(shape,blur=3,dist=2,alpha=35):
    spPr=shape._element.spPr
    eff=etree.SubElement(spPr,qn('a:effectLst'))
    sh=etree.SubElement(eff,qn('a:outerShdw'));sh.set('blurRad',str(int(blur*12700)));sh.set('dist',str(int(dist*12700)));sh.set('dir','2700000');sh.set('algn','tl');sh.set('rotWithShape','0')
    c=etree.SubElement(sh,qn('a:srgbClr'));c.set('val','000000');a=etree.SubElement(c,qn('a:alpha'));a.set('val',str(alpha*1000))
def rect(sl,x,y,w,h,fill='EDEDED',line='000000',lw=1,shadow_on=True,shape=MSO_SHAPE.RECTANGLE):
    s=sl.shapes.add_shape(shape,Inches(x),Inches(y),Inches(w),Inches(h))
    if fill is None: s.fill.background()
    else: s.fill.solid(); s.fill.fore_color.rgb=rgb(fill)
    if line is None: s.line.fill.background()
    else: s.line.color.rgb=rgb(line); s.line.width=Pt(lw)
    if shadow_on: shadow(s)
    s.text_frame.text=''
    return s
# paragraphs: list of dicts {runs:[(text,{b,u,i,size,color})], bullet, align, space_after}
def text(sl,x,y,w,h,paras,size=12,align='l',anchor='t',margin=0.0):
    tb=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=tb.text_frame;tf.word_wrap=True
    tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=Inches(margin)
    tf.vertical_anchor={'t':MSO_ANCHOR.TOP,'m':MSO_ANCHOR.MIDDLE,'b':MSO_ANCHOR.BOTTOM}[anchor]
    first=True
    for para in paras:
        p=tf.paragraphs[0] if first else tf.add_paragraph(); first=False
        p.alignment={'l':PP_ALIGN.LEFT,'c':PP_ALIGN.CENTER,'r':PP_ALIGN.RIGHT}[para.get('align',align)]
        if para.get('space_after'): p.space_after=Pt(para['space_after'])
        if para.get('bullet'):
            pPr=p._p.get_or_add_pPr();pPr.set('marL','228600');pPr.set('indent','-228600')
            bf=etree.SubElement(pPr,qn('a:buFont'));bf.set('typeface','Arial');bc=etree.SubElement(pPr,qn('a:buChar'));bc.set('char','•')
        for t,o in para['runs']:
            r=p.add_run();r.text=t;f=r.font;f.name='Arial';f.size=Pt(o.get('size',para.get('size',size)));f.bold=o.get('b',False);f.italic=o.get('i',False);f.underline=o.get('u',False)
            f.color.rgb=o.get('color',BLACK)
    return tb
def R(t,**o): return (t,o)
def P(*runs,**o): d={'runs':list(runs)}; d.update(o); return d
def title(sl,t):
    s=rect(sl,3.17,.22,7.0,.58,fill='F2F2F2',lw=1.5)
    tf=s.text_frame;tf.margin_left=tf.margin_right=Inches(.05);tf.vertical_anchor=MSO_ANCHOR.MIDDLE
    p=tf.paragraphs[0];p.alignment=PP_ALIGN.CENTER;r=p.add_run();r.text=t;r.font.name='Arial';r.font.size=Pt(24);r.font.bold=True;r.font.color.rgb=BLACK
def box(sl,x,y,w,h,label,paras,size=12,fill='EDEDED',anchor='t',align='l'):
    rect(sl,x,y,w,h,fill=fill)
    ps=[]
    if label: ps.append(P(R(label,b=True,u=True),space_after=4))
    ps+=paras
    text(sl,x+.12,y+.08,w-.24,h-.16,ps,size=size,anchor=anchor,align=align)
def para(s,**o): return P(R(s),**o)
def lead(label,rest,**o): return P(R(label+' ',b=True),R(rest),**o)
def bullets(items,size=None):
    out=[]
    for it in items:
        if '. ' in it and len(it.split('. ')[0])<40:
            a,b=it.split('. ',1); out.append(P(R(a+'. ',b=True),R(b),bullet=True,space_after=5))
        else: out.append(P(R(it),bullet=True,space_after=5))
    return out
def numbered(items):
    out=[]
    for i,it in enumerate(items,1):
        if '. ' in it and len(it.split('. ')[0])<30:
            a,b=it.split('. ',1); out.append(P(R(f'{i}. ',b=True),R(a+'. ',b=True),R(b),space_after=5))
        else: out.append(P(R(f'{i}. ',b=True),R(it),space_after=5))
    return out
def stat(sl,x,y,w,big,label):
    rect(sl,x,y,w,1.05,fill='FFFFFF',shadow_on=False)
    text(sl,x,y+.08,w,.55,[P(R(big,b=True,size=24,color=NAVY),align='c')])
    text(sl,x+.1,y+.62,w-.2,.38,[P(R(label,size=9.5,color=MUTED),align='c')])
def cover_pic(sl,path,x,y,w,h):
    im=Image.open(path);iw,ih=im.size;ar=iw/ih;br=w/h
    p=sl.shapes.add_picture(path,Inches(x),Inches(y),Inches(w),Inches(h))
    if ar>br: c=(1-br/ar)/2; p.crop_left=c; p.crop_right=c
    else: c=(1-ar/br)/2; p.crop_top=c; p.crop_bottom=c
    p.line.color.rgb=BLACK; p.line.width=Pt(1)
    return p
def pic(sl,n,x,y,w,h): cover_pic(sl,FR+'c%02d.jpg'%n,x,y,w,h)
def vid(sl,n,x,y,w,h,label=None):
    mp=VD+'s%02d.mp4'%n
    if os.path.exists(mp):
        mv=sl.shapes.add_movie(mp,Inches(x),Inches(y),Inches(w),Inches(h),poster_frame_image=SC+'vid/s%02d/cover.png'%n,mime_type='video/mp4')
        text(sl,x,y+h+.03,w,.22,[P(R('Click the image to play the animation',i=True,size=8,color=MUTED))])
    else:
        cover_pic(sl,FR+'f%02d.jpg'%n,x,y,w,h)
    if label: text(sl,x,y-.3,w,.26,[P(R(label,b=True,size=10))])
def notes(sl,t): sl.notes_slide.notes_text_frame.text=t
def new(): return prs.slides.add_slide(LAYOUT)
def num(sl,n): text(sl,12.5,6.85,.6,.25,[P(R(str(n),size=9,color=MUTED),align='r')])
N=0
def slide():
    global N; N+=1; sl=new(); return sl

# ===== 1 Title
sl=slide()
p=sl.shapes.add_picture(AS+'tgfd.png',Inches(4.4),Inches(1.25),Inches(4.53),Inches(4.8))
# fade the crest
blip=p._element.blipFill.blip; am=etree.SubElement(blip,qn('a:alphaModFix')); am.set('amt','28000')
text(sl,1,2.3,11.33,1.1,[P(R('J-FARP SEIZURE',b=True,size=48),align='c')])
text(sl,1,3.4,11.33,.5,[P(R('Concept brief · Ground Task Force mission set',b=True,size=20),align='c')])
text(sl,1,3.95,11.33,.5,[P(R('Dog Co · 2-506 IN · OCT 26',b=True,size=20),align='c')])
notes(sl,cap(1))
# ===== 2 BLUF
sl=slide();title(sl,'BLUF')
box(sl,.5,1.25,12.33,1.15,None,[P(R('Bottom line: ',b=True,u=True),R('A rifle company (-) can seize and secure a jump FARP 180–200 NM forward and hold it for one turn of the attack element, putting Division deep targets 100–150 NM beyond it inside attack aviation combat radius.'))],size=15,anchor='m')
stat(sl,.5,2.6,2.9,'3 hours','on the ground, wheels down to extraction');stat(sl,3.65,2.6,2.9,'82 PAX','Lift 1 · 2x CH-47 · 2x UH-60');stat(sl,6.8,2.6,2.9,'H + 2:00','FARP operations complete');stat(sl,9.95,2.6,2.88,'28–29 OCT 26','proof of concept · Operation Storm Surge')
box(sl,.5,3.9,7.9,2.95,'Decisions this brief supports',numbered(['Endorse. J-FARP seizure as a Ground Task Force mission set for the deep attack.','Resource. The full-distance proof of concept, Operation Storm Surge, 28–29 OCT 26.','Transition. On validation, hand the mission set to the Division pathfinder company.']),size=15)
box(sl,8.6,3.9,4.23,2.95,'Agenda',numbered(['Mission and problem','Concept of the operation','Task organization and phasing','Execution by phase','Risk and proof of concept']),size=13)
num(sl,N);notes(sl,'Bottom line up front, the numbers that matter, and the decisions this brief supports.')
# ===== 3 Mission
sl=slide();title(sl,'GTF Mission Statement')
box(sl,.8,1.55,11.73,1.55,None,[P(R('On order, an Infantry Company (-) conducts Air Assault Operations to seize a J-FARP to enable Combat Aviation Assets to mass combat power in the Division Deep Area to defeat enemy forces and seize key terrain.'),align='c')],size=20,anchor='m')
box(sl,.8,3.4,5.7,3.4,'Purpose',[para('Place Class III and Class V inside the Division deep area for one turn of the attack element, so the deep attack launches with full fuel and ordnance against targets that are otherwise beyond combat radius.')],size=14)
box(sl,6.83,3.4,5.7,3.4,'Key tasks',bullets(['Seize the FARP site and the objectives covering its approaches.','Establish the rapid refuel and rearm point from the WETCOW and FATCOW.','Turn the attack element forward with full fuel and ordnance.','Defend the site for the duration of FARP operations.','Close the site and extract through the PZ.']),size=14)
num(sl,N);notes(sl,cap(5))
# ===== 4 Problem
sl=slide();title(sl,'The Problem: Operational Reach')
vid(sl,2,.5,1.5,7.1,3.99,'Attack aviation from the ISB reaches bingo short of the Division deep target')
box(sl,7.85,1.25,4.98,2.35,'The problem',[para('Attack aviation combat radius is bounded by fuel. Ordnance, reserve and time on station draw on the same tank. From the intermediate staging base the Division deep target is out of reach.')],size=12.5)
stat(sl,7.85,3.75,2.4,'180–200 NM','ISB to J-FARP site');stat(sl,10.43,3.75,2.4,'100–150 NM','J-FARP to DIV deep target')
box(sl,7.85,4.95,4.98,1.85,'Without fuel forward',[para('The attack element turns for home short of the objective. Terrain beyond that line stays unshaped until Class III is positioned forward.')],size=12.5)
box(sl,.5,5.75,7.1,1.05,None,[lead('With a J-FARP:','refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius, with full station time on arrival.')],size=12.5,anchor='m')
num(sl,N);notes(sl,cap(2))
# ===== 5 Concept
sl=slide();title(sl,'Concept of the Operation')
vid(sl,5,.5,1.5,7.1,3.99,'Three lifts: Ground Task Force seizes, FARP Task Force establishes, attack element turns forward')
box(sl,7.85,1.25,4.98,1.75,'Lift 1 · Ground Task Force',[P(R('2x CH-47 · 2x UH-60 · 82 PAX',b=True)),para('GTF air assault. FARP seizure. Pathfinder operations.')],size=13)
box(sl,7.85,3.15,4.98,1.75,'Lift 2 · FARP Task Force',[P(R('2x CH-47 · WETCOW · FATCOW · 18 PAX',b=True)),para('FARP TF air assault. FARP establishment by ECHO / 6-101.')],size=13)
box(sl,7.85,5.05,4.98,1.75,'Lift 3 · Deep Attack',[P(R('2x AH-64',b=True)),para('Refuel and rearm at the J-FARP with rotors turning, continue to the DIV deep target (100–150 NM).')],size=13)
box(sl,.5,5.75,7.1,1.05,'End state',[para('FARP operations complete NLT H+2:00. Attack element over the deep target with full fuel. Site closed, force reconsolidated at the ISB by H+3:00.')],size=11.5)
num(sl,N);notes(sl,cap(5))
# ===== 6 Task org
sl=slide();title(sl,'Proposed Ground Task Force')
pic(sl,6,.5,1.25,5.6,5.55)
el=[('Element 1 · Interdict · 54 PAX','Occupies OBJ Husky, Poodle and Bernese, the objectives covering the avenues of approach to the site. Interdicts enemy forces approaching the FARP.'),('Element 2 · Seize · 12 PAX','Seizes the FARP site from LZ Vulture. Reports the site secured NLT H+1:00.'),('Element 3 · Secure · 9 PAX','Local security on the aircraft and refuel points for the duration of FARP operations. Pathfinders mark PZ Cedar and the refuel point.')]
for i,(a,b) in enumerate(el): box(sl,6.35,1.25+i*1.9,6.48,1.75,a,[para(b)],size=14)
num(sl,N);notes(sl,cap(6))
# ===== 7 Phasing
sl=slide();title(sl,'Phasing and H-Hour Sequence')
ln=sl.shapes.add_connector(1,Inches(1.2),Inches(1.95),Inches(12.1),Inches(1.95));ln.line.color.rgb=BLACK;ln.line.width=Pt(2)
ms=[('H − 2:30','Wheels up'),('H-hour','Wheels down'),('H + 1:00','FARP seized'),('H + 2:00','FARP ops complete'),('H + 3:00','Reconsolidated')]
for i,(a,b) in enumerate(ms):
    x=1.2+i*2.725; d=rect(sl,x-.12,1.83,.24,.24,fill='7030A0',shadow_on=False,shape=MSO_SHAPE.OVAL)
    text(sl,x-1.2,1.35,2.4,.4,[P(R(a,b=True,size=15),align='c')]); text(sl,x-1.2,2.15,2.4,.3,[P(R(b,size=10,color=MUTED),align='c')])
ph=[('Phase I · Condition setting','B: Planning start  E: Wheels up','Air assault planning (IPC, AMCM, AMB). FARP site selection. Cold loads, rehearsals, leader certification.'),('Phase II · Air assault and seizure','B: Wheels down (H-hour)  E: FARP seized (H+1:00)','Lift 1 lands three LZs. Objectives occupied. Site seized and PZ marked.'),('Phase III · FARP operations','B: H+1:00  E: FARP ops complete (H+2:00)','Lift 2 establishes the point. Lift 3 refuels, rearms and departs for the deep target.'),('Phase IV · FARP defense','Concurrent with Phase III','Ground force holds the perimeter for the window the aircraft are on the ground. Site closed and force extracted by H+3:00.')]
for i,(a,b,c) in enumerate(ph): box(sl,.5+i*3.1,2.7,2.98,4.1,a,[P(R(b,i=True),space_after=6),lead('Key tasks:',c)],size=13)
num(sl,N);notes(sl,cap(8))
# ===== 8 Phase II
sl=slide();title(sl,'Phase II: Air Assault and Seizure')
vid(sl,10,.5,1.5,7.1,3.99,'H-hour to H+1:00 · Lift 1 lands, objectives occupied, PZ Cedar marked')
box(sl,7.85,1.25,4.98,1.1,'Phase II',[para('B: Wheels down (H-hour)   E: FARP seized (H+1:00)'),lead('Key tasks:','GTF air assault, FARP seizure, pathfinder operations.')],size=11.5)
box(sl,7.85,2.5,4.98,4.3,'Scheme of maneuver',bullets(['H-hour. Three serials land LZ Vulture, Owl and Falcon on separate headings.','Seize. Element 2 seizes the FARP site from LZ Vulture.','Interdict. Element 1 occupies OBJ Husky, Poodle and Bernese, oriented on the avenues of approach.','Establish. Element 3 secures the site. Pathfinders mark PZ Cedar and the refuel point inside the triangle. Secured NLT H+1:00.']),size=13)
box(sl,.5,5.75,7.1,1.05,None,[P(R('Night air assault, dismounted land navigation, seize. ',b=True),R('Chalk-to-LZ assignments shown are the Storm Surge AMB plan; the CONOP Lift 1 package is 2x CH-47 and 2x UH-60.'))],size=11.5,anchor='m')
num(sl,N);notes(sl,cap(10))
# ===== 9 Phase III
sl=slide();title(sl,'Phase III: FARP Operations')
vid(sl,11,.5,1.5,7.1,3.99,'H+1:00 to H+2:00 · WETCOW and FATCOW land, Apaches refuel, rearm and turn forward')
box(sl,7.85,1.25,4.98,1.1,'Phase III',[para('B: FARP seized (H+1:00)   E: FARP ops complete (H+2:00)'),lead('Key tasks:','FARP TF air assault, FARP establishment, deep attack air assault.')],size=11.5)
box(sl,7.85,2.5,4.98,3.1,'Scheme of maneuver',bullets(['H+1:00. Lift 2 lands the 18-PAX FARP crew: WETCOW with fuel, FATCOW with AH-64 munitions.','Establish. Refuel points from the WETCOW, the rearm point from the FATCOW. ECHO / 6-101 operates the points.','Turn. Lift 3 lands, refuels and rearms with rotors turning, and departs for the deep target with full fuel.']),size=13)
stat(sl,7.85,5.75,2.4,'60 NM','AH reach added by the WETCOW');stat(sl,10.43,5.75,2.4,'185 NM','extended range, no refuel, ERFS')
box(sl,.5,5.75,7.1,1.05,None,[P(R('Pathfinders control the zone. ',b=True),R('Element 3 secures the aircraft and the points. FARP operations complete NLT H+2:00.'))],size=11.5,anchor='m')
num(sl,N);notes(sl,cap(11))
# ===== 10 Phase IV
sl=slide();title(sl,'Phase IV: FARP Defense')
vid(sl,12,.5,1.5,7.1,3.99,'Objectives engage along the avenues of approach while aircraft are on the ground')
box(sl,7.85,1.25,4.98,1.1,'Phase IV',[para('Concurrent with FARP operations'),lead('Key task:','FARP defense. ATP 3-04.17: the FARP is a high-value, high-signature target with minimal organic security. Security is a ground force task.')],size=11)
box(sl,7.85,2.5,4.98,4.3,'Security posture',bullets(['All-round security. A mobile enemy can approach from any direction. 360 degrees, no element split.','Early warning in depth. OPs on the avenues, then the objectives, then the centre.','Mutual support. Any approach under fire from two positions. The triangular patrol base delivers this with three elements.','Organic fires. Mortar section registered on the avenues.','Limited exposure. The FARP displaces or closes before the enemy can mass.']),size=12.5)
box(sl,.5,5.75,3.45,1.05,'PACE',[para('Wickr over TACSAT-IW · TAK / ICE · Signal over MUOS · cell or HF.')],size=10.5)
box(sl,4.15,5.75,3.45,1.05,'Displacement criteria',[para('Site compromised, or the alternate better positioned for the next turn.')],size=10.5)
num(sl,N);notes(sl,cap(12))
# ===== 11 Risk
sl=slide();title(sl,'Risk and Mitigation')
rk=[('Risk to force: distance from sustainment','A company (-) with limited sustainment 180 NM from FCKY. Casualty evacuation and resupply are aviation tasks that must be planned, not assumed.'),('Risk to mission: high-signature fuel point','Rotors turning, fuel flowing, aircraft on the ground on a schedule. The window is the vulnerability.'),('Training risk: newly qualified pathfinders','The company assumes risk training with newly qualified pathfinders; a Division pathfinder company is the intended recipient of the mission set.'),('Mitigation: Class I carried','1x DOS of CL I issued prior to W/U and 1x DOS carried per soldier on the aircraft.'),('Mitigation: deliberate rehearsals','Air assault planning with 6-101 CAB, cold-load training on the airframes and leader certification before execution.'),('Mitigation: proof of concept','Validate the task force on a training objective at full distance before the concept is relied on for a deep attack.')]
for i,(a,b) in enumerate(rk): box(sl,.5+(i%3)*4.15,1.3+(i//3)*2.8,4.03,2.6,a,[para(b)],size=14,fill='F8E0E0' if i<3 else 'E2EFDA')
num(sl,N);notes(sl,cap(14))
# ===== 12 Proof of concept
sl=slide();title(sl,'Proof of Concept: Operation Storm Surge')
pic(sl,15,.5,1.25,6.4,3.6)
box(sl,7.15,1.25,5.68,1.6,'Mission',[para('Dog Company conducts a Company (-) Air Assault into Tullahoma, TN to secure an expeditionary FARP on 29 0130Z OCT 26 to enable freedom of movement for 6-101 CAB and proof the concept for FARP seizure operations.')],size=11.5)
box(sl,7.15,3.0,5.68,2.5,'Deliverables',bullets(['Extended range: 185 NM with no refuel, ERFS employment on Serial 1.','WETCOW enables AH operations 60 NM beyond the objective.','3x WETCOWs, 5x FARP locations, simultaneous, secured and dispersed.','Air-deliverable Shadow TAC.']),size=12.5)
tl=[('2200Z','Wheels up'),('0030Z','W/D · H-hour'),('0130Z','FARP seized'),('0230Z','FARP ops complete'),('0330Z','Reconsolidated')]
for i,(a,b) in enumerate(tl): box(sl,.5+i*1.3,5.05,1.2,.95,None,[P(R(a,b=True,size=13),align='c'),P(R(b,size=8.5),align='c')],size=9,anchor='m')
box(sl,.5,6.15,6.4,.65,None,[lead('A/C package:','3x CH-47 (1x PAX, 1x FATCOW) · 1x UH-60 · 83 PAX · night · CAAF to Tullahoma, TN · 184 NM')],size=10.5,anchor='m')
box(sl,7.15,5.65,5.68,1.15,None,[lead('Training objectives:','07-CO-5135 Operate a company command post · 011-60Z-1054 Select a landing zone / pickup zone.')],size=10.5,anchor='m')
num(sl,N);notes(sl,cap(15))
# ===== 13 End state
sl=slide();title(sl,'End State and Decisions')
vid(sl,4,.5,1.5,7.1,3.99,'The attack element on the Division deep target with full fuel and ordnance')
box(sl,7.85,1.25,4.98,2.1,'End state',[P(R('Fuel forward is reach forward. ',b=True),R('An infantry company (-), three hours on the ground, extends the Division deep fight by 200 nautical miles for one turn of the attack element. Site closed, force reconsolidated with 100 percent accountability.'))],size=13.5)
box(sl,7.85,3.5,4.98,3.3,'Decisions requested',numbered(['Endorse J-FARP seizure as a Ground Task Force mission set.','Resource Operation Storm Surge, 28–29 OCT 26.','On validation, transition the mission set to the Division pathfinder company.']),size=14)
box(sl,.5,5.75,7.1,1.05,None,[P(R('Seize the site. Establish the FARP. Turn the deep attack. Close the site.',b=True),align='c')],size=14,anchor='m')
num(sl,N);notes(sl,cap(16))
# ===== Backup divider
sl=slide()
text(sl,1,2.7,11.33,1.0,[P(R('BACKUP',b=True,size=44),align='c')])
text(sl,1,3.75,11.33,.5,[P(R('Purpose · FARP doctrine · Reverse planning · Site selection · Reconsolidation · Terms and sources',size=14,color=MUTED),align='c')])
notes(sl,'Backup slides.')
def backup_tag(sl):
    s=rect(sl,10.5,.33,1.5,.34,fill='FFFFFF',line='C00000',shadow_on=False);tf=s.text_frame;tf.vertical_anchor=MSO_ANCHOR.MIDDLE;p=tf.paragraphs[0];p.alignment=PP_ALIGN.CENTER;r=p.add_run();r.text='BACKUP';r.font.name='Arial';r.font.size=Pt(10);r.font.bold=True;r.font.color.rgb=RED
# B1 Purpose
sl=slide();title(sl,'Purpose: The Division Deep Area');backup_tag(sl)
pic(sl,4,.5,1.25,5.6,5.55)
box(sl,6.35,1.25,6.48,1.6,None,[para('The Division deep area is the terrain beyond the close fight, where aviation shapes enemy forces before the main effort arrives. Reaching it is a sustainment problem before it is a tactical one.')],size=12.5)
box(sl,6.35,3.0,6.48,2.5,'What a forward FARP buys',bullets(['Reach. Refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius.','Station time. Aircraft arrive on the objective with a full fuel load rather than a reserve.','Tempo. Aircraft turn forward at the FARP in minutes instead of returning hours to the ISB.','Mass. More sorties on the objective within the same window.']),size=12)
box(sl,6.35,5.65,6.48,1.15,None,[P(R('For the 101st, the J-FARP is the enabler for the large-scale, long-range air assault. The deep attack sets conditions on the follow-on objective. The J-FARP sets conditions for the deep attack.',i=True))],size=11.5,anchor='m')
num(sl,N);notes(sl,cap(4))
# B2 Doctrine
sl=slide();title(sl,'FARP Variants in Doctrine');backup_tag(sl)
box(sl,.5,1.25,12.33,.95,None,[lead('FM 3-04 and ATP 3-04.17:','a temporary facility providing fuel and ammunition to aviation units in combat. Positioning Class III and Class V forward reduces turnaround distance and extends time in the fight. Four variants trade permanence for responsiveness.')],size=12.5,anchor='m')
v=[('Active · manned and operating','Equipment emplaced, crews in position, refueling in progress. The standard FARP supporting sustained operations.'),('Silent · established, not operating','Emplaced and concealed. Activated only when aircraft are inbound, to limit signature until required.'),('Jump · minimal footprint, air-delivered','Rapidly displaceable, with Class III and Class V delivered on the supporting aircraft. Established for a single deep operation, then displaced or closed.'),('Rolling · displaces with the force','Moves by bounds behind the maneuvering unit to keep aviation support within reach of the advance.')]
for i,(a,b) in enumerate(v): box(sl,.5+(i%2)*6.25,2.4+(i//2)*2.25,6.08,2.05,a,[para(b)],size=12.5,fill='FFF2CC' if i==2 else 'EDEDED')
num(sl,N);notes(sl,cap(3))
# B3 Reverse planning
sl=slide();title(sl,'Reverse Planning Sequence');backup_tag(sl)
box(sl,.5,1.25,12.33,.95,None,[para('Air assault doctrine develops five plans in reverse sequence, each constraining the next. For a FARP seizure the ground tactical plan carries a requirement most air assaults do not: the site must accept a CH-47 with internal fuel aboard.')],size=12.5,anchor='m')
steps=[('5 · Staging plan','Assembly, sequencing and PZ posture.'),('4 · Loading plan','Chalk and load assignments by aircraft.'),('3 · Air movement plan','Routes, timing and formation.'),('2 · Landing plan','Zone, heading and landing sequence.'),('1 · Ground tactical plan','Actions on the objective.')]
for i,(a,b) in enumerate(steps):
    x=.5+i*2.5; box(sl,x,2.45,2.3,1.6,a,[para(b)],size=11.5,fill='FFF2CC' if i==4 else 'EDEDED')
    if i<4:
        c=sl.shapes.add_connector(1,Inches(x+2.32),Inches(3.25),Inches(x+2.48),Inches(3.25));c.line.color.rgb=BLACK;c.line.width=Pt(1.5)
pc=[('Planning events · IPC, AMCM, AMB','Initial planning conference, air mission coordination meeting and air mission brief. One integrated plan between the aviation and ground task forces.'),('Pathfinders · LZ / PZ control','Select, mark and control landing and pickup zones, provide terminal guidance and conduct PZ operations for extraction.'),('Rehearsals · cold loads and leader certification','Cold-load training on the airframes, leader certification and deliberate rehearsals mitigate operating 180 NM from the ISB.')]
for i,(a,b) in enumerate(pc): box(sl,.5+i*4.15,4.3,4.03,2.5,a,[para(b)],size=12)
num(sl,N);notes(sl,cap(7))
# B4 Site selection
sl=slide();title(sl,'Phase I: FARP Site Selection');backup_tag(sl)
pic(sl,9,.5,1.25,5.6,5.55)
box(sl,6.35,1.25,6.48,.8,None,[lead('Key tasks:','air assault planning and FARP site selection. The site must satisfy three requirements concurrently.')],size=12,anchor='m')
box(sl,6.35,2.2,6.48,4.6,'Site requirements',bullets(['Aircraft. Dimensions and surface for a CH-47 with internal fuel aboard. Approach and departure headings clear of obstacles. Slope and dust within limits for night landing.','Target. Close enough to the Division deep target to preserve station time. Far enough that the fuel point is not on the objective.','Defense. Limited avenues of approach, cover for the security elements and an alternate site for displacement. Storm Surge plans three WETCOWs across five FARP locations.']),size=12.5)
num(sl,N);notes(sl,cap(9))
# B5 Reconsolidation
sl=slide();title(sl,'Reconsolidation: H+2:00 to H+3:00');backup_tag(sl)
pic(sl,13,.5,1.25,5.6,5.55)
box(sl,6.35,1.25,6.48,.95,None,[para('Reconsolidation distinguishes a jump FARP from a base. The fuel point exists for one turn of the attack element and is then closed.')],size=12,anchor='m')
box(sl,6.35,2.35,6.48,4.45,'Sequence',bullets(['H+2:00. FARP operations complete. Refuel and rearm points recovered. WETCOW and FATCOW lift first.','PZ operations. Pathfinders control the pickup: chalks staged, aircraft inbound 090, staggered left, weapons tight.','Collapse in reverse order. Objectives fold in last, under mortar coverage until the mortar section itself lifts.','H+3:00. Reconsolidated at FCKY with 100 percent accountability. FARP seizure task force validated.']),size=12.5)
num(sl,N);notes(sl,cap(13))
# B6 Terms
sl=slide();title(sl,'Terms and Sources');backup_tag(sl)
terms=[('J-FARP','Jump forward arming and refueling point'),('ISB','Intermediate staging base'),('GTF','Ground task force'),('WETCOW','CH-47 carrying fuel tanks: the Class III point'),('FATCOW','CH-47 carrying munitions for the AH-64: the Class V point'),('ERFS','Extended range fuel system'),('R3P','Rapid refuel point'),('L2A2','Large-scale, long-range air assault'),('IPC · AMCM · AMB','Initial planning conference · air mission coordination meeting · air mission brief'),('W/U · W/D','Wheels up · wheels down'),('OP · CCP','Observation post · casualty collection point'),('NOMAD','Unmanned ground vehicle carrying the C2 package'),('MDS','Mission design series')]
box(sl,.5,1.25,6.0,5.55,'Terms',[P(R(a+'  ',b=True),R(b),space_after=4) for a,b in terms],size=11)
box(sl,6.75,1.25,6.08,5.55,'Sources',[lead('Doctrine.','FM 3-04 Army Aviation, ATP 3-04.17 Techniques for Forward Arming and Refueling Points, FM 3-99 Airborne and Air Assault Operations, FM 3-90 Tactics, ATP 3-21.8 Infantry Platoon and Squad, TC 3-21.76 Ranger Handbook. Doctrinal statements are paraphrased and should be verified against current editions.',space_after=8),lead('Unit sources.','Mission, task organization, lift packages, phasing, timeline and imagery are from the Dog Company J-FARP Seizure CONOP and the Operation Storm Surge AMCM / AMB of 06 OCT 26.',space_after=8),lead('Imagery.','Scene renders and animations are from the animated brief. The wide-area terrain is illustrative and not real geography; objective-area scenes are drawn over the unit\'s overhead imagery. GTF elements and strengths are as directed by the GTF commander; chalk-to-LZ and objective assignments follow the Storm Surge AMB.')],size=11)
num(sl,N);notes(sl,'Terms and sources.')
prs.save(OUT);print('saved',OUT,N,'slides')
