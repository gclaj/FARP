const pptxgen=require('pptxgenjs');const fs=require('fs');const path=require('path');
const {applyTheme}=require('/root/.claude/skills/synced/41bc01b3-6bb5-4167-b63d-2218bfcbb52e_d7e8d10f-d5e7-46e2-89f0-06b91c176297/pptx/scripts/apply_theme.js');
const FR='/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/frames/';
const S=JSON.parse(fs.readFileSync('/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/slides.json','utf8'));
const cap=n=>S[n-1].cap;
const THEME={name:'J-FARP Night',headFontFace:'Arial',bodyFontFace:'Arial',colors:{dk1:'070B16',lt1:'EEF0F8',dk2:'141A2B',lt2:'AAB2CF',accent1:'F4B94A',accent2:'79D4FF',accent3:'FF7384',accent4:'8BE3B0',accent5:'6B7498',accent6:'FFD98A',hlink:'79D4FF',folHlink:'AAB2CF'}};
const pres=new pptxgen();pres.layout='LAYOUT_WIDE';pres.theme={headFontFace:THEME.headFontFace,bodyFontFace:THEME.bodyFontFace};
pres.author='Dog Company, 2-506 IN';pres.title='J-FARP Seizure';pres.subject='Concept of the operation';
const C=pres.SchemeColor;const W=13.333,H=7.5;
const FOOT='Dog Company · 2-506 IN · 3BCT, 101st ABN DIV (AASLT) · with 6-101 CAB';
pres.defineSlideMaster({title:'CONTENT',background:{color:THEME.colors.dk1},objects:[
  {placeholder:{options:{name:'title',type:'title',x:.6,y:.38,w:12.1,h:.9,fontSize:30,bold:true,color:C.background1,align:'left',margin:0,valign:'top'},text:''}},
  {text:{text:FOOT,options:{x:.6,y:7.05,w:9,h:.3,fontSize:9,color:THEME.colors.accent5,margin:0}}}],
  slideNumber:{x:12.3,y:7.05,w:.5,h:.3,fontSize:9,color:THEME.colors.accent5,align:'right'}});
pres.defineSlideMaster({title:'DARK',background:{color:THEME.colors.dk1},objects:[]});
// helpers
const img=(sl,n,x,y,w,h,opt)=>{const o=Object.assign({},opt||{});const full=o.full;delete o.full;sl.addImage(Object.assign({path:FR+(full?'f':'c')+String(n).padStart(2,'0')+'.jpg',x,y,w,h,sizing:{type:'cover',w,h},objectName:'Scene '+n},o))};
const title=(sl,t)=>sl.addText(t,{placeholder:'title'});
const eyebrow=(sl,t,x,y,w)=>sl.addText(t.toUpperCase(),{x,y,w,h:.3,fontSize:10,color:THEME.colors.accent1,charSpacing:3,margin:0,isTextBox:true,objectName:'Eyebrow'});
const para=(sl,t,x,y,w,h,opt)=>sl.addText(t,Object.assign({x,y,w,h,fontSize:14,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Body'},opt||{}));
const bullets=(sl,items,x,y,w,h,opt)=>{const runs=[];items.forEach((it,i)=>{const m=/^([^.]+\.)\s(.*)$/.exec(it);if(m){runs.push({text:m[1]+' ',options:{bold:true,color:THEME.colors.accent6,bullet:true,paraSpaceAfter:8}});runs.push({text:m[2],options:{breakLine:i<items.length-1}})}else runs.push({text:it,options:{bullet:true,paraSpaceAfter:8,breakLine:i<items.length-1}})});
  sl.addText(runs,Object.assign({x,y,w,h,fontSize:14,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Bullets'},opt||{}))};
const card=(sl,x,y,w,h,tag,head,body,opt)=>{opt=opt||{};sl.addShape(pres.ShapeType.roundRect,{x,y,w,h,rectRadius:.08,fill:{color:THEME.colors.dk2},line:{color:'2A3350',width:.75},objectName:'Card'});
  let cy=y+.16;if(tag){sl.addText(tag.toUpperCase(),{x:x+.2,y:cy,w:w-.4,h:.26,fontSize:9,color:opt.tagColor||THEME.colors.accent2,charSpacing:2,margin:0,isTextBox:true,objectName:'Card tag'});cy+=.3}
  if(head){sl.addText(head,{x:x+.2,y:cy,w:w-.4-(opt.stat?1.2:0),h:.4,fontSize:opt.headSize||16,bold:true,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Card head'});cy+=.44}
  if(body)sl.addText(body,{x:x+.2,y:cy,w:w-.4-(opt.stat?1.2:0),h:y+h-cy-.14,fontSize:opt.bodySize||13,color:C.background2,margin:0,valign:'top',isTextBox:true,objectName:'Card body'});
  if(opt.stat)sl.addText([{text:opt.stat,options:{fontSize:34,bold:true,color:THEME.colors.accent1,breakLine:true}},{text:'PAX',options:{fontSize:10,color:C.background2,charSpacing:2}}],{x:x+w-1.3,y:y+.4,w:1.1,h:1.1,align:'right',margin:0,valign:'top',isTextBox:true,objectName:'Card stat'})};
const stat=(sl,x,y,w,big,label)=>{sl.addText(big,{x,y,w,h:.8,fontSize:28,bold:true,color:THEME.colors.accent1,margin:0,isTextBox:true,objectName:'Stat'});sl.addText(label.toUpperCase(),{x,y:y+.8,w,h:.3,fontSize:9,color:C.background2,charSpacing:2,margin:0,isTextBox:true,objectName:'Stat label'})};
let sl;
// ---------- 1 Title
pres.addSection({title:'Concept'});
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Concept'});img(sl,1,0,0,W,H,{full:true});
sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'070B16',transparency:35},line:{color:'070B16',transparency:100},objectName:'Scrim'});
sl.addShape(pres.ShapeType.roundRect,{x:.7,y:.7,w:7.4,h:3.9,rectRadius:.1,fill:{color:'0E1322',transparency:18},line:{color:'2A3350',width:.75},objectName:'Panel'});
eyebrow(sl,'Concept of the operation · J-FARP seizure',1.1,1.05,6.6);
sl.addText([{text:'J-FARP ',options:{color:C.background1}},{text:'Seizure.',options:{color:THEME.colors.accent1,italic:true}}],{x:1.1,y:1.4,w:6.6,h:1.1,fontSize:54,bold:true,margin:0,isTextBox:true,objectName:'Title'});
para(sl,S[0].lines[2].replace('[p] ',''),1.1,2.65,6.6,1.7,{fontSize:16});
sl.addText(FOOT,{x:.7,y:6.9,w:9,h:.3,fontSize:10,color:C.background2,margin:0,isTextBox:true,objectName:'Footer'});
sl.addNotes(cap(1));
// ---------- 2 Reach
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'Deep targets lie beyond combat radius');
eyebrow(sl,'The problem · operational reach',.6,1.3,6);
img(sl,2,6.9,1.7,5.85,5.0);
para(sl,'Attack aviation combat radius is bounded by fuel load. Ordnance, required reserve and time on station all draw on the same tank. When the Division deep target sits beyond that radius, the deep attack cannot be executed from the intermediate staging base.',.6,1.7,5.9,1.9);
stat(sl,.6,3.75,2.9,'180–200 NM','ISB to J-FARP site');stat(sl,3.6,3.75,2.9,'100–150 NM','J-FARP site to DIV deep TGT');
para(sl,[{text:'Bingo fuel. ',options:{bold:true,color:THEME.colors.accent6}},{text:'The attack element reaches bingo short of the objective and must return to base. Terrain beyond that line remains unshaped until fuel is positioned forward.'}],.6,5.1,5.9,1.4);
sl.addNotes(cap(2));
// ---------- 3 Doctrine
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'A temporary facility providing fuel and ammunition to aviation units in combat');
eyebrow(sl,'Doctrine · forward arming and refueling point',.6,1.45,8);
img(sl,3,.6,1.9,4.6,4.9);
para(sl,'FM 3-04 and ATP 3-04.17 define the FARP by that purpose. Positioning Class III and Class V forward reduces turnaround distance and extends the time an aircraft can remain in the fight. Four variants are described, each trading permanence for responsiveness.',5.5,1.9,7.2,1.1);
const v=[['Active','Manned and operating','Equipment emplaced, crews in position, refueling in progress. The standard FARP supporting sustained operations.'],['Silent','Established, not operating','Emplaced and concealed. Activated only when aircraft are inbound, to limit signature until required.'],['Jump','Minimal footprint, air-delivered','Rapidly displaceable, with Class III and Class V delivered on the supporting aircraft. Established to support a single deep operation, then displaced or closed.'],['Rolling','Displaces with the force','Moves by bounds behind the maneuvering unit to keep aviation support within reach of the advance.']];
v.forEach((c,i)=>card(sl,5.5+(i%2)*3.65,3.15+Math.floor(i/2)*1.85,3.55,1.7,c[0],c[1],c[2],{tagColor:i===2?THEME.colors.accent1:THEME.colors.accent2,headSize:14,bodySize:11.5}));
sl.addNotes(cap(3));
// ---------- 4 Purpose
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'A FARP extends operational reach');
eyebrow(sl,'Purpose · the Division deep area',.6,1.3,6);
img(sl,4,6.9,1.7,5.85,5.0);
para(sl,'The Division deep area is the terrain beyond the close fight, where aviation shapes enemy forces before the main effort arrives. Reaching it is a sustainment problem before it is a tactical one.',.6,1.7,5.9,1.1);
bullets(sl,['Reach. Refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius.','Station time. Aircraft arrive on the objective with a full fuel load rather than a reserve.','Tempo. Aircraft turn forward at the FARP in minutes instead of returning hours to the ISB.','Mass. More sorties on the objective within the same window. The mission statement is explicit: mass combat power in the Division deep area.'],.6,2.9,5.9,2.6);
para(sl,'For the 101st, the J-FARP is the enabler for the large-scale, long-range air assault. The deep attack sets conditions on the follow-on objective. The J-FARP sets conditions for the deep attack.',.6,5.6,5.9,1.2,{italic:true,color:C.background2});
sl.addNotes(cap(4));
// ---------- 5 CONOP
pres.addSection({title:'Scheme of maneuver'});
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Scheme of maneuver'});img(sl,5,0,0,W,H,{full:true});
sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'070B16',transparency:45},line:{color:'070B16',transparency:100},objectName:'Scrim'});
sl.addShape(pres.ShapeType.roundRect,{x:.6,y:.5,w:7.6,h:2.55,rectRadius:.1,fill:{color:'0E1322',transparency:15},line:{color:'2A3350',width:.75},objectName:'Panel'});
eyebrow(sl,'Concept of the operation · three lifts',.95,.75,6.9);
sl.addText('Seize the site. Establish the FARP. Turn the deep attack.',{x:.95,y:1.05,w:6.9,h:.8,fontSize:24,bold:true,color:C.background1,margin:0,isTextBox:true,objectName:'Title'});
para(sl,[{text:'Mission. ',options:{bold:true,color:THEME.colors.accent6}},{text:'On order, an Infantry Company (-) conducts Air Assault Operations to seize a J-FARP to enable Combat Aviation Assets to mass combat power in the Division Deep Area to defeat enemy forces and seize key terrain.'}],.95,1.9,6.9,1.1,{fontSize:12.5});
const lifts=[['Lift 1 – Ground Task Force','2x CH47 · 2x UH60 · C2 (NOMAD) · 82 PAX','GTF air assault. FARP seizure. Pathfinder operations.',THEME.colors.accent2],['Lift 2 – FARP Task Force','2x CH47 · FATCOW · WETCOW · 18 PAX (ECHO / 6-101 FARP crew)','FARP TF air assault. FARP establishment.',THEME.colors.accent1],['Lift 3 – Deep Attack','2x AH64','Deep attack air assault. Refuel and rearm at the J-FARP, continue to the DIV deep TGT (100–150 NM).',THEME.colors.accent3]];
lifts.forEach((l,i)=>card(sl,.6+i*4.1,5.25,3.95,1.75,l[0],l[1],l[2],{tagColor:l[3],headSize:13,bodySize:12}));
sl.addText(FOOT,{x:.6,y:7.1,w:9,h:.3,fontSize:9,color:THEME.colors.accent5,margin:0,isTextBox:true,objectName:'Footer'});
sl.addNotes(cap(5));
// ---------- 6 GTF
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Scheme of maneuver'});title(sl,'Proposed Ground Task Force');
eyebrow(sl,'Task organization · Lift 1 · 2x CH47 · 2x UH60',.6,1.3,8);
img(sl,6,.6,1.8,5.3,5.0);
const el=[['Element 1','Interdict','Interdict enemy forces along the avenues of approach leading to the FARP site. Occupies the objectives that form the perimeter around the site.','54'],['Element 2','Seize','Seize the FARP site. Clears the site and reports it secured NLT H+1:00.','12'],['Element 3','Secure','Secure the FARP site. Local security on the aircraft and refuel points for the duration of FARP operations.','9']];
el.forEach((e,i)=>card(sl,6.2,1.8+i*1.7,6.55,1.55,e[0],e[1],e[2],{stat:e[3],headSize:16,bodySize:12.5}));
sl.addNotes(cap(6));
// ---------- 7 Planning
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Scheme of maneuver'});title(sl,'Reverse planning from the ground tactical plan');
eyebrow(sl,'Doctrine · air assault planning',.6,1.3,6);
para(sl,'Airborne and air assault doctrine develops five plans in reverse sequence. Each constrains the next. For a FARP seizure, the ground tactical plan is governed by a requirement most air assaults do not carry: the site must accept a CH-47 with internal fuel aboard.',.6,1.7,12.1,.9,{fontSize:13});
const steps=[['5','Staging plan','Assembly, sequencing and PZ posture.'],['4','Loading plan','Chalk and load assignments by aircraft.'],['3','Air movement plan','Routes, timing and formation.'],['2','Landing plan','Zone, heading and landing sequence.'],['1','Ground tactical plan','Actions on the objective.']];
steps.forEach((s,i)=>{const x=.6+i*2.47;sl.addShape(pres.ShapeType.ellipse,{x:x,y:2.75,w:.5,h:.5,fill:{color:i===4?THEME.colors.accent1:THEME.colors.dk2},line:{color:i===4?THEME.colors.accent1:'2A3350',width:.75},objectName:'Step'});
  sl.addText(s[0],{x:x,y:2.75,w:.5,h:.5,fontSize:14,bold:true,color:i===4?'070B16':C.background1,align:'center',valign:'middle',margin:0,isTextBox:true,objectName:'Step number'});
  if(i<4)sl.addShape(pres.ShapeType.line,{x:x+.55,y:3.0,w:1.8,h:0,line:{color:'2A3350',width:1,endArrowType:'triangle'},objectName:'Arrow'});
  sl.addText(s[1],{x:x,y:3.35,w:2.3,h:.4,fontSize:14,bold:true,color:C.background1,margin:0,isTextBox:true,objectName:'Step title'});
  sl.addText(s[2],{x:x,y:3.75,w:2.3,h:.6,fontSize:11.5,color:C.background2,margin:0,valign:'top',isTextBox:true,objectName:'Step body'})});
const pc=[['Planning events','IPC · AMCM · AMB','Initial planning conference, air mission coordination meeting and air mission brief. The aviation task force and the ground task force produce a single integrated plan.'],['Pathfinders','LZ / PZ control','Select, mark and control landing and pickup zones, provide terminal guidance to inbound aircraft and conduct PZ operations for extraction.'],['Rehearsals','Cold loads and leader certification','Cold-load training on the supporting airframes, leader certification and deliberate rehearsals mitigate the risk of operating 180 NM from the ISB.']];
pc.forEach((c,i)=>card(sl,.6+i*4.1,4.6,3.95,2.3,c[0],c[1],c[2],{headSize:14,bodySize:12}));
sl.addNotes(cap(7));
// ---------- 8 Phasing
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Scheme of maneuver'});title(sl,'Four phases, three hours on the ground');
eyebrow(sl,'FARP seizure concept · proposed phasing',.6,1.3,6);
const ms=[['H − 2:30','W/U'],['H-hour','W/D'],['H + 1:00','FARP seized'],['H + 2:00','FARP ops complete'],['H + 3:00','Reconsolidated']];
sl.addShape(pres.ShapeType.line,{x:1.2,y:2.45,w:10.9,h:0,line:{color:'2A3350',width:2},objectName:'Timeline'});
ms.forEach((m,i)=>{const x=1.2+i*2.725;sl.addShape(pres.ShapeType.ellipse,{x:x-.11,y:2.34,w:.22,h:.22,fill:{color:THEME.colors.accent1},line:{color:THEME.colors.accent1},objectName:'Milestone'});
  sl.addText(m[0],{x:x-1.2,y:1.85,w:2.4,h:.4,fontSize:16,bold:true,color:C.background1,align:'center',margin:0,isTextBox:true,objectName:'Milestone time'});
  sl.addText(m[1].toUpperCase(),{x:x-1.2,y:2.65,w:2.4,h:.3,fontSize:9,color:C.background2,align:'center',charSpacing:2,margin:0,isTextBox:true,objectName:'Milestone label'})});
const ph=[['Phase I','Condition setting','Air assault planning · FARP site selection'],['Phase II','Air assault & seizure','GTF air assault (Lift 1) · FARP seizure · Pathfinder operations'],['Phase III','FARP operations','FARP TF air assault (Lift 2) · FARP establishment · Deep attack air assault (Lift 3)'],['Phase IV','FARP defense','FARP defense']];
ph.forEach((p,i)=>card(sl,.6+i*3.08,3.3,2.95,3.4,p[0],p[1],'Key tasks: '+p[2],{headSize:16,bodySize:13,tagColor:THEME.colors.accent1}));
sl.addNotes(cap(8));
// ---------- 9 Phase I
pres.addSection({title:'Execution'});
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Select the site. Plan in reverse from it');
eyebrow(sl,'Phase I · condition setting',.6,1.3,6);
img(sl,9,6.9,1.7,5.85,5.0);
para(sl,'Key tasks: air assault planning and FARP site selection. The site must satisfy three requirements concurrently.',.6,1.7,5.9,.8);
bullets(sl,['Aircraft. Dimensions and surface for a CH-47 with internal fuel aboard. Approach and departure headings clear of obstacles. Slope and dust within limits for night landing.','Target. Close enough to the Division deep target to preserve station time. Far enough that the fuel point is not on the objective.','Defense. Limited avenues of approach, cover for the security elements and an alternate site for displacement. Storm Surge plans three WETCOWs across five FARP locations.'],.6,2.6,5.9,3.4);
sl.addNotes(cap(9));
// ---------- 10 Phase II
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Land. Seize the objectives. Establish the PZ');
eyebrow(sl,'Phase II · air assault and seizure',.6,1.3,6);
img(sl,10,.6,1.7,6.2,5.2);
para(sl,'Key tasks: GTF air assault (Lift 1), FARP seizure, pathfinder operations. Night air assault, dismounted land navigation, seize.',7.1,1.7,5.65,.8,{fontSize:13});
bullets(sl,['H-hour. Wheels down. CH-47 on LZ Vulture inbound 158°. CH-47 on LZ Owl inbound 063°. UH-60 on LZ Falcon inbound 155°.','Seize the site. Element 2 moves from LZ Vulture and seizes the FARP site. Element 1 moves to OBJ Husky, oriented north.','Interdict the approaches. Element 1 occupies OBJ Poodle from LZ Falcon, oriented east along the road, and OBJ Bernese from LZ Owl, oriented south-east, interdicting the avenues of approach.','Establish the PZ. Element 3 secures the site. Pathfinders mark PZ Cedar and the R3P inside the triangle formed by the objectives. FARP secured NLT H+1:00.'],7.1,2.5,5.65,3.9,{fontSize:12.5});
para(sl,'Chalk-to-LZ assignments shown are the Storm Surge AMB plan (3 aircraft). The CONOP Lift 1 package is 2x CH47 / 2x UH60.',7.1,6.4,5.65,.55,{fontSize:10,color:C.background2});
sl.addNotes(cap(10));
// ---------- 11 Phase III
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Establish the R3P. Turn the deep attack');
eyebrow(sl,'Phase III · FARP operations',.6,1.3,6);
img(sl,11,.6,1.7,6.2,5.2);
para(sl,'Key tasks: FARP TF air assault (Lift 2), FARP establishment, deep attack air assault (Lift 3). Pathfinders control the zone.',7.1,1.7,5.65,.8,{fontSize:13});
bullets(sl,['H+1:00. Lift 2 lands. 2x CH47 with the 18-PAX FARP crew land the marked zone: WETCOW with fuel tanks, FATCOW with AH-64 munitions.','FARP establishment. Refuel points run from the WETCOW; the rearm point is served from the FATCOW. ECHO / 6-101 operates the points. Element 3 secures the zone.','Lift 3 lands. AH-64s land the points, refuel from the WETCOW and rearm from the FATCOW with rotors turning.','Turn forward. The attack element departs for the Division deep target with full fuel. FARP operations complete NLT H+2:00.'],7.1,2.5,5.65,3.2,{fontSize:12.5});
stat(sl,7.1,5.75,2.7,'60 NM','AH operations extended by WETCOW');stat(sl,10.0,5.75,2.75,'185 NM','Extended range, no refuel, ERFS');
sl.addNotes(cap(11));
// ---------- 12 Phase IV
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Security is a ground force task');
eyebrow(sl,'Phase IV · security posture',.6,1.3,6);
img(sl,12,.6,1.7,5.4,3.4);
para(sl,'Key task: FARP Defense. ATP 3-04.17 treats the FARP as a high-value, high-signature target. The aviation unit brings minimal organic security; the ground task force provides it. The posture follows the principles of the defense applied at company scale.',.6,5.25,5.4,1.7,{fontSize:12});
bullets(sl,['All-round security. The FARP is approached from any direction by a mobile enemy. The perimeter must cover 360 degrees with no element split.','Early warning in depth. Scout observation posts forward on the avenues of approach, then the objectives, then the centre. Contact is made as far from the aircraft as possible.','Mutual support. Positions are sited so that any approach is under fire from two positions. The triangular patrol base is the doctrinal shape that delivers this with three elements.','Organic fires. The mortar section, registered on the avenues, is the company\'s responsive indirect fire. Aircraft on the points are the protected asset, not part of the perimeter.','Limited exposure. Time on site is the vulnerability. A jump FARP displaces or closes before the enemy can mass against it.'],6.3,1.7,6.45,4.0,{fontSize:12});
card(sl,6.3,5.75,3.15,1.2,'PACE',null,'Primary Wickr over TACSAT-IW. Alternate TAK / ICE. Contingency Signal over MUOS. Emergency cell or HF.',{bodySize:11});
card(sl,9.6,5.75,3.15,1.2,'Displacement criteria',null,'Site compromised, or the alternate site better positioned for the next turn. The FARP moves before it is forced to.',{bodySize:11});
sl.addNotes(cap(12));
// ---------- 13 Reconsolidation
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Close the site. The deep attack is on the objective');
eyebrow(sl,'Reconsolidation · H+2:00 to H+3:00',.6,1.3,6);
img(sl,13,6.9,1.7,5.85,5.0);
para(sl,'Storm Surge Phase 5, reconsolidation, distinguishes a jump FARP from a base. The fuel point exists for one turn of the attack element and is then closed.',.6,1.7,5.9,.9);
bullets(sl,['H+2:00. FARP operations complete. Refuel and rearm points recovered. WETCOW and FATCOW lift first.','PZ operations. Pathfinders control the pickup: chalks staged, aircraft inbound 090, staggered left, weapons tight.','Collapse in reverse order. Objectives fold in last, in reverse order of occupation, under mortar coverage until the mortar section itself lifts.','H+3:00. Reconsolidated at FCKY. Friendly end state: 100 percent accountability of personnel and equipment, FARP seizure task force validated.'],.6,2.7,5.9,3.6);
sl.addNotes(cap(13));
// ---------- 14 Risk
pres.addSection({title:'Risk and validation'});
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Risk and validation'});title(sl,'Risk');
const rk=[['Risk to force','Distance from sustainment','A company (-) with limited sustainment 180 nautical miles from FCKY. Casualty evacuation and resupply are aviation tasks that must be planned, not assumed.',THEME.colors.accent3],['Risk to mission','High-signature fuel point','Rotors turning, fuel flowing, aircraft on the ground on a schedule. The window is the vulnerability. The concept mitigates with objective security and a compressed timeline.',THEME.colors.accent3],['Training risk','Newly qualified pathfinders','The company assumes risk by training with newly qualified pathfinders, because a Division pathfinder company is the intended recipient of this mission set.',THEME.colors.accent3],['Mitigation','Class I carried','1× DOS of CL I issued prior to W/U and 1× DOS carried per soldier on the aircraft.',THEME.colors.accent1],['Mitigation','Deliberate rehearsals','Air assault planning with 6-101 CAB, cold-load training on the supporting airframes, and leader certification prior to execution.',THEME.colors.accent1],['Mitigation','Proof of concept','Validate the task force on a training objective at full distance before the concept is relied upon for a deep attack.',THEME.colors.accent1]];
rk.forEach((r,i)=>card(sl,.6+(i%3)*4.1,1.5+Math.floor(i/3)*2.65,3.95,2.45,r[0],r[1],r[2],{tagColor:r[3],headSize:16,bodySize:12.5}));
sl.addNotes(cap(14));
// ---------- 15 Storm Surge
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Risk and validation'});title(sl,'Validate at full distance');
eyebrow(sl,'Proof of concept · Operation Storm Surge · 28–29 OCT 26',.6,1.3,8);
img(sl,15,.6,1.7,6.2,4.1);
para(sl,[{text:'Mission. ',options:{bold:true,color:THEME.colors.accent6}},{text:'Dog Company conducts a Company (-) Air Assault into Tullahoma, TN to secure an expeditionary FARP on 29 0130Z OCT 26 to enable freedom of movement for 6-101 CAB and proof the concept for FARP seizure operations.'}],7.1,1.7,5.65,1.3,{fontSize:12.5});
bullets(sl,['Extended range: 185 NM with no refuel, ERFS employment on Serial 1.','WETCOW enables AH operations 60 NM beyond the objective.','3× WETCOWs, 5× FARP locations, simultaneous, secured and dispersed, supporting multiple MDS.','Hippo concept validated at 370 NM total, with integration of sustainment and C2 nodes including an air-deliverable Shadow TAC.'],7.1,3.05,5.65,2.7,{fontSize:12.5});
const tl=[['2200Z','W/U'],['0030Z','W/D · H-hour'],['0130Z','FARP seized'],['0230Z','FARP ops complete'],['0330Z','Reconsolidated']];
tl.forEach((m,i)=>{const x=.6+i*1.25;sl.addText(m[0],{x,y:5.95,w:1.22,h:.35,fontSize:14,bold:true,color:THEME.colors.accent1,margin:0,isTextBox:true,objectName:'Time'});sl.addText(m[1].toUpperCase(),{x,y:6.3,w:1.22,h:.5,fontSize:7.5,color:C.background2,margin:0,valign:'top',isTextBox:true,objectName:'Time label'})});
para(sl,'A/C package: 3× CH-47 (1× PAX, 1× FATCOW) · 1× UH-60 · 83 PAX · night · CAAF to Tullahoma, TN · 184 NM',7.1,5.95,5.65,.8,{fontSize:11,color:C.background2});
sl.addNotes(cap(15));
// ---------- 16 End state
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Risk and validation'});img(sl,16,0,0,W,H,{full:true});
sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'070B16',transparency:40},line:{color:'070B16',transparency:100},objectName:'Scrim'});
sl.addShape(pres.ShapeType.roundRect,{x:.7,y:.7,w:7.6,h:3.75,rectRadius:.1,fill:{color:'0E1322',transparency:18},line:{color:'2A3350',width:.75},objectName:'Panel'});
eyebrow(sl,'End state',1.1,1.0,6.8);
sl.addText([{text:'Fuel forward is ',options:{color:C.background1}},{text:'reach forward.',options:{color:THEME.colors.accent1,italic:true}}],{x:1.1,y:1.3,w:6.8,h:1.0,fontSize:40,bold:true,margin:0,isTextBox:true,objectName:'Title'});
para(sl,'Seize the site. Establish the FARP. Turn the deep attack. Close the site. An infantry company (-), three hours on the ground, extends the Division\'s deep fight by 200 nautical miles.',1.1,2.85,6.8,1.4,{fontSize:15});
sl.addText(FOOT,{x:.7,y:6.9,w:9,h:.3,fontSize:10,color:C.background2,margin:0,isTextBox:true,objectName:'Footer'});
sl.addNotes(cap(16));
// ---------- 17 Terms and sources
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Risk and validation'});title(sl,'Terms and sources');
const terms=[['J-FARP','Jump forward arming and refueling point'],['ISB','Intermediate staging base'],['GTF','Ground task force'],['WETCOW','CH-47 carrying fuel tanks: the Class III point'],['FATCOW','CH-47 carrying munitions for the AH-64: the Class V point'],['ERFS','Extended range fuel system'],['R3P','Rapid refuel point'],['L2A2','Large-scale, long-range air assault'],['IPC · AMCM · AMB','Initial planning conference · air mission coordination meeting · air mission brief'],['W/U · W/D','Wheels up · wheels down'],['OP · CCP','Observation post · casualty collection point'],['NOMAD','Unmanned ground vehicle carrying the C2 package'],['MDS','Mission design series']];
const truns=[];terms.forEach((t,i)=>{truns.push({text:t[0]+'  ',options:{bold:true,color:THEME.colors.accent6}});truns.push({text:t[1],options:{breakLine:i<terms.length-1,paraSpaceAfter:5}})});
sl.addText(truns,{x:.6,y:1.45,w:5.6,h:5.4,fontSize:11.5,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Terms'});
const src=[{text:'Doctrine. ',options:{bold:true,color:THEME.colors.accent6}},{text:'FM 3-04 Army Aviation, ATP 3-04.17 Techniques for Forward Arming and Refueling Points, FM 3-99 Airborne and Air Assault Operations, FM 3-90 Tactics, ATP 3-21.8 Infantry Platoon and Squad and TC 3-21.76 Ranger Handbook. Doctrinal statements are paraphrased and should be verified against current editions.',options:{breakLine:true,paraSpaceAfter:8}},
 {text:'Unit sources. ',options:{bold:true,color:THEME.colors.accent6}},{text:'Mission, intent, task organization, lift packages, phasing, timeline, deliverables, aerial imagery and apron photographs are from the Dog Company J-FARP Seizure CONOP and the Operation Storm Surge AMCM / AMB of 06 OCT 26.',options:{breakLine:true,paraSpaceAfter:8}},
 {text:'Imagery. ',options:{bold:true,color:THEME.colors.accent6}},{text:'Scene renders are from the animated brief. The wide-area terrain is built from illustrative terrain-board renders and is not real geography; the objective-area scenes are drawn over the unit\'s overhead imagery. Lift packages, phasing and key tasks are quoted from the CONOP as briefed; the Ground Task Force elements and strengths are as directed by the GTF commander; chalk-to-LZ and objective assignments follow the Storm Surge AMB.'}];
sl.addText(src,{x:6.6,y:1.45,w:6.15,h:5.4,fontSize:11.5,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Sources'});
sl.addNotes('Terms and sources for the brief.');
(async()=>{await pres.writeFile({fileName:'J-FARP_Seizure.pptx'});await applyTheme('J-FARP_Seizure.pptx',THEME);console.log('written')})();
