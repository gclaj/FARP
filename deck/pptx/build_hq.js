const pptxgen=require('pptxgenjs');const fs=require('fs');
const {applyTheme}=require('/root/.claude/skills/synced/41bc01b3-6bb5-4167-b63d-2218bfcbb52e_d7e8d10f-d5e7-46e2-89f0-06b91c176297/pptx/scripts/apply_theme.js');
const SC='/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/';const FR=SC+'frames/',VD=SC+(process.env.VIDDIR||'vid')+'/',AS=__dirname+'/assets/';const OUT=process.env.OUT||'J-FARP_Seizure_HQ.pptx';
const S=JSON.parse(fs.readFileSync(SC+'slides.json','utf8'));const cap=n=>S[n-1].cap;
const THEME={name:'Currahee',headFontFace:'Arial',bodyFontFace:'Arial',colors:{dk1:'000000',lt1:'FFFFFF',dk2:'1F1F1F',lt2:'EDEDED',accent1:'7030A0',accent2:'FFC000',accent3:'C00000',accent4:'1F4E79',accent5:'595959',accent6:'00B050',hlink:'1F4E79',folHlink:'595959'}};
const pres=new pptxgen();pres.layout='LAYOUT_WIDE';pres.theme={headFontFace:'Arial',bodyFontFace:'Arial'};
pres.author='Dog Company, 2-506 IN';pres.title='J-FARP Seizure - Concept brief';
const C=pres.SchemeColor,B=THEME.colors;const W=13.333,H=7.5;
const CLASS='/CONTROLLED UNCLASSIFIED INFORMATION/';
const chrome=[
 {image:{path:AS+'tgfd.png',x:4.35,y:1.3,w:4.63,h:4.9,transparency:90}},
 {image:{path:AS+'spade506.jpg',x:.25,y:.12,w:.72,h:.78}},
 {image:{path:AS+'rifles.png',x:12.35,y:.12,w:.75,h:.8}},
 {line:{x:0,y:1.0,w:W,h:0,line:{color:'000000',width:1.5}}},
 {rect:{x:0,y:7.12,w:W,h:.38,fill:{color:'000000'}}},
 {rect:{x:3.4,y:7.15,w:5.0,h:.32,fill:{color:B.accent1}}},
 {text:{text:CLASS,options:{x:3.4,y:7.15,w:5.0,h:.32,fontSize:9,bold:true,color:'FFFFFF',align:'center',valign:'middle',margin:0}}},
 {text:{text:'Air Assault!',options:{x:10.6,y:7.14,w:2.5,h:.34,fontSize:12,bold:true,italic:true,color:B.accent2,align:'right',valign:'middle',margin:0}}},
 {text:{text:'Dog Company · 2-506 IN · 3BCT, 101st ABN DIV (AASLT)',options:{x:.25,y:7.14,w:3.1,h:.34,fontSize:8,color:'FFFFFF',valign:'middle',margin:0}}}];
pres.defineSlideMaster({title:'HQ',background:{color:'FFFFFF'},objects:[...chrome,
  {placeholder:{options:{name:'title',type:'title',x:3.17,y:.22,w:7.0,h:.58,fontSize:24,bold:true,color:'000000',align:'center',valign:'middle',fill:{color:'F2F2F2'},line:{color:'000000',width:1.5},shadow:{type:'outer',blur:3,offset:2,angle:45,color:'000000',opacity:.35}},text:''}}],
  slideNumber:{x:12.5,y:6.85,w:.6,h:.25,fontSize:9,color:B.accent5,align:'right'}});
pres.defineSlideMaster({title:'HQBACK',background:{color:'FFFFFF'},objects:[...chrome,
  {placeholder:{options:{name:'title',type:'title',x:3.17,y:.22,w:7.0,h:.58,fontSize:24,bold:true,color:'000000',align:'center',valign:'middle',fill:{color:'F2F2F2'},line:{color:'000000',width:1.5},shadow:{type:'outer',blur:3,offset:2,angle:45,color:'000000',opacity:.35}},text:''}},
  {text:{text:'BACKUP',options:{x:10.5,y:.33,w:1.5,h:.34,fontSize:10,bold:true,color:B.accent3,align:'center',valign:'middle',margin:0,line:{color:B.accent3,width:1}}}}],
  slideNumber:{x:12.5,y:6.85,w:.6,h:.25,fontSize:9,color:B.accent5,align:'right'}});
pres.defineSlideMaster({title:'HQPLAIN',background:{color:'FFFFFF'},objects:[...chrome]});
// helpers
const title=(sl,t)=>sl.addText(t,{placeholder:'title'});
const box=(sl,x,y,w,h,label,body,opt)=>{opt=opt||{};sl.addShape(pres.ShapeType.rect,{x,y,w,h,fill:{color:opt.fill||'EDEDED'},line:{color:'000000',width:1},shadow:{type:'outer',blur:2,offset:1.5,angle:45,color:'000000',opacity:.3},objectName:'Box'});
  const runs=[];if(label)runs.push({text:label,options:{bold:true,underline:true,breakLine:true,paraSpaceAfter:4}});
  if(Array.isArray(body))runs.push(...body);else if(body)runs.push({text:body});
  sl.addText(runs,{x:x+.12,y:y+.08,w:w-.24,h:h-.16,fontSize:opt.size||12,color:'000000',margin:0,valign:opt.valign||'top',align:opt.align||'left',isTextBox:true,objectName:'Box text'})};
const bulletsRuns=(items,bodySize)=>{const runs=[];items.forEach((it,i)=>{const m=/^([^.]+\.)\s(.*)$/.exec(it);if(m){runs.push({text:m[1]+' ',options:{bold:true,bullet:true,paraSpaceAfter:5}});runs.push({text:m[2],options:{breakLine:i<items.length-1}})}else runs.push({text:it,options:{bullet:true,paraSpaceAfter:5,breakLine:i<items.length-1}})});return runs};
const bbox=(sl,x,y,w,h,label,items,opt)=>box(sl,x,y,w,h,label,bulletsRuns(items),opt);
const numRuns=items=>{const runs=[];items.forEach((it,i)=>{runs.push({text:(i+1)+'. ',options:{bold:true}});const m=/^([^.]+\.)\s(.*)$/.exec(it);if(m){runs.push({text:m[1]+' ',options:{bold:true}});runs.push({text:m[2],options:{breakLine:i<items.length-1,paraSpaceAfter:5}})}else runs.push({text:it,options:{breakLine:i<items.length-1,paraSpaceAfter:5}})});return runs};
const stat=(sl,x,y,w,big,label)=>{sl.addShape(pres.ShapeType.rect,{x,y,w,h:1.05,fill:{color:'FFFFFF'},line:{color:'000000',width:1},objectName:'Stat box'});sl.addText(big,{x,y:y+.08,w,h:.55,fontSize:24,bold:true,color:B.accent4,align:'center',margin:0,isTextBox:true,objectName:'Stat'});sl.addText(label,{x:x+.1,y:y+.62,w:w-.2,h:.38,fontSize:9.5,color:B.accent5,align:'center',valign:'top',margin:0,isTextBox:true,objectName:'Stat label'})};
const pic=(sl,n,x,y,w,h)=>{sl.addImage({path:FR+'c'+String(n).padStart(2,'0')+'.jpg',x,y,w,h,sizing:{type:'cover',w,h},objectName:'Scene '+n});sl.addShape(pres.ShapeType.rect,{x,y,w,h,fill:{color:'000000',transparency:100},line:{color:'000000',width:1},objectName:'Frame'})};
const vid=(sl,n,x,y,w,h,label)=>{const mp=VD+'s'+String(n).padStart(2,'0')+'.mp4';
  if(fs.existsSync(mp)){const cover='image/jpeg;base64,'+fs.readFileSync(SC+'vid/s'+String(n).padStart(2,'0')+'/f00000.jpg').toString('base64');sl.addMedia({type:'video',path:mp,cover,x,y,w,h,objectName:'Animation '+n});sl.addShape(pres.ShapeType.rect,{x,y,w,h,fill:{color:'000000',transparency:100},line:{color:'000000',width:1},objectName:'Frame'});
    sl.addText('Click the image to play the animation',{x,y:y+h+.03,w,h:.22,fontSize:8,italic:true,color:B.accent5,margin:0,isTextBox:true,objectName:'Video hint'})}
  else{sl.addImage({path:FR+'f'+String(n).padStart(2,'0')+'.jpg',x,y,w,h,sizing:{type:'cover',w,h},objectName:'Scene '+n});sl.addShape(pres.ShapeType.rect,{x,y,w,h,fill:{color:'000000',transparency:100},line:{color:'000000',width:1},objectName:'Frame'})}
  if(label)sl.addText(label,{x,y:y-.3,w,h:.26,fontSize:10,bold:true,color:'000000',margin:0,isTextBox:true,objectName:'Figure label'})};
let sl;
// ===== 1 Title
pres.addSection({title:'Opening'});
sl=pres.addSlide({masterName:'HQPLAIN',sectionTitle:'Opening'});
sl.addImage({path:AS+'tgfd.png',x:4.4,y:1.25,w:4.53,h:4.8,transparency:78,objectName:'Crest'});
sl.addText('J-FARP SEIZURE',{x:1,y:2.3,w:11.33,h:1.1,fontSize:48,bold:true,color:'000000',align:'center',margin:0,isTextBox:true,objectName:'Title'});
sl.addText('Concept brief · Ground Task Force mission set',{x:1,y:3.4,w:11.33,h:.5,fontSize:20,bold:true,color:'000000',align:'center',margin:0,isTextBox:true,objectName:'Subtitle'});
sl.addText('Dog Co · 2-506 IN · OCT 26',{x:1,y:3.95,w:11.33,h:.5,fontSize:20,bold:true,color:'000000',align:'center',margin:0,isTextBox:true,objectName:'Unit'});
sl.addNotes(cap(1));
// ===== 2 BLUF
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Opening'});title(sl,'BLUF');
box(sl,.5,1.25,12.33,1.15,null,[{text:'Bottom line: ',options:{bold:true,underline:true}},{text:'A rifle company (-) can seize and secure a jump FARP 180–200 NM forward and hold it for one turn of the attack element, putting Division deep targets 100–150 NM beyond it inside attack aviation combat radius.'}],{size:15,valign:'middle'});
stat(sl,.5,2.6,2.9,'3 hours','on the ground, wheels down to extraction');stat(sl,3.65,2.6,2.9,'82 PAX','Lift 1 · 2x CH-47 · 2x UH-60');stat(sl,6.8,2.6,2.9,'H + 2:00','FARP operations complete');stat(sl,9.95,2.6,2.88,'28–29 OCT 26','proof of concept · Operation Storm Surge');
box(sl,.5,3.9,7.9,2.95,'Decisions this brief supports',numRuns(['Endorse. J-FARP seizure as a Ground Task Force mission set for the deep attack.','Resource. The full-distance proof of concept, Operation Storm Surge, 28–29 OCT 26.','Transition. On validation, hand the mission set to the Division pathfinder company.']),{size:15});
box(sl,8.6,3.9,4.23,2.95,'Agenda',numRuns(['Mission and problem','Concept of the operation','Task organization and phasing','Execution by phase','Risk and proof of concept']),{size:13});
sl.addNotes('Bottom line up front, the numbers that matter, and the decisions this brief supports.');
// ===== 3 Mission
pres.addSection({title:'Concept'});
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Concept'});title(sl,'GTF Mission Statement');
box(sl,.8,1.55,11.73,1.55,null,'On order, an Infantry Company (-) conducts Air Assault Operations to seize a J-FARP to enable Combat Aviation Assets to mass combat power in the Division Deep Area to defeat enemy forces and seize key terrain.',{size:20,valign:'middle',align:'center'});
box(sl,.8,3.4,5.7,3.4,'Purpose',[{text:'Place Class III and Class V inside the Division deep area for one turn of the attack element, so the deep attack launches with full fuel and ordnance against targets that are otherwise beyond combat radius.'}],{size:14});
bbox(sl,6.83,3.4,5.7,3.4,'Key tasks',['Seize the FARP site and the objectives covering its approaches.','Establish the rapid refuel and rearm point from the WETCOW and FATCOW.','Turn the attack element forward with full fuel and ordnance.','Defend the site for the duration of FARP operations.','Close the site and extract through the PZ.'],{size:14});
sl.addNotes(cap(5));
// ===== 4 Problem
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Concept'});title(sl,'The Problem: Operational Reach');
vid(sl,2,.5,1.5,7.1,3.99,'Attack aviation from the ISB reaches bingo short of the Division deep target');
box(sl,7.85,1.25,4.98,2.35,'The problem',[{text:'Attack aviation combat radius is bounded by fuel. Ordnance, reserve and time on station draw on the same tank. From the intermediate staging base the Division deep target is out of reach.'}],{size:12.5});
stat(sl,7.85,3.75,2.4,'180–200 NM','ISB to J-FARP site');stat(sl,10.43,3.75,2.4,'100–150 NM','J-FARP to DIV deep target');
box(sl,7.85,4.95,4.98,1.85,'Without fuel forward',[{text:'The attack element turns for home short of the objective. Terrain beyond that line stays unshaped until Class III is positioned forward.'}],{size:12.5});
box(sl,.5,5.75,7.1,1.05,null,[{text:'With a J-FARP: ',options:{bold:true}},{text:'refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius, with full station time on arrival.'}],{size:12.5,valign:'middle'});
sl.addNotes(cap(2));
// ===== 5 Concept of the operation
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Concept'});title(sl,'Concept of the Operation');
vid(sl,5,.5,1.5,7.1,3.99,'Three lifts: Ground Task Force seizes, FARP Task Force establishes, attack element turns forward');
box(sl,7.85,1.25,4.98,1.75,'Lift 1 · Ground Task Force',[{text:'2x CH-47 · 2x UH-60 · 82 PAX',options:{bold:true,breakLine:true}},{text:'GTF air assault. FARP seizure. Pathfinder operations.'}],{size:13});
box(sl,7.85,3.15,4.98,1.75,'Lift 2 · FARP Task Force',[{text:'2x CH-47 · WETCOW · FATCOW · 18 PAX',options:{bold:true,breakLine:true}},{text:'FARP TF air assault. FARP establishment by ECHO / 6-101.'}],{size:13});
box(sl,7.85,5.05,4.98,1.75,'Lift 3 · Deep Attack',[{text:'2x AH-64',options:{bold:true,breakLine:true}},{text:'Refuel and rearm at the J-FARP with rotors turning, continue to the DIV deep target (100–150 NM).'}],{size:13});
box(sl,.5,5.75,7.1,1.05,'End state','FARP operations complete NLT H+2:00. Attack element over the deep target with full fuel. Site closed, force reconsolidated at the ISB by H+3:00.',{size:11.5});
sl.addNotes(cap(5));
// ===== 6 Task organization
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Concept'});title(sl,'Proposed Ground Task Force');
pic(sl,6,.5,1.25,5.6,5.55);
const el=[['Element 1 · Interdict · 54 PAX','Occupies OBJ Husky, Poodle and Bernese, the objectives covering the avenues of approach to the site. Interdicts enemy forces approaching the FARP.'],['Element 2 · Seize · 12 PAX','Seizes the FARP site from LZ Vulture. Reports the site secured NLT H+1:00.'],['Element 3 · Secure · 9 PAX','Local security on the aircraft and refuel points for the duration of FARP operations. Pathfinders mark PZ Cedar and the refuel point.']];
el.forEach((e,i)=>box(sl,6.35,1.25+i*1.9,6.48,1.75,e[0],e[1],{size:14}));
sl.addNotes(cap(6));
// ===== 7 Phasing
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Concept'});title(sl,'Phasing and H-Hour Sequence');
const ms=[['H − 2:30','Wheels up'],['H-hour','Wheels down'],['H + 1:00','FARP seized'],['H + 2:00','FARP ops complete'],['H + 3:00','Reconsolidated']];
sl.addShape(pres.ShapeType.line,{x:1.2,y:1.95,w:10.9,h:0,line:{color:'000000',width:2},objectName:'Timeline'});
ms.forEach((m,i)=>{const x=1.2+i*2.725;sl.addShape(pres.ShapeType.ellipse,{x:x-.12,y:1.83,w:.24,h:.24,fill:{color:B.accent1},line:{color:'000000',width:1},objectName:'Milestone'});
  sl.addText(m[0],{x:x-1.2,y:1.35,w:2.4,h:.4,fontSize:15,bold:true,color:'000000',align:'center',margin:0,isTextBox:true,objectName:'Milestone time'});
  sl.addText(m[1],{x:x-1.2,y:2.15,w:2.4,h:.3,fontSize:10,color:B.accent5,align:'center',margin:0,isTextBox:true,objectName:'Milestone label'})});
const ph=[['Phase I · Condition setting','B: Planning start  E: Wheels up','Air assault planning (IPC, AMCM, AMB). FARP site selection. Cold loads, rehearsals, leader certification.'],['Phase II · Air assault and seizure','B: Wheels down (H-hour)  E: FARP seized (H+1:00)','Lift 1 lands three LZs. Objectives occupied. Site seized and PZ marked.'],['Phase III · FARP operations','B: H+1:00  E: FARP ops complete (H+2:00)','Lift 2 establishes the point. Lift 3 refuels, rearms and departs for the deep target.'],['Phase IV · FARP defense','Concurrent with Phase III','Ground force holds the perimeter for the window the aircraft are on the ground. Site closed and force extracted by H+3:00.']];
ph.forEach((p,i)=>box(sl,.5+i*3.1,2.7,2.98,4.1,p[0],[{text:p[1],options:{italic:true,breakLine:true,paraSpaceAfter:6}},{text:'Key tasks: ',options:{bold:true}},{text:p[2]}],{size:13}));
sl.addNotes(cap(8));
// ===== 8 Phase II
pres.addSection({title:'Execution'});
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Execution'});title(sl,'Phase II: Air Assault and Seizure');
vid(sl,10,.5,1.5,7.1,3.99,'H-hour to H+1:00 · Lift 1 lands, objectives occupied, PZ Cedar marked');
box(sl,7.85,1.25,4.98,1.1,'Phase II',[{text:'B: Wheels down (H-hour)   E: FARP seized (H+1:00)',options:{breakLine:true}},{text:'Key tasks: ',options:{bold:true}},{text:'GTF air assault, FARP seizure, pathfinder operations.'}],{size:11.5});
bbox(sl,7.85,2.5,4.98,4.3,'Scheme of maneuver',['H-hour. Three serials land LZ Vulture, Owl and Falcon on separate headings.','Seize. Element 2 seizes the FARP site from LZ Vulture.','Interdict. Element 1 occupies OBJ Husky, Poodle and Bernese, oriented on the avenues of approach.','Establish. Element 3 secures the site. Pathfinders mark PZ Cedar and the refuel point inside the triangle. Secured NLT H+1:00.'],{size:13});
box(sl,.5,5.75,7.1,1.05,null,[{text:'Night air assault, dismounted land navigation, seize. ',options:{bold:true}},{text:'Chalk-to-LZ assignments shown are the Storm Surge AMB plan; the CONOP Lift 1 package is 2x CH-47 and 2x UH-60.'}],{size:11.5,valign:'middle'});
sl.addNotes(cap(10));
// ===== 9 Phase III
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Execution'});title(sl,'Phase III: FARP Operations');
vid(sl,11,.5,1.5,7.1,3.99,'H+1:00 to H+2:00 · WETCOW and FATCOW land, Apaches refuel, rearm and turn forward');
box(sl,7.85,1.25,4.98,1.1,'Phase III',[{text:'B: FARP seized (H+1:00)   E: FARP ops complete (H+2:00)',options:{breakLine:true}},{text:'Key tasks: ',options:{bold:true}},{text:'FARP TF air assault, FARP establishment, deep attack air assault.'}],{size:11.5});
bbox(sl,7.85,2.5,4.98,3.1,'Scheme of maneuver',['H+1:00. Lift 2 lands the 18-PAX FARP crew: WETCOW with fuel, FATCOW with AH-64 munitions.','Establish. Refuel points from the WETCOW, the rearm point from the FATCOW. ECHO / 6-101 operates the points.','Turn. Lift 3 lands, refuels and rearms with rotors turning, and departs for the deep target with full fuel.'],{size:13});
stat(sl,7.85,5.75,2.4,'60 NM','AH reach added by the WETCOW');stat(sl,10.43,5.75,2.4,'185 NM','extended range, no refuel, ERFS');
box(sl,.5,5.75,7.1,1.05,null,[{text:'Pathfinders control the zone. ',options:{bold:true}},{text:'Element 3 secures the aircraft and the points. FARP operations complete NLT H+2:00.'}],{size:11.5,valign:'middle'});
sl.addNotes(cap(11));
// ===== 10 Phase IV
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Execution'});title(sl,'Phase IV: FARP Defense');
vid(sl,12,.5,1.5,7.1,3.99,'Objectives engage along the avenues of approach while aircraft are on the ground');
box(sl,7.85,1.25,4.98,1.1,'Phase IV',[{text:'Concurrent with FARP operations',options:{breakLine:true}},{text:'Key task: ',options:{bold:true}},{text:'FARP defense. ATP 3-04.17: the FARP is a high-value, high-signature target with minimal organic security. Security is a ground force task.'}],{size:11});
bbox(sl,7.85,2.5,4.98,4.3,'Security posture',['All-round security. A mobile enemy can approach from any direction. 360 degrees, no element split.','Early warning in depth. OPs on the avenues, then the objectives, then the centre.','Mutual support. Any approach under fire from two positions. The triangular patrol base delivers this with three elements.','Organic fires. Mortar section registered on the avenues.','Limited exposure. The FARP displaces or closes before the enemy can mass.'],{size:12.5});
box(sl,.5,5.75,3.45,1.05,'PACE','Wickr over TACSAT-IW · TAK / ICE · Signal over MUOS · cell or HF.',{size:10.5});
box(sl,4.15,5.75,3.45,1.05,'Displacement criteria','Site compromised, or the alternate better positioned for the next turn.',{size:10.5});
sl.addNotes(cap(12));
// ===== 11 Risk
pres.addSection({title:'Risk and validation'});
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Risk and validation'});title(sl,'Risk and Mitigation');
const rk=[['Risk to force: distance from sustainment','A company (-) with limited sustainment 180 NM from FCKY. Casualty evacuation and resupply are aviation tasks that must be planned, not assumed.'],['Risk to mission: high-signature fuel point','Rotors turning, fuel flowing, aircraft on the ground on a schedule. The window is the vulnerability.'],['Training risk: newly qualified pathfinders','The company assumes risk training with newly qualified pathfinders; a Division pathfinder company is the intended recipient of the mission set.'],['Mitigation: Class I carried','1x DOS of CL I issued prior to W/U and 1x DOS carried per soldier on the aircraft.'],['Mitigation: deliberate rehearsals','Air assault planning with 6-101 CAB, cold-load training on the airframes and leader certification before execution.'],['Mitigation: proof of concept','Validate the task force on a training objective at full distance before the concept is relied on for a deep attack.']];
rk.forEach((r,i)=>box(sl,.5+(i%3)*4.15,1.3+Math.floor(i/3)*2.8,4.03,2.6,r[0],r[1],{size:14,fill:i<3?'F8E0E0':'E2EFDA'}));
sl.addNotes(cap(14));
// ===== 12 Proof of concept
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Risk and validation'});title(sl,'Proof of Concept: Operation Storm Surge');
pic(sl,15,.5,1.25,6.4,3.6);
box(sl,7.15,1.25,5.68,1.6,'Mission','Dog Company conducts a Company (-) Air Assault into Tullahoma, TN to secure an expeditionary FARP on 29 0130Z OCT 26 to enable freedom of movement for 6-101 CAB and proof the concept for FARP seizure operations.',{size:11.5});
bbox(sl,7.15,3.0,5.68,2.5,'Deliverables',['Extended range: 185 NM with no refuel, ERFS employment on Serial 1.','WETCOW enables AH operations 60 NM beyond the objective.','3x WETCOWs, 5x FARP locations, simultaneous, secured and dispersed.','Air-deliverable Shadow TAC.'],{size:12.5});
const tl=[['2200Z','Wheels up'],['0030Z','W/D · H-hour'],['0130Z','FARP seized'],['0230Z','FARP ops complete'],['0330Z','Reconsolidated']];
tl.forEach((m,i)=>box(sl,.5+i*1.3,5.05,1.2,.95,null,[{text:m[0],options:{bold:true,breakLine:true,fontSize:13}},{text:m[1],options:{fontSize:8.5}}],{size:9,align:'center',valign:'middle'}));
box(sl,.5,6.15,6.4,.65,null,[{text:'A/C package: ',options:{bold:true}},{text:'3x CH-47 (1x PAX, 1x FATCOW) · 1x UH-60 · 83 PAX · night · CAAF to Tullahoma, TN · 184 NM'}],{size:10.5,valign:'middle'});
box(sl,7.15,5.65,5.68,1.15,null,[{text:'Training objectives: ',options:{bold:true}},{text:'07-CO-5135 Operate a company command post · 011-60Z-1054 Select a landing zone / pickup zone.'}],{size:10.5,valign:'middle'});
sl.addNotes(cap(15));
// ===== 13 End state
sl=pres.addSlide({masterName:'HQ',sectionTitle:'Risk and validation'});title(sl,'End State and Decisions');
vid(sl,4,.5,1.5,7.1,3.99,'The attack element on the Division deep target with full fuel and ordnance');
box(sl,7.85,1.25,4.98,2.1,'End state',[{text:'Fuel forward is reach forward. ',options:{bold:true}},{text:'An infantry company (-), three hours on the ground, extends the Division deep fight by 200 nautical miles for one turn of the attack element. Site closed, force reconsolidated with 100 percent accountability.'}],{size:13.5});
box(sl,7.85,3.5,4.98,3.3,'Decisions requested',numRuns(['Endorse J-FARP seizure as a Ground Task Force mission set.','Resource Operation Storm Surge, 28–29 OCT 26.','On validation, transition the mission set to the Division pathfinder company.']),{size:14});
box(sl,.5,5.75,7.1,1.05,null,[{text:'Seize the site. Establish the FARP. Turn the deep attack. Close the site.',options:{bold:true}}],{size:14,valign:'middle',align:'center'});
sl.addNotes(cap(16));
// ===== Backup
pres.addSection({title:'Backup'});
sl=pres.addSlide({masterName:'HQPLAIN',sectionTitle:'Backup'});
sl.addText('BACKUP',{x:1,y:2.7,w:11.33,h:1.0,fontSize:44,bold:true,color:'000000',align:'center',margin:0,isTextBox:true,objectName:'Divider'});
sl.addText('Purpose · FARP doctrine · Reverse planning · Site selection · Reconsolidation · Terms and sources',{x:1,y:3.75,w:11.33,h:.5,fontSize:14,color:B.accent5,align:'center',margin:0,isTextBox:true,objectName:'Divider sub'});
sl.addNotes('Backup slides.');
// B1 Purpose
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'Purpose: The Division Deep Area');
pic(sl,4,.5,1.25,5.6,5.55);
box(sl,6.35,1.25,6.48,1.6,null,'The Division deep area is the terrain beyond the close fight, where aviation shapes enemy forces before the main effort arrives. Reaching it is a sustainment problem before it is a tactical one.',{size:12.5});
bbox(sl,6.35,3.0,6.48,2.5,'What a forward FARP buys',['Reach. Refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius.','Station time. Aircraft arrive on the objective with a full fuel load rather than a reserve.','Tempo. Aircraft turn forward at the FARP in minutes instead of returning hours to the ISB.','Mass. More sorties on the objective within the same window.'],{size:12});
box(sl,6.35,5.65,6.48,1.15,null,[{text:'For the 101st, the J-FARP is the enabler for the large-scale, long-range air assault. The deep attack sets conditions on the follow-on objective. The J-FARP sets conditions for the deep attack.',options:{italic:true}}],{size:11.5,valign:'middle'});
sl.addNotes(cap(4));
// B2 Doctrine
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'FARP Variants in Doctrine');
box(sl,.5,1.25,12.33,.95,null,[{text:'FM 3-04 and ATP 3-04.17: ',options:{bold:true}},{text:'a temporary facility providing fuel and ammunition to aviation units in combat. Positioning Class III and Class V forward reduces turnaround distance and extends time in the fight. Four variants trade permanence for responsiveness.'}],{size:12.5,valign:'middle'});
const v=[['Active · manned and operating','Equipment emplaced, crews in position, refueling in progress. The standard FARP supporting sustained operations.'],['Silent · established, not operating','Emplaced and concealed. Activated only when aircraft are inbound, to limit signature until required.'],['Jump · minimal footprint, air-delivered','Rapidly displaceable, with Class III and Class V delivered on the supporting aircraft. Established for a single deep operation, then displaced or closed.'],['Rolling · displaces with the force','Moves by bounds behind the maneuvering unit to keep aviation support within reach of the advance.']];
v.forEach((c,i)=>box(sl,.5+(i%2)*6.25,2.4+Math.floor(i/2)*2.25,6.08,2.05,c[0],c[1],{size:12.5,fill:i===2?'FFF2CC':'EDEDED'}));
sl.addNotes(cap(3));
// B3 Reverse planning
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'Reverse Planning Sequence');
box(sl,.5,1.25,12.33,.95,null,'Air assault doctrine develops five plans in reverse sequence, each constraining the next. For a FARP seizure the ground tactical plan carries a requirement most air assaults do not: the site must accept a CH-47 with internal fuel aboard.',{size:12.5,valign:'middle'});
const steps=[['5','Staging plan','Assembly, sequencing and PZ posture.'],['4','Loading plan','Chalk and load assignments by aircraft.'],['3','Air movement plan','Routes, timing and formation.'],['2','Landing plan','Zone, heading and landing sequence.'],['1','Ground tactical plan','Actions on the objective.']];
steps.forEach((s,i)=>{const x=.5+i*2.5;box(sl,x,2.45,2.3,1.6,s[0]+' · '+s[1],s[2],{size:11.5,fill:i===4?'FFF2CC':'EDEDED'});if(i<4)sl.addShape(pres.ShapeType.line,{x:x+2.32,y:3.25,w:.16,h:0,line:{color:'000000',width:1.5,endArrowType:'triangle'},objectName:'Arrow'})});
const pc=[['Planning events · IPC, AMCM, AMB','Initial planning conference, air mission coordination meeting and air mission brief. One integrated plan between the aviation and ground task forces.'],['Pathfinders · LZ / PZ control','Select, mark and control landing and pickup zones, provide terminal guidance and conduct PZ operations for extraction.'],['Rehearsals · cold loads and leader certification','Cold-load training on the airframes, leader certification and deliberate rehearsals mitigate operating 180 NM from the ISB.']];
pc.forEach((c,i)=>box(sl,.5+i*4.15,4.3,4.03,2.5,c[0],c[1],{size:12}));
sl.addNotes(cap(7));
// B4 Site selection
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'Phase I: FARP Site Selection');
pic(sl,9,.5,1.25,5.6,5.55);
box(sl,6.35,1.25,6.48,.8,null,[{text:'Key tasks: ',options:{bold:true}},{text:'air assault planning and FARP site selection. The site must satisfy three requirements concurrently.'}],{size:12,valign:'middle'});
bbox(sl,6.35,2.2,6.48,4.6,'Site requirements',['Aircraft. Dimensions and surface for a CH-47 with internal fuel aboard. Approach and departure headings clear of obstacles. Slope and dust within limits for night landing.','Target. Close enough to the Division deep target to preserve station time. Far enough that the fuel point is not on the objective.','Defense. Limited avenues of approach, cover for the security elements and an alternate site for displacement. Storm Surge plans three WETCOWs across five FARP locations.'],{size:12.5});
sl.addNotes(cap(9));
// B5 Reconsolidation
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'Reconsolidation: H+2:00 to H+3:00');
pic(sl,13,.5,1.25,5.6,5.55);
box(sl,6.35,1.25,6.48,.95,null,'Reconsolidation distinguishes a jump FARP from a base. The fuel point exists for one turn of the attack element and is then closed.',{size:12,valign:'middle'});
bbox(sl,6.35,2.35,6.48,4.45,'Sequence',['H+2:00. FARP operations complete. Refuel and rearm points recovered. WETCOW and FATCOW lift first.','PZ operations. Pathfinders control the pickup: chalks staged, aircraft inbound 090, staggered left, weapons tight.','Collapse in reverse order. Objectives fold in last, under mortar coverage until the mortar section itself lifts.','H+3:00. Reconsolidated at FCKY with 100 percent accountability. FARP seizure task force validated.'],{size:12.5});
sl.addNotes(cap(13));
// B6 Terms
sl=pres.addSlide({masterName:'HQBACK',sectionTitle:'Backup'});title(sl,'Terms and Sources');
const terms=[['J-FARP','Jump forward arming and refueling point'],['ISB','Intermediate staging base'],['GTF','Ground task force'],['WETCOW','CH-47 carrying fuel tanks: the Class III point'],['FATCOW','CH-47 carrying munitions for the AH-64: the Class V point'],['ERFS','Extended range fuel system'],['R3P','Rapid refuel point'],['L2A2','Large-scale, long-range air assault'],['IPC · AMCM · AMB','Initial planning conference · air mission coordination meeting · air mission brief'],['W/U · W/D','Wheels up · wheels down'],['OP · CCP','Observation post · casualty collection point'],['NOMAD','Unmanned ground vehicle carrying the C2 package'],['MDS','Mission design series']];
const truns=[];terms.forEach((t,i)=>{truns.push({text:t[0]+'  ',options:{bold:true}});truns.push({text:t[1],options:{breakLine:i<terms.length-1,paraSpaceAfter:4}})});
box(sl,.5,1.25,6.0,5.55,'Terms',truns,{size:11});
box(sl,6.75,1.25,6.08,5.55,'Sources',[{text:'Doctrine. ',options:{bold:true}},{text:'FM 3-04 Army Aviation, ATP 3-04.17 Techniques for Forward Arming and Refueling Points, FM 3-99 Airborne and Air Assault Operations, FM 3-90 Tactics, ATP 3-21.8 Infantry Platoon and Squad, TC 3-21.76 Ranger Handbook. Doctrinal statements are paraphrased and should be verified against current editions.',options:{breakLine:true,paraSpaceAfter:8}},{text:'Unit sources. ',options:{bold:true}},{text:'Mission, task organization, lift packages, phasing, timeline and imagery are from the Dog Company J-FARP Seizure CONOP and the Operation Storm Surge AMCM / AMB of 06 OCT 26.',options:{breakLine:true,paraSpaceAfter:8}},{text:'Imagery. ',options:{bold:true}},{text:'Scene renders and animations are from the animated brief. The wide-area terrain is illustrative and not real geography; objective-area scenes are drawn over the unit\'s overhead imagery. GTF elements and strengths are as directed by the GTF commander; chalk-to-LZ and objective assignments follow the Storm Surge AMB.'}],{size:11});
sl.addNotes('Terms and sources.');
(async()=>{await pres.writeFile({fileName:OUT});await applyTheme(OUT,THEME);console.log('written')})();
