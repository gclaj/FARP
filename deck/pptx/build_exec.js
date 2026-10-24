const pptxgen=require('pptxgenjs');const fs=require('fs');const path=require('path');
const {applyTheme}=require('/root/.claude/skills/synced/41bc01b3-6bb5-4167-b63d-2218bfcbb52e_d7e8d10f-d5e7-46e2-89f0-06b91c176297/pptx/scripts/apply_theme.js');
const FR='/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/frames/';
const S=JSON.parse(fs.readFileSync('/tmp/claude-0/-home-user-FARP/2cd27235-ceb3-518d-bee2-e3b544119c0a/scratchpad/slides.json','utf8'));
const cap=n=>S[n-1].cap;
const THEME={name:'J-FARP Night',headFontFace:'Arial',bodyFontFace:'Arial',colors:{dk1:'070B16',lt1:'EEF0F8',dk2:'141A2B',lt2:'AAB2CF',accent1:'F4B94A',accent2:'79D4FF',accent3:'FF7384',accent4:'8BE3B0',accent5:'6B7498',accent6:'FFD98A',hlink:'79D4FF',folHlink:'AAB2CF'}};
const pres=new pptxgen();pres.layout='LAYOUT_WIDE';pres.theme={headFontFace:THEME.headFontFace,bodyFontFace:THEME.bodyFontFace};
pres.author='Dog Company, 2-506 IN';pres.title='J-FARP Seizure - Executive brief';pres.subject='Concept of the operation';
const C=pres.SchemeColor;const W=13.333,H=7.5;
const FOOT='Dog Company · 2-506 IN · 3BCT, 101st ABN DIV (AASLT) · with 6-101 CAB';
const CLASS='UNCLASSIFIED';
const banner=[{text:{text:CLASS,options:{x:0,y:.06,w:W,h:.25,fontSize:9,bold:true,color:THEME.colors.accent4,align:'center',margin:0}}},{text:{text:CLASS,options:{x:0,y:7.2,w:W,h:.25,fontSize:9,bold:true,color:THEME.colors.accent4,align:'center',margin:0}}}];
pres.defineSlideMaster({title:'CONTENT',background:{color:THEME.colors.dk1},objects:[...banner,
  {placeholder:{options:{name:'title',type:'title',x:.6,y:.42,w:12.1,h:.9,fontSize:30,bold:true,color:C.background1,align:'left',margin:0,valign:'top'},text:''}},
  {text:{text:FOOT,options:{x:.6,y:6.92,w:9,h:.3,fontSize:9,color:THEME.colors.accent5,margin:0}}}],
  slideNumber:{x:12.3,y:6.92,w:.5,h:.3,fontSize:9,color:THEME.colors.accent5,align:'right'}});
pres.defineSlideMaster({title:'BACKUP',background:{color:THEME.colors.dk1},objects:[...banner,{placeholder:{options:{name:'title',type:'title',x:.6,y:.42,w:10.6,h:.9,fontSize:30,bold:true,color:C.background1,align:'left',margin:0,valign:'top'},text:''}},{text:{text:'BACKUP',options:{x:11.3,y:.5,w:1.4,h:.34,fontSize:10,bold:true,color:THEME.colors.accent1,align:'center',margin:0,line:{color:THEME.colors.accent1,width:.75}}}},{text:{text:FOOT,options:{x:.6,y:6.92,w:9,h:.3,fontSize:9,color:THEME.colors.accent5,margin:0}}}],
  slideNumber:{x:12.3,y:6.92,w:.5,h:.3,fontSize:9,color:THEME.colors.accent5,align:'right'}});
pres.defineSlideMaster({title:'DARK',background:{color:THEME.colors.dk1},objects:[...banner]});
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
const B=THEME.colors;
const numbered=(sl,items,x,y,w,h,opt)=>{const runs=[];items.forEach((it,i)=>{runs.push({text:(i+1)+'  ',options:{bold:true,color:B.accent1}});const m=/^([^.]+\.)\s(.*)$/.exec(it);if(m){runs.push({text:m[1]+' ',options:{bold:true}});runs.push({text:m[2],options:{breakLine:i<items.length-1,paraSpaceAfter:10}})}else runs.push({text:it,options:{breakLine:i<items.length-1,paraSpaceAfter:10}})});sl.addText(runs,Object.assign({x,y,w,h,fontSize:14,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Numbered'},opt||{}))};
// ===== 1 Title
pres.addSection({title:'Opening'});
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Opening'});img(sl,1,0,0,W,H,{full:true});
sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'070B16',transparency:35},line:{color:'070B16',transparency:100},objectName:'Scrim'});sl.addText(CLASS,{x:0,y:.06,w:W,h:.25,fontSize:9,bold:true,color:B.accent4,align:'center',margin:0,isTextBox:true,objectName:'Classification'});sl.addText(CLASS,{x:0,y:7.2,w:W,h:.25,fontSize:9,bold:true,color:B.accent4,align:'center',margin:0,isTextBox:true,objectName:'Classification'});
sl.addShape(pres.ShapeType.roundRect,{x:.7,y:.9,w:7.6,h:4.1,rectRadius:.1,fill:{color:'0E1322',transparency:18},line:{color:'2A3350',width:.75},objectName:'Panel'});
eyebrow(sl,'Concept brief · Ground Task Force mission set',1.1,1.25,6.8);
sl.addText([{text:'J-FARP ',options:{color:C.background1}},{text:'Seizure.',options:{color:B.accent1,italic:true}}],{x:1.1,y:1.6,w:6.8,h:1.1,fontSize:54,bold:true,margin:0,isTextBox:true,objectName:'Title'});
para(sl,'Air assault seizure of a jump forward arming and refueling point to extend attack aviation into the Division deep area.',1.1,2.8,6.8,1.0,{fontSize:17});
para(sl,'Dog Company, 2-506 IN · 3BCT, 101st ABN DIV (AASLT) · with 6-101 CAB · OCT 2026',1.1,4.2,6.8,.5,{fontSize:12,color:C.background2});
sl.addNotes(cap(1));
// ===== 2 BLUF
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Opening'});title(sl,'Bottom line up front');
sl.addShape(pres.ShapeType.roundRect,{x:.6,y:1.45,w:12.1,h:1.5,rectRadius:.08,fill:{color:B.dk2},line:{color:B.accent1,width:1},objectName:'BLUF panel'});
sl.addText([{text:'A rifle company (-) can seize and secure a jump FARP 180–200 NM forward and hold it for one turn of the attack element, ',options:{}},{text:'putting Division deep targets 100–150 NM beyond it inside attack aviation combat radius.',options:{bold:true,color:B.accent6}}],{x:.9,y:1.6,w:11.5,h:1.2,fontSize:17,color:C.background1,margin:0,valign:'middle',isTextBox:true,objectName:'BLUF'});
stat(sl,.6,3.25,2.8,'3 hours','on the ground, W/D to extraction');stat(sl,3.6,3.25,2.8,'82 PAX','Lift 1 · 2x CH-47 · 2x UH-60');stat(sl,6.6,3.25,2.8,'H + 2:00','FARP operations complete');stat(sl,9.6,3.25,3.1,'28–29 OCT','proof of concept · Storm Surge');
eyebrow(sl,'Decisions this brief supports',.6,4.7,8);
numbered(sl,['Endorse. J-FARP seizure as a Ground Task Force mission set for the deep attack.','Resource. The full-distance proof of concept, Operation Storm Surge, 28–29 OCT 26.','Transition. On validation, hand the mission set to the Division pathfinder company.'],.6,5.05,8.2,1.8);
card(sl,9.2,4.7,3.5,1.85,'Agenda',null,'Problem · Mission and concept · Task organization · Phasing · Execution by phase · Risk · Proof of concept',{bodySize:12.5});
sl.addNotes('Bottom line up front. The concept, the numbers that matter, and the decisions this brief supports. '+cap(1));
// ===== 3 Problem
pres.addSection({title:'Concept'});
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'Deep targets lie beyond combat radius');
eyebrow(sl,'The problem · operational reach',.6,1.35,6);
img(sl,2,6.9,1.7,5.85,5.0);
para(sl,'Attack aviation combat radius is bounded by fuel. Ordnance, reserve and time on station draw on the same tank. From the intermediate staging base, the Division deep target is out of reach.',.6,1.75,5.9,1.6,{fontSize:15});
stat(sl,.6,3.55,2.9,'180–200 NM','ISB to J-FARP site');stat(sl,3.6,3.55,2.9,'100–150 NM','J-FARP to DIV deep target');
para(sl,[{text:'Without fuel forward: ',options:{bold:true,color:B.accent6}},{text:'the attack element reaches bingo short of the objective and turns for home. Terrain beyond that line stays unshaped.'}],.6,4.95,5.9,1.5,{fontSize:15});
sl.addNotes(cap(2));
// ===== 4 Mission and concept
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'Seize the site. Establish the FARP. Turn the deep attack');
eyebrow(sl,'Mission and concept of the operation',.6,1.35,6);
sl.addShape(pres.ShapeType.roundRect,{x:.6,y:1.7,w:12.1,h:1.15,rectRadius:.08,fill:{color:B.dk2},line:{color:'2A3350',width:.75},objectName:'Mission panel'});
para(sl,[{text:'Mission.  ',options:{bold:true,color:B.accent6}},{text:'On order, an Infantry Company (-) conducts Air Assault Operations to seize a J-FARP to enable Combat Aviation Assets to mass combat power in the Division Deep Area to defeat enemy forces and seize key terrain.'}],.9,1.82,11.5,.95,{fontSize:14,valign:'middle'});
eyebrow(sl,'Key tasks',.6,3.1,4);
numbered(sl,['Seize the FARP site and the objectives that cover its approaches.','Establish the rapid refuel and rearm point from the WETCOW and FATCOW.','Turn the attack element forward with full fuel and ordnance.','Defend the site for the duration of FARP operations.','Close the site and extract through the PZ.'],.6,3.45,5.6,2.0,{fontSize:13});
eyebrow(sl,'End state',6.6,3.1,4);
para(sl,'FARP operations complete NLT H+2:00. The attack element is over the Division deep target with full fuel. The site is closed, the force reconsolidated at the ISB by H+3:00 with 100 percent accountability.',6.6,3.45,6.1,1.3,{fontSize:13});
const lifts=[['Lift 1 · Ground Task Force','2x CH-47 · 2x UH-60 · 82 PAX',B.accent2],['Lift 2 · FARP Task Force','2x CH-47 · WETCOW · FATCOW · 18 PAX',B.accent1],['Lift 3 · Deep Attack','2x AH-64 · refuel, rearm, continue',B.accent3]];
lifts.forEach((l,i)=>card(sl,.6+i*4.1,5.55,3.95,1.2,l[0],null,l[1],{tagColor:l[2],bodySize:12.5}));
sl.addNotes(cap(5));
// ===== 5 Task org
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'Proposed Ground Task Force');
eyebrow(sl,'Task organization · Lift 1 · 75 PAX in three elements',.6,1.35,8);
img(sl,6,.6,1.8,5.3,5.0);
const el=[['Element 1','Interdict','Occupies the objectives covering the avenues of approach to the site.','54'],['Element 2','Seize','Seizes the FARP site and reports it secured NLT H+1:00.','12'],['Element 3','Secure','Local security on the aircraft and refuel points for the duration of FARP operations.','9']];
el.forEach((e,i)=>card(sl,6.2,1.8+i*1.7,6.55,1.55,e[0],e[1],e[2],{stat:e[3],headSize:16,bodySize:13}));
sl.addNotes(cap(6));
// ===== 6 Phasing
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Concept'});title(sl,'Four phases, three hours on the ground');
eyebrow(sl,'Proposed phasing and H-hour sequence',.6,1.35,6);
const ms=[['H − 2:30','Wheels up'],['H-hour','Wheels down'],['H + 1:00','FARP seized'],['H + 2:00','FARP ops complete'],['H + 3:00','Reconsolidated']];
sl.addShape(pres.ShapeType.line,{x:1.2,y:2.55,w:10.9,h:0,line:{color:'2A3350',width:2},objectName:'Timeline'});
ms.forEach((m,i)=>{const x=1.2+i*2.725;sl.addShape(pres.ShapeType.ellipse,{x:x-.11,y:2.44,w:.22,h:.22,fill:{color:B.accent1},line:{color:B.accent1},objectName:'Milestone'});
  sl.addText(m[0],{x:x-1.2,y:1.95,w:2.4,h:.4,fontSize:16,bold:true,color:C.background1,align:'center',margin:0,isTextBox:true,objectName:'Milestone time'});
  sl.addText(m[1].toUpperCase(),{x:x-1.2,y:2.75,w:2.4,h:.3,fontSize:9,color:C.background2,align:'center',charSpacing:2,margin:0,isTextBox:true,objectName:'Milestone label'})});
const ph=[['Phase I','Condition setting','Air assault planning. FARP site selection. Rehearsals and leader certification.'],['Phase II','Air assault and seizure','Lift 1 lands. Objectives occupied. Site seized and PZ marked NLT H+1:00.'],['Phase III','FARP operations','Lift 2 establishes the point. Lift 3 refuels, rearms and departs. Complete NLT H+2:00.'],['Phase IV','FARP defense','Ground force holds the perimeter for the window the aircraft are on the ground.']];
ph.forEach((p,i)=>card(sl,.6+i*3.08,3.45,2.95,2.5,p[0],p[1],p[2],{headSize:16,bodySize:13,tagColor:B.accent1}));
sl.addNotes(cap(8));
// ===== 7 Phase II
pres.addSection({title:'Execution'});
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Phase II · Land, seize the objectives, establish the PZ');
eyebrow(sl,'H-hour to H + 1:00 · Lift 1',.6,1.35,6);
img(sl,10,.6,1.75,6.2,5.0);
bullets(sl,['H-hour. Three serials land LZ Vulture, Owl and Falcon on separate headings.','Seize. Element 2 seizes the FARP site from LZ Vulture.','Interdict. Element 1 occupies OBJ Husky, Poodle and Bernese, oriented on the avenues of approach.','Establish. Element 3 secures the site. Pathfinders mark PZ Cedar and the refuel point inside the triangle. Secured NLT H+1:00.'],7.1,1.75,5.65,4.4,{fontSize:15});
para(sl,'Chalk-to-LZ assignments shown are the Storm Surge AMB plan.',7.1,6.3,5.65,.4,{fontSize:10,color:C.background2});
sl.addNotes(cap(10));
// ===== 8 Phase III
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Phase III · Establish the point, turn the deep attack');
eyebrow(sl,'H + 1:00 to H + 2:00 · Lifts 2 and 3',.6,1.35,6);
img(sl,11,.6,1.75,6.2,5.0);
bullets(sl,['H+1:00. Lift 2 lands the 18-PAX FARP crew: WETCOW with fuel, FATCOW with AH-64 munitions.','Establish. Refuel points run from the WETCOW, the rearm point from the FATCOW. ECHO / 6-101 operates them.','Turn. Lift 3 lands the points, refuels and rearms with rotors turning, and departs for the deep target with full fuel. Complete NLT H+2:00.'],7.1,1.75,5.65,3.6,{fontSize:15});
stat(sl,7.1,5.5,2.7,'60 NM','AH reach added by the WETCOW');stat(sl,10.0,5.5,2.75,'185 NM','Extended range, no refuel, ERFS');
sl.addNotes(cap(11));
// ===== 9 Phase IV
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Execution'});title(sl,'Phase IV · Security is a ground force task');
eyebrow(sl,'FARP defense · principles of the defense at company scale',.6,1.35,8);
img(sl,12,.6,1.75,5.4,3.6);
para(sl,'ATP 3-04.17 treats the FARP as a high-value, high-signature target with minimal organic security. The ground task force provides it.',.6,5.5,5.4,1.2,{fontSize:13,color:C.background2});
bullets(sl,['All-round security. A mobile enemy can approach from any direction. The perimeter covers 360 degrees with no element split.','Early warning in depth. Observation posts on the avenues, then the objectives, then the centre. Contact is made as far from the aircraft as possible.','Mutual support. Any approach is under fire from two positions. The triangular patrol base delivers this with three elements.','Limited exposure. Time on site is the vulnerability. The FARP displaces or closes before the enemy can mass.'],6.3,1.75,6.45,3.7,{fontSize:13});
card(sl,6.3,5.55,3.15,1.2,'PACE',null,'Primary Wickr over TACSAT-IW. Alternate TAK / ICE. Contingency Signal over MUOS. Emergency cell or HF.',{bodySize:11});
card(sl,9.6,5.55,3.15,1.2,'Displacement criteria',null,'Site compromised, or the alternate site better positioned for the next turn. The FARP moves before it is forced to.',{bodySize:11});
sl.addNotes(cap(12));
// ===== 10 Risk
pres.addSection({title:'Risk and validation'});
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Risk and validation'});title(sl,'Risk and mitigation');
const rk=[['Risk to force','Distance from sustainment','A company (-) with limited sustainment 180 NM from FCKY. Casualty evacuation and resupply are aviation tasks that must be planned, not assumed.',B.accent3],['Risk to mission','High-signature fuel point','Rotors turning, fuel flowing, aircraft on the ground on a schedule. The window is the vulnerability.',B.accent3],['Training risk','Newly qualified pathfinders','The company assumes risk training with newly qualified pathfinders; a Division pathfinder company is the intended recipient of the mission set.',B.accent3],['Mitigation','Class I carried','1× DOS of CL I issued prior to W/U and 1× DOS carried per soldier on the aircraft.',B.accent1],['Mitigation','Deliberate rehearsals','Air assault planning with 6-101 CAB, cold-load training on the airframes and leader certification before execution.',B.accent1],['Mitigation','Proof of concept','Validate the task force on a training objective at full distance before the concept is relied on for a deep attack.',B.accent1]];
rk.forEach((r,i)=>card(sl,.6+(i%3)*4.1,1.5+Math.floor(i/3)*2.65,3.95,2.45,r[0],r[1],r[2],{tagColor:r[3],headSize:16,bodySize:13}));
sl.addNotes(cap(14));
// ===== 11 Proof of concept
sl=pres.addSlide({masterName:'CONTENT',sectionTitle:'Risk and validation'});title(sl,'Proof of concept · Operation Storm Surge');
eyebrow(sl,'28–29 OCT 26 · CAAF to Tullahoma, TN · 184 NM · night',.6,1.35,8);
img(sl,15,.6,1.75,6.2,4.1);
para(sl,[{text:'Mission.  ',options:{bold:true,color:B.accent6}},{text:'Dog Company conducts a Company (-) Air Assault into Tullahoma, TN to secure an expeditionary FARP on 29 0130Z OCT 26 to enable freedom of movement for 6-101 CAB and proof the concept for FARP seizure operations.'}],7.1,1.75,5.65,1.4,{fontSize:13});
bullets(sl,['185 NM with no refuel, ERFS employment on Serial 1.','WETCOW enables AH operations 60 NM beyond the objective.','3× WETCOWs across 5× FARP locations, simultaneous, secured and dispersed.','Package: 3× CH-47 (1× PAX, 1× FATCOW), 1× UH-60, 83 PAX.'],7.1,3.3,5.65,2.5,{fontSize:13});
const tl=[['2200Z','Wheels up'],['0030Z','W/D · H-hour'],['0130Z','FARP seized'],['0230Z','FARP ops complete'],['0330Z','Reconsolidated']];
tl.forEach((m,i)=>{const x=.6+i*1.25;sl.addText(m[0],{x,y:5.95,w:1.22,h:.35,fontSize:14,bold:true,color:B.accent1,margin:0,isTextBox:true,objectName:'Time'});sl.addText(m[1].toUpperCase(),{x,y:6.3,w:1.22,h:.5,fontSize:7.5,color:C.background2,margin:0,valign:'top',isTextBox:true,objectName:'Time label'})});
sl.addNotes(cap(15));
// ===== 12 End state and decisions
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Risk and validation'});img(sl,16,0,0,W,H,{full:true});
sl.addShape(pres.ShapeType.rect,{x:0,y:0,w:W,h:H,fill:{color:'070B16',transparency:40},line:{color:'070B16',transparency:100},objectName:'Scrim'});sl.addText(CLASS,{x:0,y:.06,w:W,h:.25,fontSize:9,bold:true,color:B.accent4,align:'center',margin:0,isTextBox:true,objectName:'Classification'});sl.addText(CLASS,{x:0,y:7.2,w:W,h:.25,fontSize:9,bold:true,color:B.accent4,align:'center',margin:0,isTextBox:true,objectName:'Classification'});
sl.addShape(pres.ShapeType.roundRect,{x:.7,y:.8,w:7.9,h:5.6,rectRadius:.1,fill:{color:'0E1322',transparency:15},line:{color:'2A3350',width:.75},objectName:'Panel'});
eyebrow(sl,'End state',1.1,1.1,6.8);
sl.addText([{text:'Fuel forward is ',options:{color:C.background1}},{text:'reach forward.',options:{color:B.accent1,italic:true}}],{x:1.1,y:1.4,w:7.1,h:1.0,fontSize:38,bold:true,margin:0,isTextBox:true,objectName:'Title'});
para(sl,'An infantry company (-), three hours on the ground, extends the Division deep fight by 200 nautical miles for one turn of the attack element.',1.1,2.85,7.1,.9,{fontSize:15});
eyebrow(sl,'Decisions requested',1.1,3.9,6.8);
numbered(sl,['Endorse J-FARP seizure as a Ground Task Force mission set.','Resource Operation Storm Surge, 28–29 OCT 26.','On validation, transition the mission set to the Division pathfinder company.'],1.1,4.25,7.1,1.9,{fontSize:14});
sl.addText(FOOT,{x:.7,y:6.9,w:9,h:.3,fontSize:10,color:C.background2,margin:0,isTextBox:true,objectName:'Footer'});
sl.addNotes(cap(16));
// ===== Backup
pres.addSection({title:'Backup'});
sl=pres.addSlide({masterName:'DARK',sectionTitle:'Backup'});
sl.addText('Backup',{x:.7,y:2.9,w:12,h:1.1,fontSize:54,bold:true,color:C.background1,margin:0,isTextBox:true,objectName:'Divider'});
sl.addText('Purpose · Doctrine · Reverse planning · Site selection · Reconsolidation · Terms and sources',{x:.7,y:4.05,w:12,h:.5,fontSize:14,color:C.background2,margin:0,isTextBox:true,objectName:'Divider sub'});
sl.addNotes('Backup slides.');
// B1 Purpose
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'A FARP extends operational reach');
eyebrow(sl,'Purpose · the Division deep area',.6,1.35,6);
img(sl,4,6.9,1.7,5.85,5.0);
para(sl,'The Division deep area is the terrain beyond the close fight, where aviation shapes enemy forces before the main effort arrives. Reaching it is a sustainment problem before it is a tactical one.',.6,1.75,5.9,1.1);
bullets(sl,['Reach. Refuel 200 NM forward and targets 300 NM from the ISB fall inside combat radius.','Station time. Aircraft arrive on the objective with a full fuel load rather than a reserve.','Tempo. Aircraft turn forward at the FARP in minutes instead of returning hours to the ISB.','Mass. More sorties on the objective within the same window.'],.6,2.95,5.9,2.5);
para(sl,'For the 101st, the J-FARP is the enabler for the large-scale, long-range air assault. The deep attack sets conditions on the follow-on objective. The J-FARP sets conditions for the deep attack.',.6,5.55,5.9,1.2,{italic:true,color:C.background2});
sl.addNotes(cap(4));
// B2 Doctrine
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'FARP variants in doctrine');
eyebrow(sl,'FM 3-04 · ATP 3-04.17 · forward arming and refueling point',.6,1.35,8);
img(sl,3,.6,1.9,4.6,4.9);
para(sl,'A temporary facility providing fuel and ammunition to aviation units in combat. Positioning Class III and Class V forward reduces turnaround distance and extends time in the fight. Four variants trade permanence for responsiveness.',5.5,1.9,7.2,1.1);
const v=[['Active','Manned and operating','Equipment emplaced, crews in position, refueling in progress. The standard FARP supporting sustained operations.'],['Silent','Established, not operating','Emplaced and concealed. Activated only when aircraft are inbound, to limit signature until required.'],['Jump','Minimal footprint, air-delivered','Rapidly displaceable, with Class III and Class V delivered on the supporting aircraft. Established for a single deep operation, then displaced or closed.'],['Rolling','Displaces with the force','Moves by bounds behind the maneuvering unit to keep aviation support within reach of the advance.']];
v.forEach((c,i)=>card(sl,5.5+(i%2)*3.65,3.15+Math.floor(i/2)*1.85,3.55,1.7,c[0],c[1],c[2],{tagColor:i===2?B.accent1:B.accent2,headSize:14,bodySize:11.5}));
sl.addNotes(cap(3));
// B3 Reverse planning
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'Reverse planning from the ground tactical plan');
eyebrow(sl,'Doctrine · air assault planning',.6,1.35,6);
para(sl,'Five plans developed in reverse sequence, each constraining the next. For a FARP seizure the ground tactical plan carries a requirement most air assaults do not: the site must accept a CH-47 with internal fuel aboard.',.6,1.75,12.1,.9,{fontSize:13});
const steps=[['5','Staging plan','Assembly, sequencing and PZ posture.'],['4','Loading plan','Chalk and load assignments by aircraft.'],['3','Air movement plan','Routes, timing and formation.'],['2','Landing plan','Zone, heading and landing sequence.'],['1','Ground tactical plan','Actions on the objective.']];
steps.forEach((s,i)=>{const x=.6+i*2.47;sl.addShape(pres.ShapeType.ellipse,{x:x,y:2.8,w:.5,h:.5,fill:{color:i===4?B.accent1:B.dk2},line:{color:i===4?B.accent1:'2A3350',width:.75},objectName:'Step'});
  sl.addText(s[0],{x:x,y:2.8,w:.5,h:.5,fontSize:14,bold:true,color:i===4?'070B16':C.background1,align:'center',valign:'middle',margin:0,isTextBox:true,objectName:'Step number'});
  if(i<4)sl.addShape(pres.ShapeType.line,{x:x+.55,y:3.05,w:1.8,h:0,line:{color:'2A3350',width:1,endArrowType:'triangle'},objectName:'Arrow'});
  sl.addText(s[1],{x:x,y:3.4,w:2.3,h:.4,fontSize:14,bold:true,color:C.background1,margin:0,isTextBox:true,objectName:'Step title'});
  sl.addText(s[2],{x:x,y:3.8,w:2.3,h:.6,fontSize:11.5,color:C.background2,margin:0,valign:'top',isTextBox:true,objectName:'Step body'})});
const pc=[['Planning events','IPC · AMCM · AMB','Initial planning conference, air mission coordination meeting and air mission brief. One integrated plan between the aviation and ground task forces.'],['Pathfinders','LZ / PZ control','Select, mark and control landing and pickup zones, provide terminal guidance and conduct PZ operations for extraction.'],['Rehearsals','Cold loads and leader certification','Cold-load training on the airframes, leader certification and deliberate rehearsals mitigate operating 180 NM from the ISB.']];
pc.forEach((c,i)=>card(sl,.6+i*4.1,4.65,3.95,2.15,c[0],c[1],c[2],{headSize:14,bodySize:12}));
sl.addNotes(cap(7));
// B4 Site selection
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'Phase I · Select the site, plan in reverse from it');
eyebrow(sl,'Condition setting · FARP site selection',.6,1.35,6);
img(sl,9,6.9,1.7,5.85,5.0);
para(sl,'The site must satisfy three requirements concurrently.',.6,1.75,5.9,.5);
bullets(sl,['Aircraft. Dimensions and surface for a CH-47 with internal fuel aboard. Approach and departure headings clear of obstacles. Slope and dust within limits for night landing.','Target. Close enough to the Division deep target to preserve station time. Far enough that the fuel point is not on the objective.','Defense. Limited avenues of approach, cover for the security elements and an alternate site for displacement. Storm Surge plans three WETCOWs across five FARP locations.'],.6,2.4,5.9,3.6);
sl.addNotes(cap(9));
// B5 Reconsolidation
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'Reconsolidation · Close the site');
eyebrow(sl,'H + 2:00 to H + 3:00',.6,1.35,6);
img(sl,13,6.9,1.7,5.85,5.0);
para(sl,'Reconsolidation distinguishes a jump FARP from a base. The fuel point exists for one turn of the attack element and is then closed.',.6,1.75,5.9,.9);
bullets(sl,['H+2:00. FARP operations complete. Refuel and rearm points recovered. WETCOW and FATCOW lift first.','PZ operations. Pathfinders control the pickup: chalks staged, aircraft inbound 090, staggered left, weapons tight.','Collapse in reverse order. Objectives fold in last, under mortar coverage until the mortar section itself lifts.','H+3:00. Reconsolidated at FCKY with 100 percent accountability. FARP seizure task force validated.'],.6,2.75,5.9,3.6);
sl.addNotes(cap(13));
// B6 Terms and sources
sl=pres.addSlide({masterName:'BACKUP',sectionTitle:'Backup'});title(sl,'Terms and sources');
const terms=[['J-FARP','Jump forward arming and refueling point'],['ISB','Intermediate staging base'],['GTF','Ground task force'],['WETCOW','CH-47 carrying fuel tanks: the Class III point'],['FATCOW','CH-47 carrying munitions for the AH-64: the Class V point'],['ERFS','Extended range fuel system'],['R3P','Rapid refuel point'],['L2A2','Large-scale, long-range air assault'],['IPC · AMCM · AMB','Initial planning conference · air mission coordination meeting · air mission brief'],['W/U · W/D','Wheels up · wheels down'],['OP · CCP','Observation post · casualty collection point'],['NOMAD','Unmanned ground vehicle carrying the C2 package'],['MDS','Mission design series']];
const truns=[];terms.forEach((t,i)=>{truns.push({text:t[0]+'  ',options:{bold:true,color:B.accent6}});truns.push({text:t[1],options:{breakLine:i<terms.length-1,paraSpaceAfter:5}})});
sl.addText(truns,{x:.6,y:1.5,w:5.6,h:5.3,fontSize:11.5,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Terms'});
const src=[{text:'Doctrine. ',options:{bold:true,color:B.accent6}},{text:'FM 3-04 Army Aviation, ATP 3-04.17 Techniques for Forward Arming and Refueling Points, FM 3-99 Airborne and Air Assault Operations, FM 3-90 Tactics, ATP 3-21.8 Infantry Platoon and Squad and TC 3-21.76 Ranger Handbook. Doctrinal statements are paraphrased and should be verified against current editions.',options:{breakLine:true,paraSpaceAfter:8}},
 {text:'Unit sources. ',options:{bold:true,color:B.accent6}},{text:'Mission, task organization, lift packages, phasing, timeline and imagery are from the Dog Company J-FARP Seizure CONOP and the Operation Storm Surge AMCM / AMB of 06 OCT 26.',options:{breakLine:true,paraSpaceAfter:8}},
 {text:'Imagery. ',options:{bold:true,color:B.accent6}},{text:'Scene renders are from the animated brief. The wide-area terrain is illustrative and not real geography; objective-area scenes are drawn over the unit\'s overhead imagery. Ground Task Force elements and strengths are as directed by the GTF commander; chalk-to-LZ and objective assignments follow the Storm Surge AMB.'}];
sl.addText(src,{x:6.6,y:1.5,w:6.15,h:5.3,fontSize:11.5,color:C.background1,margin:0,valign:'top',isTextBox:true,objectName:'Sources'});
sl.addNotes('Terms and sources.');
(async()=>{await pres.writeFile({fileName:'J-FARP_Seizure_Executive.pptx'});await applyTheme('J-FARP_Seizure_Executive.pptx',THEME);console.log('written')})();
