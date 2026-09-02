'use strict';
const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
const POS=['QB','RB','WR','TE','OT','OG','C','EDGE','DT','LB','CB','S','K','P'];
const POS_COUNTS={QB:4,RB:5,WR:10,TE:5,OT:7,OG:6,C:3,EDGE:7,DT:7,LB:8,CB:8,S:7,K:2,P:2};
const FIRST=['Marcus','Darius','Eli','Jordan','Devin','Trevor','Mason','Jamal','Evan','Aaron','Tyler','Carter','Malik','Isaiah','Noah','Caleb','Andre','Micah','Dante','Logan','Xavier','Bryce','Miles','Cole','Jaylen','Cam','Roman','Nico','Zion','Trey','Dominic','Rashad','Gavin','Khalil','Luke','Jalen','Emmett','Malachi'];
const LAST=['Grant','Brooks','Coleman','Hayes','Wallace','Kemp','Voss','Rourke','Price','Banks','Foster','Bell','Crawford','Morris','King','Lewis','Dawson','Cross','Reed','Marsh','Holloway','Cole','Harper','Bennett','Rivers','Stone','Woods','Bryant','Knox','Mercer','Hawkins','Maddox','Sutton','Parker','Monroe','Whitaker'];
const STYLES={
 QB:['Rhythm Distributor','Field Architect','Off-Script Creator','Vertical Hunter','Power Creator','Run-First Weapon','Toolsy Project','Backyard Magician'],
 RB:['Gap Hammer','Space Cutter','Contact Glider','Third-Down Weapon','One-Cut Burner','Patient Slasher'],
 WR:['Boundary Bully','Space Hunter','Route Sculptor','Vertical Glider','Catch-Point Ace','YAC Creator','Motion Weapon'],
 TE:['Seam Hunter','Inline Mauler','Move Chesspiece','Red-Zone Tower','Backfield Hybrid'],
 OT:['Island Protector','Movement Tackle','Power Edge','Long-Frame Project'],OG:['Drive Blocker','Phone-Booth Anchor','Pulling Guard','Interior Technician'],C:['Protection Quarterback','Leverage Center','Movement Pivot','Power Pivot'],
 EDGE:['Pocket Wrecker','Speed Bender','Power Setter','Chaos Rusher','Run-Side Anchor'],DT:['Pocket Crusher','Gap Eater','Penetrator','Two-Gap Anchor','Interior Disruptor'],LB:['Range Hunter','Box Enforcer','Coverage Rover','Pressure Backer','Traffic Director'],
 CB:['Mirror Corner','Press Eraser','Ball Hunter','Boundary Fighter','Slot Shadow'],S:['Centerfielder','Box Hammer','Coverage Eraser','Hybrid Rover','Trigger Safety'],K:['Pressure Leg','Range Kicker','Placement Specialist'],P:['Field Flipper','Hang-Time Specialist','Directional Punter']
};
const BODY={QB:[[70,78],[185,245]],RB:[[66,73],[175,235]],WR:[[67,78],[165,230]],TE:[[73,80],[220,275]],OT:[[75,81],[275,350]],OG:[[72,78],[280,350]],C:[[71,77],[275,335]],EDGE:[[72,79],[215,285]],DT:[[71,78],[270,350]],LB:[[70,77],[205,255]],CB:[[67,75],[165,215]],S:[[68,77],[180,225]],K:[[68,76],[165,225]],P:[[70,78],[175,230]]};
const OFF_SCHEMES={
 'Tempo Spread':{pass:.61,pace:91,traits:['speed','iq','technique'],qb:['Rhythm Distributor','Off-Script Creator','Run-First Weapon']},
 'Vertical Strike':{pass:.66,pace:73,traits:['speed','power','composure'],qb:['Vertical Hunter','Power Creator','Backyard Magician']},
 'Rhythm Control':{pass:.58,pace:67,traits:['iq','technique','composure'],qb:['Rhythm Distributor','Field Architect']},
 'Ground Pressure':{pass:.37,pace:58,traits:['power','technique','durability'],qb:['Power Creator','Field Architect']},
 'Option Motion':{pass:.35,pace:82,traits:['speed','versatility','iq'],qb:['Run-First Weapon','Off-Script Creator']},
 'Heavy Play Action':{pass:.46,pace:55,traits:['power','iq','technique'],qb:['Field Architect','Vertical Hunter','Power Creator']},
 'Multiple':{pass:.51,pace:68,traits:['versatility','iq','technique'],qb:['Field Architect','Off-Script Creator','Rhythm Distributor']}
};
const DEF_SCHEMES={
 'Pressure Multiple':{pressure:88,coverage:66,run:73,traits:['speed','technique','versatility']},
 'Odd Front Control':{pressure:70,coverage:68,run:88,traits:['power','technique','iq']},
 'Four-Down Attack':{pressure:82,coverage:65,run:80,traits:['power','speed','technique']},
 'Match Quarters':{pressure:66,coverage:91,run:69,traits:['iq','speed','technique']},
 'Aggressive Man':{pressure:79,coverage:86,run:65,traits:['speed','technique','composure']},
 'Split-Safety Control':{pressure:61,coverage:88,run:76,traits:['iq','technique','versatility']},
 'Contain & Rally':{pressure:57,coverage:78,run:86,traits:['iq','power','durability']}
};
const PRIORITIES=['Prestige','NIL','Development','Early Role','Coaching','Academics','Scheme Fit','Winning'];
const OFF_POS=new Set(['QB','RB','WR','TE','OT','OG','C','K','P']);
let schools=window.SCHOOL_DATA||[],universe=null;

const rng=(a,b)=>Math.random()*(b-a)+a, gi=(a,b)=>Math.floor(rng(a,b+1)), pick=a=>a[gi(0,a.length-1)], clamp=(v,a,b)=>Math.max(a,Math.min(b,v));
function gauss(){let u=0,v=0;while(!u)u=Math.random();while(!v)v=Math.random();return Math.sqrt(-2*Math.log(u))*Math.cos(2*Math.PI*v)}
function uid(){return (crypto&&crypto.randomUUID)?crypto.randomUUID():`${Date.now()}-${Math.random()}`}
function grade(v){return v>=94?'A+':v>=90?'A':v>=86?'A-':v>=82?'B+':v>=78?'B':v>=74?'B-':v>=70?'C+':v>=66?'C':v>=62?'C-':v>=57?'D+':v>=52?'D':'F'}
function stars(v){return v>=91?5:v>=82?4:v>=70?3:v>=58?2:1}
function heightStr(i){return `${Math.floor(i/12)}'${i%12}"`}
function avg(a){return a.length?a.reduce((x,y)=>x+y,0)/a.length:55}
async function loadSchools(){schools=window.SCHOOL_DATA||[];return schools}
function staffRating(base=65){return clamp(Math.round(base+gauss()*12),25,99)}
function generateCoach(role,school){
 const base=56+school.prestige*.16+school.resources*.08;
 return {id:uid(),name:`${pick(FIRST)} ${pick(LAST)}`,role,age:gi(31,66),recruiting:staffRating(base),development:staffRating(base),evaluation:staffRating(base),playCall:staffRating(base),adaptability:staffRating(base),loyalty:staffRating(63),ambition:staffRating(68),years:gi(0,6)};
}
function generateStaff(s){return {HC:generateCoach('Head Coach',s),OC:generateCoach('Offensive Coordinator',s),DC:generateCoach('Defensive Coordinator',s),RC:generateCoach('Recruiting Coordinator',s),SC:generateCoach('Strength & Performance',s)}}
function generatePlayer(school,pos,idx){
 const [hr,wr]=BODY[pos],h=gi(hr[0],hr[1]),w=gi(wr[0],wr[1]),classYr=pick(['FR','FR','SO','SO','JR','JR','SR']);
 const program=school.prestige*.48+school.resources*.16+school.development*.22+school.facilities*.14;
 const trueNow=clamp(Math.round(42+program*.38+gauss()*8+(classYr==='FR'?-5:classYr==='SO'?-2:classYr==='SR'?3:1)),38,97);
 const attrs={speed:clamp(Math.round(trueNow+gauss()*10+(['WR','CB','RB','S'].includes(pos)?7:['DT','OG','C'].includes(pos)?-11:0)),25,99),power:clamp(Math.round(trueNow+gauss()*9+(['DT','OG','C','OT','EDGE','TE'].includes(pos)?6:['CB','WR'].includes(pos)?-8:0)),25,99),technique:clamp(Math.round(trueNow+gauss()*9),25,99),iq:clamp(Math.round(trueNow+gauss()*11),20,99),composure:clamp(Math.round(66+gauss()*15),15,99),durability:clamp(Math.round(72+gauss()*14),20,99),versatility:clamp(Math.round(64+gauss()*17),15,99)};
 const dev=clamp(Math.round(61+school.development*.18+gauss()*16),20,99),work=clamp(Math.round(65+gauss()*18),15,99),coach=clamp(Math.round(68+gauss()*16),15,99),upside=clamp(Math.round(trueNow+(100-trueNow)*(dev/100)*rng(.28,.72)),trueNow,99);
 const scoutErr=gauss()*((100-school.development)/18+2.2),perceived=clamp(Math.round(trueNow+scoutErr),35,99),perceivedUpside=clamp(Math.round(upside+gauss()*6),perceived,99),morale=clamp(Math.round(72+gauss()*13),20,99);
 let role=idx<Math.ceil((POS_COUNTS[pos]||4)*.35)?'Starter mix':idx<Math.ceil((POS_COUNTS[pos]||4)*.7)?'Rotation':'Development';if(classYr==='FR'&&role==='Development')role='Redshirt candidate';
 return {id:uid(),name:`${pick(FIRST)} ${pick(LAST)}`,pos,year:classYr,height:h,weight:w,style:pick(STYLES[pos]),trueNow,upside,perceived,perceivedUpside,dev,work,coach,morale,role,...attrs,stats:newStats(),career:{games:0},origin:'Initial roster'};
}
function newStats(){return {games:0,passYds:0,passTD:0,int:0,rushYds:0,rushTD:0,recYds:0,recTD:0,tackles:0,sacks:0,intDef:0}}
function generateRoster(s){let r=[];Object.entries(POS_COUNTS).forEach(([p,n])=>{for(let i=0;i<n;i++)r.push(generatePlayer(s,p,i))});while(r.length<93)r.push(generatePlayer(s,pick(POS),99));return r}
function generateRecruitPool(n){
 const regions=[...new Set(schools.map(s=>s.conference))];let out=[];
 for(let i=0;i<n;i++){const pos=pick(POS.filter(x=>!['K','P'].includes(x))),raw=clamp(Math.round(57+gauss()*12.5),36,98),up=clamp(Math.round(raw+rng(4,22)+gauss()*5),raw,99);out.push({id:uid(),name:`${pick(FIRST)} ${pick(LAST)}`,pos,stars:stars(raw),style:pick(STYLES[pos]),homeRegion:pick(regions),priority:pick(PRIORITIES),trueNow:raw,upside:up,scout:clamp(Math.round(raw+gauss()*6),35,99),scoutUp:clamp(Math.round(up+gauss()*8),40,99),targeted:false,interest:gi(5,45),relationship:gi(0,20),leader:null,committed:null,visit:false,work:clamp(Math.round(65+gauss()*18),15,99),dev:clamp(Math.round(68+gauss()*18),15,99)})}
 return out.sort((a,b)=>b.stars-a.stars||b.scoutUp-a.scoutUp);
}
function initUniverse(){
 const os=Object.keys(OFF_SCHEMES),ds=Object.keys(DEF_SCHEMES);
 const teams=schools.map(s=>({...s,nickname:'',staff:generateStaff(s),offScheme:pick(os),defScheme:pick(ds),roster:generateRoster(s),w:0,l:0,cw:0,cl:0,pf:0,pa:0,sos:0,rank:null,champ:false,schedule:[],history:[],commits:[]}));
 universe={version:'0.4',year:2027,week:0,phase:'regular',teams,recruits:generateRecruitPool(900),history:[],latest:[],confChamps:[],champion:null,lastDetailedGame:null,movementLog:[],offseasonDone:false};buildSchedule();ranked();
}
