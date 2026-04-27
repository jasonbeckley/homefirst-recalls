const express = require("express");
const axios = require("axios");
const XLSX = require("xlsx");
const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const { FaExclamationTriangle, FaTools, FaFire, FaBolt, FaDollarSign, FaClock, FaCalendarWeek } = require("react-icons/fa");
const { execSync } = require("child_process");
const fs = require("fs");
const path = require("path");
const FormData = require("form-data");

const app = express();
app.use(express.json({ limit: "50mb" }));

const SLACK_TOKEN = process.env.SLACK_TOKEN;
const SLACK_CHANNEL = process.env.SLACK_CHANNEL || "C0AUBTULV8E";
const WEBHOOK_SECRET = process.env.WEBHOOK_SECRET || "homefirst2026";
const GMAIL_TOKEN = process.env.GMAIL_TOKEN;

const C = {
  navy:"0D1B3E", teal:"00A896", orange:"F97316", offWhite:"F4F6FA",
  white:"FFFFFF", slate:"475569", red:"DC2626", lightRed:"FEF2F2",
  amber:"D97706", green:"059669", muted:"94A3B8",
  plumbing:"1D4ED8", hvac:"00A896", elec:"F97316",
};

async function makeIcon(IconComponent, color, size=256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(IconComponent, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}
const shadow = () => ({ type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.10 });

function getBU(bu) {
  bu = String(bu||"");
  if (bu.includes("HVAC")) return "HVAC";
  if (bu.includes("Plumbing")) return "Plumbing";
  if (bu.includes("Electric")) return "Electrical";
  return "Other";
}
function getReason(s) {
  s = String(s||"").toLowerCase();
  if (s.includes("not work")||s.includes("not turn")||s.includes("wont turn")||s.includes("not switching")) return "Not Working";
  if (s.includes("leak")) return "Leak";
  if (s.includes("loose")||s.includes("came out")||s.includes("ducting")) return "Install Fault";
  if (s.includes("running")) return "Cont. Running";
  return "Other";
}
function ordinal(n) {
  return n + (n>=11&&n<=13?"th":{1:"st",2:"nd",3:"rd"}[n%10]||"th");
}
function formatDateRange(jobs) {
  const months = ["January","February","March","April","May","June","July","August","September","October","November","December"];
  const days = jobs.map(j => j.date).filter(Boolean).map(d => new Date(d)).filter(d => !isNaN(d)).sort((a,b)=>a-b);
  if (!days.length) return "This Week";
  const first = days[0], last = days[days.length-1];
  const dow = ["Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"];
  const year = last.getFullYear();
  const fMonth = months[first.getMonth()], lMonth = months[last.getMonth()];
  return fMonth === lMonth
    ? dow[first.getDay()]+" "+ordinal(first.getDate())+" - "+dow[last.getDay()]+" "+ordinal(last.getDate())+" "+lMonth+" "+year
    : dow[first.getDay()]+" "+ordinal(first.getDate())+" "+fMonth+" - "+dow[last.getDay()]+" "+ordinal(last.getDate())+" "+lMonth+" "+year;
}
function fmt(n) { return n>=1000 ? "$"+(n/1000).toFixed(1)+"K" : "$"+Math.round(n); }

function parseXLSX(buffer) {
  const wb = XLSX.read(buffer, { type:"buffer", cellDates:true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
  return rows
    .filter(r => r["Business Unit"] && /^\d+$/.test(String(r["Job #"]||"").trim()))
    .map(r => ({
      jobNum: String(r["Job #"]).trim(),
      bu: getBU(r["Business Unit"]),
      hours: parseFloat(r["Total Hours Worked"]) || 0,
      cost: parseFloat(r["Jobs Total Costs"]) || 0,
      summary: String(r["Summary"] || ""),
      reason: getReason(r["Summary"]),
      date: r["Completion Date"] ? new Date(r["Completion Date"]) : null,
    }));
}

async function buildSlide(jobs, outputPath) {
  const n=jobs.length, tc=jobs.reduce((s,j)=>s+j.cost,0), th=jobs.reduce((s,j)=>s+j.hours,0);
  const ac=n?tc/n:0, ah=n?th/n:0;
  const byBU={};
  jobs.forEach(j=>{if(!byBU[j.bu])byBU[j.bu]={recalls:0,cost:0,hours:0};byBU[j.bu].recalls++;byBU[j.bu].cost+=j.cost;byBU[j.bu].hours+=j.hours;});
  const rc={};
  jobs.forEach(j=>{rc[j.reason]=(rc[j.reason]||0)+1;});
  const reasons=Object.entries(rc).sort((a,b)=>b[1]-a[1]).slice(0,4).reverse();
  const top3=[...jobs].sort((a,b)=>b.hours-a.hours).slice(0,3);
  const dateRange=formatDateRange(jobs);
  const divOrder=["Plumbing","HVAC","Electrical"];
  const divColors={Plumbing:C.plumbing,HVAC:C.hvac,Electrical:C.elec};
  const divIcons={Plumbing:FaTools,HVAC:FaFire,Electrical:FaBolt};

  const pres=new pptxgen();
  pres.layout="LAYOUT_16x9";
  const s=pres.addSlide();
  s.background={color:C.offWhite};

  s.addShape(pres.shapes.RECTANGLE,{x:0,y:0,w:10,h:0.75,fill:{color:C.navy},line:{type:"none"}});
  const warnIcon=await makeIcon(FaExclamationTriangle,"#F97316");
  s.addImage({data:warnIcon,x:0.3,y:0.14,w:0.42,h:0.42});
  s.addText("RECALLS - "+dateRange.toUpperCase(),{x:0.85,y:0,w:8.0,h:0.75,fontSize:16,bold:true,color:C.white,fontFace:"Calibri",valign:"middle",margin:0});
  s.addText("HomeFirst Services",{x:7.6,y:0,w:2.1,h:0.75,fontSize:11,color:"A0AEC0",fontFace:"Calibri",align:"right",valign:"middle",margin:0});

  const kpis=[
    {label:"Total Recalls",value:String(n),sub:"this week",icon:FaTools,iconColor:"#0D1B3E",accent:C.navy},
    {label:"Total Cost",value:fmt(tc),sub:"all recalls",icon:FaDollarSign,iconColor:"#059669",accent:"059669"},
    {label:"Hours Consumed",value:th.toFixed(1),sub:"labour time",icon:FaClock,iconColor:"#D97706",accent:"D97706"},
    {label:"Avg Cost/Recall",value:fmt(ac),sub:"per job",icon:FaFire,iconColor:"#DC2626",accent:"DC2626"},
    {label:"Avg Hours/Recall",value:ah.toFixed(1)+" hrs",sub:"per job",icon:FaCalendarWeek,iconColor:"#00A896",accent:"00A896"},
  ];
  const kW=1.82,kH=1.0,kGap=0.045,kY=0.85;
  for(let i=0;i<kpis.length;i++){
    const kx=0.1+i*(kW+kGap),k=kpis[i];
    s.addShape(pres.shapes.RECTANGLE,{x:kx,y:kY,w:kW,h:kH,fill:{color:C.white},line:{type:"none"},shadow:shadow()});
    s.addShape(pres.shapes.RECTANGLE,{x:kx,y:kY,w:kW,h:0.055,fill:{color:k.accent},line:{type:"none"}});
    const ico=await makeIcon(k.icon,k.iconColor);
    s.addImage({data:ico,x:kx+0.12,y:kY+0.12,w:0.3,h:0.3});
    s.addText(k.value,{x:kx+0.05,y:kY+0.13,w:kW-0.1,h:0.38,fontSize:17,bold:true,color:C.navy,fontFace:"Calibri",align:"center",valign:"middle",margin:0});
    s.addText(k.label,{x:kx+0.05,y:kY+0.53,w:kW-0.1,h:0.25,fontSize:9.5,bold:true,color:C.slate,fontFace:"Calibri",align:"center",margin:0});
    s.addText(k.sub,{x:kx+0.05,y:kY+0.76,w:kW-0.1,h:0.2,fontSize:8,color:C.muted,fontFace:"Calibri",align:"center",margin:0});
  }

  const dW=3.12,dH=1.12,dGap=0.07,dY=2.0;
  for(let i=0;i<divOrder.length;i++){
    const name=divOrder[i],d=byBU[name]||{recalls:0,cost:0,hours:0},col=divColors[name],dx=0.1+i*(dW+dGap);
    s.addShape(pres.shapes.RECTANGLE,{x:dx,y:dY,w:dW,h:dH,fill:{color:C.white},line:{type:"none"},shadow:shadow()});
    s.addShape(pres.shapes.RECTANGLE,{x:dx,y:dY,w:0.06,h:dH,fill:{color:col},line:{type:"none"}});
    const ico=await makeIcon(divIcons[name],"#"+col);
    s.addImage({data:ico,x:dx+0.15,y:dY+0.08,w:0.38,h:0.38});
    s.addText(name,{x:dx+0.62,y:dY+0.06,w:dW-0.7,h:0.3,fontSize:13,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    const stats=[{label:"Recalls",val:String(d.recalls)},{label:"Total Cost",val:fmt(d.cost)},{label:"Total Hours",val:d.hours.toFixed(1)+"h"}];
    const sW=(dW-0.7)/3;
    stats.forEach((st,j)=>{const sx=dx+0.62+j*sW;s.addText(st.val,{x:sx,y:dY+0.42,w:sW,h:0.3,fontSize:13,bold:true,color:"#"+col,fontFace:"Calibri",align:"center",margin:0});s.addText(st.label,{x:sx,y:dY+0.72,w:sW,h:0.22,fontSize:8,color:C.muted,fontFace:"Calibri",align:"center",margin:0});});
  }

  const rX=0.1,rY=3.22,rW=4.7,rH=1.72;
  s.addShape(pres.shapes.RECTANGLE,{x:rX,y:rY,w:rW,h:rH,fill:{color:C.white},line:{type:"none"},shadow:shadow()});
  s.addText("Top Recall Reasons",{x:rX+0.15,y:rY+0.1,w:rW-0.3,h:0.28,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
  s.addChart(pres.charts.BAR,[{name:"Recalls",labels:reasons.map(r=>r[0]),values:reasons.map(r=>r[1])}],{
    x:rX+0.05,y:rY+0.35,w:rW-0.1,h:rH-0.42,barDir:"bar",
    chartColors:["0D1B3E","DC2626","94A3B8","F97316"],
    chartArea:{fill:{color:"FFFFFF"}},catAxisLabelColor:"475569",valAxisLabelColor:"94A3B8",
    valGridLine:{color:"E2E8F0",size:0.5},catGridLine:{style:"none"},
    showValue:true,dataLabelColor:"FFFFFF",dataLabelPosition:"inEnd",showLegend:false,dataLabelFontSize:10,
  });

  const tX=5.0,tY=3.22,tW=4.9,tH=1.72;
  s.addShape(pres.shapes.RECTANGLE,{x:tX,y:tY,w:tW,h:tH,fill:{color:C.white},line:{type:"none"},shadow:shadow()});
  s.addText("Top 3 Jobs by Hours",{x:tX+0.15,y:tY+0.1,w:tW-0.3,h:0.28,fontSize:11,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
  const rowH=0.41,rowY0=tY+0.42;
  top3.forEach((j,i)=>{
    const ry=rowY0+i*(rowH+0.04),col=divColors[j.bu]||C.slate;
    const desc=j.summary.replace(/Possible recall:\s*/i,"").substring(0,55);
    s.addShape(pres.shapes.RECTANGLE,{x:tX+0.12,y:ry,w:tW-0.24,h:rowH,fill:{color:i%2===0?"F8FAFC":"FFFFFF"},line:{type:"none"}});
    s.addShape(pres.shapes.RECTANGLE,{x:tX+0.12,y:ry,w:0.05,h:rowH,fill:{color:col},line:{type:"none"}});
    s.addShape(pres.shapes.OVAL,{x:tX+0.22,y:ry+0.09,w:0.24,h:0.24,fill:{color:col},line:{type:"none"}});
    s.addText(String(i+1),{x:tX+0.22,y:ry+0.09,w:0.24,h:0.24,fontSize:10,bold:true,color:C.white,fontFace:"Calibri",align:"center",valign:"middle",margin:0});
    s.addText("#"+j.jobNum,{x:tX+0.52,y:ry+0.03,w:0.72,h:0.2,fontSize:9.5,bold:true,color:C.navy,fontFace:"Calibri",margin:0});
    s.addText(j.bu,{x:tX+0.52,y:ry+0.22,w:0.72,h:0.16,fontSize:8,color:C.muted,fontFace:"Calibri",margin:0});
    s.addText(desc,{x:tX+1.28,y:ry+0.04,w:2.22,h:0.34,fontSize:8,color:C.slate,fontFace:"Calibri",valign:"middle",margin:0});
    s.addText(j.hours.toFixed(1)+" hrs",{x:tX+3.52,y:ry+0.03,w:1.2,h:0.2,fontSize:10,bold:true,color:"#"+col,fontFace:"Calibri",align:"right",margin:0});
    s.addText(fmt(j.cost),{x:tX+3.52,y:ry+0.22,w:1.2,h:0.16,fontSize:8,color:C.muted,fontFace:"Calibri",align:"right",margin:0});
  });

  const top1=top3[0];
  const flagText=top1&&th>0&&top1.hours/th>0.3
    ? "Job #"+top1.jobNum+" ("+top1.bu+", "+top1.hours.toFixed(1)+" hrs, "+fmt(top1.cost)+") = "+Math.round(top1.hours/th*100)+"% of all recall hours this week - review with Scott."
    : n+" recalls completed this week across Plumbing, HVAC and Electrical.";
  s.addShape(pres.shapes.RECTANGLE,{x:0.1,y:5.08,w:9.8,h:0.4,fill:{color:C.lightRed},line:{pt:1,color:"DC2626"},shadow:shadow()});
  const warnIco=await makeIcon(FaExclamationTriangle,"#DC2626");
  s.addImage({data:warnIco,x:0.22,y:5.16,w:0.24,h:0.24});
  s.addText(flagText,{x:0.55,y:5.08,w:9.2,h:0.4,fontSize:9.5,color:"991B1B",fontFace:"Calibri",valign:"middle",margin:0});

  await pres.writeFile({fileName:outputPath});
}

async function postPDFToSlack(pdfPath, dateRange) {
  const form=new FormData();
  form.append("channels",SLACK_CHANNEL);
  form.append("initial_comment","PDF Recalls Report - "+dateRange+"\nAuto-generated from ServiceTitan.");
  form.append("filename",path.basename(pdfPath));
  form.append("file",fs.createReadStream(pdfPath));
  const resp=await axios.post("https://slack.com/api/files.upload",form,{
    headers:{...form.getHeaders(),Authorization:"Bearer "+SLACK_TOKEN},
  });
  if(!resp.data.ok) throw new Error("Slack error: "+resp.data.error);
  return resp.data;
}

app.post("/webhook", async (req, res) => {
  const {secret,attachment_url,attachment_base64,message_id,gmail_token}=req.body;
  if(secret!==WEBHOOK_SECRET) return res.status(401).json({error:"Unauthorized"});
  console.log("Webhook received");
  res.json({status:"processing"});
  try {
    let xlsxBuffer;
    const token=gmail_token||GMAIL_TOKEN;
    if(attachment_base64){
      xlsxBuffer=Buffer.from(attachment_base64,"base64");
    } else if(attachment_url){
      const resp=await axios.get(attachment_url,{responseType:"arraybuffer",headers:token?{Authorization:"Bearer "+token}:{}});
      xlsxBuffer=Buffer.from(resp.data);
    } else if(message_id&&token){
      const msgResp=await axios.get("https://gmail.googleapis.com/gmail/v1/users/me/messages/"+message_id+"?format=full",{headers:{Authorization:"Bearer "+token}});
      const findAtt=(parts=[])=>{for(const p of parts){if(p.filename?.endsWith(".xlsx")&&p.body?.attachmentId)return p.body.attachmentId;const f=findAtt(p.parts);if(f)return f;}return null;};
      const attId=findAtt(msgResp.data.payload?.parts||[msgResp.data.payload]);
      if(!attId) throw new Error("No xlsx attachment found");
      const attResp=await axios.get("https://gmail.googleapis.com/gmail/v1/users/me/messages/"+message_id+"/attachments/"+attId,{headers:{Authorization:"Bearer "+token}});
      xlsxBuffer=Buffer.from(attResp.data.data.replace(/-/g,"+").replace(/_/g,"/"),"base64");
    } else {
      throw new Error("No attachment_url, attachment_base64, or message_id provided");
    }
    const jobs=parseXLSX(xlsxBuffer);
    console.log("Parsed "+jobs.length+" jobs");
    const dateRange=formatDateRange(jobs);
    const pptxPath="/tmp/recalls_"+Date.now()+".pptx";
    const pdfPath=pptxPath.replace(".pptx",".pdf");
    await buildSlide(jobs,pptxPath);
    execSync('soffice --headless --convert-to pdf --outdir /tmp "'+pptxPath+'"',{timeout:60000});
    await postPDFToSlack(pdfPath,dateRange);
    console.log("Posted to Slack");
    fs.unlinkSync(pptxPath);
    fs.unlinkSync(pdfPath);
  } catch(err){
    console.error("Error:",err.message);
    await axios.post("https://slack.com/api/chat.postMessage",{channel:SLACK_CHANNEL,text:"Recalls automation failed: "+err.message},{headers:{Authorization:"Bearer "+SLACK_TOKEN}}).catch(()=>{});
  }
});

app.get("/",(req,res)=>res.json({status:"ok",service:"HomeFirst Recalls Automation"}));

const PORT=process.env.PORT||3000;
app.listen(PORT,()=>console.log("Server running on port "+PORT));
