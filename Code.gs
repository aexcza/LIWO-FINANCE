const CONFIG = {
  SPREADSHEET_ID: "19G1z2Yl56XwQpQxPNLzV4JUhxWy7f_hXohS8wDyeIZ8",
  PROJECT_NAME: "LIWO FINANCE TRACKER",
  SESSION_HOURS: 12,

  BOOTSTRAP_ADMIN_USERNAME: "admin",
  BOOTSTRAP_ADMIN_NAME: "LIWO Administrator",
  BOOTSTRAP_ADMIN_PASSWORD: "Althea0610",

  DEFAULT_INVITE_CODE: "liwoecfinance",

  RECEIPT_FOLDER_NAME: "LIWO Finance Receipts",
  MAX_RECEIPT_BYTES: 5 * 1024 * 1024,
  REPORT_FOLDER_NAME: "LIWO Finance Reports",
  REPORT_RECIPIENTS: [
    "valezachristian0821@gmail.com",
    "tristanvelasco4@gmail.com"
  ],
  REPORT_WEEKLY_DAY: "SUNDAY",
  REPORT_WEEKLY_HOUR: 17
};
const SHEETS={
 settings:["Key","Value"],
 users:["Username","Name","PasswordHash","Role","Active","CreatedAt","UpdatedAt"],
 clients:["ClientID","Client / Project Name","Reference","Contract Budget","Active","CreatedAt","UpdatedAt"],
 payments:["Timestamp","Date","ClientID","Payment Ref.","Description / Milestone","Due Amount","Amount Paid","Payment Method","Notes","Entered By","Username"],
 expenses:["Timestamp","Date","ClientID","Type","Category","Payee / Supplier","Description","Amount","Payment Method","Receipt / Ref.","Approved By","Notes","Entered By","Username","Receipt File URL","Receipt File ID"],
 budget:["Category","Budget"],tools:["Timestamp","Tool","Tool ID","Borrowed By","Project","Date Borrowed","Expected Return","Date Returned","Status","Notes","Entered By","Username","ClientID"],cash_balances:["Account","Balance","UpdatedAt","UpdatedBy","Notes"],audit:["Timestamp","Action","Username","Name","Details"],
  milestones:["MilestoneID","ClientID","Milestone","DueDate","Amount","Status","PaidAmount","Notes","CreatedAt","UpdatedAt","CreatedBy"],
  receipt_verification:["ReceiptID","Status","VerifiedBy","VerifiedAt","Notes"],
  notification_reads:["Username","LastReadAt"],
  cash_reconciliation:["Timestamp","BankExpected","CashOnHandExpected","ActualBank","ActualCashOnHand","Variance","Notes","ReconciledBy"],
  project_budgets:["ClientID","Category","Budget","UpdatedAt","UpdatedBy"],
  report_runs:["Timestamp","Period","ReportType","FileId","FileUrl","CreatedBy"],
  calculator_history:["Timestamp","Calculator","Inputs","Result","ClientID","CreatedBy","Username"],
  bills:["Timestamp","BillID","ClientID","Bill / Payee","Category","DueDate","Amount","Status","Notes","Entered By","Username","UpdatedAt"],

};
function setup(){
  ensureSheet_("weekly_report_config",[
    "Key","Value","UpdatedAt","UpdatedBy"
  ]);
  ensureSheet_("report_recipients",[
    "Email","Enabled","UpdatedAt","UpdatedBy"
  ]);
 const s=ss();Object.keys(SHEETS).forEach(n=>{let sh=s.getSheetByName(n)||s.insertSheet(n);ensureHeaders_(sh,SHEETS[n]);sh.setFrozenRows(1);sh.getRange(1,1,1,SHEETS[n].length).setFontWeight("bold")});
 let sm=settingsMap(),st=s.getSheetByName("settings");
 if(sm.ProjectName===undefined)st.appendRow(["ProjectName",CONFIG.PROJECT_NAME]);
 if(sm.InviteCodeHash===undefined){if(CONFIG.DEFAULT_INVITE_CODE==="CHANGE_THIS_INVITE_CODE")throw Error("Change DEFAULT_INVITE_CODE before setup().");st.appendRow(["InviteCodeHash",hash_(CONFIG.DEFAULT_INVITE_CODE)])}
 if(sm.RegistrationOpen===undefined)st.appendRow(["RegistrationOpen",true]);
 if(sm.ReportRecipients===undefined)st.appendRow(["ReportRecipients",""]);
 // Record Status migration: fund transactions no longer require approval.
 let ex=s.getSheetByName("expenses");if(ex){if(ex.getLastColumn()<17)ex.getRange(1,17).setValue("Record Status");else ex.getRange(1,17).setValue("Record Status");if(ex.getLastRow()>1)ex.getRange(2,17,ex.getLastRow()-1,1).setValue("Recorded");}
 let u=s.getSheetByName("users");if(u.getLastRow()===1){if(CONFIG.BOOTSTRAP_ADMIN_PASSWORD==="CHANGE_THIS_BEFORE_SETUP")throw Error("Change BOOTSTRAP_ADMIN_PASSWORD before setup().");u.appendRow([CONFIG.BOOTSTRAP_ADMIN_USERNAME,CONFIG.BOOTSTRAP_ADMIN_NAME,hash_(CONFIG.BOOTSTRAP_ADMIN_PASSWORD),"Admin",true,new Date(),new Date()])}
 return"Setup complete";
}
function doGet(){return json({ok:true,service:"LIWO Finance Tracker"})}

/* =========================
   LIWO AUTOMATED REPORTING
   ========================= */


/* Reporting compatibility helpers.
   These normalize the newer reporting code to this Code.gs file's
   existing spreadsheet/audit helpers. */
function readSheet_(sheetName){
  const sh=ss().getSheetByName(sheetName);
  if(!sh || sh.getLastRow()<2)return [];
  const values=sh.getDataRange().getValues();
  const headers=values[0].map(function(h){return String(h||"").trim();});
  return values.slice(1).map(function(row){
    const o={};
    headers.forEach(function(h,i){if(h)o[h]=row[i];});
    return o;
  });
}

function audit_(u,action,entity,details,extra){
  const actor=u||{username:"SYSTEM",name:"SYSTEM"};
  let detail=String(details||"");
  if(extra)detail += (detail?" | ":"")+String(extra);
  const username=String(actor.username||"SYSTEM");
  const name=String(actor.name||"SYSTEM");
  audit(action,{username,name},String(entity||detail||""));
}

function getOrCreateReceiptFolder_(){
  return receiptRootFolder_();
}

function getReportRecipients_(){
  // Admin-controlled recipients stored in report_recipients.
  // Fall back to the configured defaults until the Admin saves recipients.
  const sh=ss().getSheetByName("report_recipients");
  if(sh && sh.getLastRow()>1){
    const out=sh.getDataRange().getValues().slice(1).filter(function(r){
      const email=String(r[0]||"").trim();
      const enabled=String(r[1]).toLowerCase()!=="false";
      return enabled && email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
    }).map(function(r){return String(r[0]).trim().toLowerCase();});
    if(out.length) return Array.from(new Set(out));
  }
  return (CONFIG.REPORT_RECIPIENTS || []).slice();
}

function setReportRecipients_(u, emails){
  adminOnly(u);
  const clean = [];
  const seen = {};
  (emails||[]).forEach(function(x){
    const email=String(x||"").trim().toLowerCase();
    if(email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) && !seen[email]){
      seen[email]=true; clean.push(email);
    }
  });
  const sh=ensureSheet_("report_recipients",["Email","Enabled","UpdatedAt","UpdatedBy"]);
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,4).clearContent();
  if(clean.length) sh.getRange(2,1,clean.length,4).setValues(clean.map(e=>[e,true,nowISO(),u.username]));
  audit_(u,"REPORT_RECIPIENTS_UPDATED","SYSTEM","Weekly/monthly report recipients updated",JSON.stringify(clean));
  return {ok:true,recipients:clean};
}

function getReportConfig_(u){
  adminOnly(u);
  const recipients=getReportRecipients_();
  const sh=ss().getSheetByName("weekly_report_config");
  const cfg={weeklyEnabled:false,weeklyDay:"SUNDAY",weeklyHour:17,recipients:recipients};
  if(sh && sh.getLastRow()>1){
    sh.getDataRange().getValues().slice(1).forEach(function(r){
      const key=String(r[0]||"").trim();
      const val=r[1];
      if(key==="weeklyEnabled") cfg.weeklyEnabled=String(val).toLowerCase()==="true";
      if(key==="weeklyDay") cfg.weeklyDay=String(val||"SUNDAY").toUpperCase();
      if(key==="weeklyHour") cfg.weeklyHour=Math.max(0,Math.min(23,Number(val)||17));
    });
  }
  return {ok:true,config:cfg};
}

function saveReportConfig_(u, cfg){
  adminOnly(u);
  const sh=ensureSheet_("weekly_report_config",["Key","Value","UpdatedAt","UpdatedBy"]);
  const values=[
    ["weeklyEnabled",!!cfg.weeklyEnabled,nowISO(),u.username],
    ["weeklyDay",String(cfg.weeklyDay||"SUNDAY").toUpperCase(),nowISO(),u.username],
    ["weeklyHour",Math.max(0,Math.min(23,Number(cfg.weeklyHour)||17)),nowISO(),u.username]
  ];
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,4).clearContent();
  sh.getRange(2,1,values.length,4).setValues(values);
  audit_(u,"REPORT_CONFIG_UPDATED","SYSTEM","Weekly automated report configuration updated",JSON.stringify(cfg));
  return {ok:true,config:getReportConfig_(u).config};
}

function getProjectFinancialSummary_(startDate,endDate){
  const clients=readSheet_("clients");
  const payments=readSheet_("payments");
  const expenses=readSheet_("expenses");
  const tz=Session.getScriptTimeZone()||"Asia/Manila";
  const hasRange=!!(startDate||endDate);
  const startKey=startDate?Utilities.formatDate(new Date(startDate),tz,"yyyy-MM-dd"):null;
  const endKey=endDate?Utilities.formatDate(new Date(endDate),tz,"yyyy-MM-dd"):null;

  const inRange=function(value){
    if(!hasRange)return true;
    if(value===""||value===null||value===undefined)return false;
    const d=new Date(value);
    if(isNaN(d.getTime()))return false;
    const key=Utilities.formatDate(d,tz,"yyyy-MM-dd");
    return (!startKey||key>=startKey) && (!endKey||key<=endKey);
  };

  const projectMap={};
  clients.forEach(function(c){
    const id=String(c.ClientID||"").trim();
    if(!id)return;
    projectMap[id]={
      clientId:id,
      project:String(c.ProjectName||c.ClientName||"").trim(),
      reference:String(c.Reference||"").trim(),
      contract:Number(c.ContractValue||c.Budget||0)||0,
      payments:0,
      expenses:0,
      otherIncome:0,
      refunds:0,
      netMovement:0,
      uncollected:0,
      profit:0,
      margin:0,
      periodTransactionCount:0,
      paymentCount:0,
      expenseCount:0,
      categories:{}
    };
  });

  // Lifetime payments are retained only for the true contract balance.
  const lifetimePayments={};
  payments.forEach(function(p){
    const id=String(p.ClientID||"").trim();
    if(!projectMap[id])return;
    const amount=Number(p.Amount||p.AmountPaid||0)||0;
    lifetimePayments[id]=(lifetimePayments[id]||0)+amount;
    if(inRange(p.Date||p.Timestamp)){
      projectMap[id].payments+=amount;
      projectMap[id].paymentCount++;
      projectMap[id].periodTransactionCount++;
    }
  });

  expenses.forEach(function(e){
    const id=String(e.ClientID||"").trim();
    if(!projectMap[id] || !inRange(e.Date||e.Timestamp))return;
    const amount=Number(e.Amount||0)||0;
    const type=String(e.Type||"Expense").trim().toLowerCase();
    const category=String(e.Category||"Uncategorized").trim()||"Uncategorized";
    projectMap[id].periodTransactionCount++;

    if(type==="other income"){
      projectMap[id].otherIncome+=amount;
      return;
    }
    if(type==="refund"){
      projectMap[id].refunds+=amount;
      return;
    }

    projectMap[id].expenses+=amount;
    projectMap[id].expenseCount++;
    projectMap[id].categories[category]=(projectMap[id].categories[category]||0)+amount;
  });

  return Object.keys(projectMap).map(function(id){
    const x=projectMap[id];
    const collectedLifetime=lifetimePayments[id]||0;
    x.uncollected=Math.max(0,x.contract-collectedLifetime);
    x.profit=x.payments+x.otherIncome-x.refunds-x.expenses;
    x.netMovement=x.payments+x.otherIncome-x.refunds-x.expenses;
    x.margin=x.payments>0?(x.profit/x.payments)*100:0;
    x.collectionPct=x.contract>0?(collectedLifetime/x.contract)*100:0;
    x.spentPctOfReceived=x.payments>0?(x.expenses/x.payments)*100:0;
    x.remainingPctOfReceived=x.payments>0?((x.payments-x.expenses)/x.payments)*100:0;
    x.budgetUsed=x.contract>0?(x.expenses/x.contract)*100:0;

    const categoryEntries=Object.keys(x.categories).map(function(k){
      return {name:k,amount:x.categories[k]};
    }).sort(function(a,b){return b.amount-a.amount;});
    x.topExpenseCategory=categoryEntries[0]||{name:"No expense category recorded",amount:0};
    x.secondExpenseCategory=categoryEntries[1]||null;
    x.categoryBreakdown=categoryEntries;
    x.health=x.budgetUsed>=100 || x.profit<0 ? "Critical" :
      (x.budgetUsed>=80 || x.collectionPct<50 ? "Attention" : "Healthy");
    return x;
  });
}
function buildFinancialReportHtml_(periodLabel, rows, reportType){
  reportType=String(reportType||"Weekly");
  let totalContract=0,totalPayments=0,totalExpenses=0,totalOtherIncome=0,totalRefunds=0,totalProfit=0,totalUncollected=0;
  rows.forEach(function(r){
    totalContract+=r.contract; totalPayments+=r.payments; totalExpenses+=r.expenses;
    totalOtherIncome+=r.otherIncome||0; totalRefunds+=r.refunds||0;
    totalProfit+=r.profit; totalUncollected+=r.uncollected;
  });

  const netMovement=totalPayments+totalOtherIncome-totalRefunds-totalExpenses;
  const collectionPct=totalContract>0?(totalPayments/totalContract)*100:0;
  const spentPct=totalPayments>0?(totalExpenses/totalPayments)*100:0;
  const remainingReceived=totalPayments-totalExpenses;
  const remainingPct=totalPayments>0?(remainingReceived/totalPayments)*100:0;
  const margin=totalPayments>0?(totalProfit/totalPayments)*100:0;

  const money=function(n){
    return "₱"+Number(n||0).toLocaleString("en-PH",{minimumFractionDigits:2,maximumFractionDigits:2});
  };
  const pct=function(n){return Number(n||0).toFixed(1)+"%";};
  const esc=function(s){
    return String(s??"").replace(/[&<>"']/g,function(m){
      return({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"})[m];
    });
  };

  const projectSections=rows.map(function(r){
    const top=r.topExpenseCategory||{name:"No expense category recorded",amount:0};
    const second=r.secondExpenseCategory;
    const topPct=r.expenses>0?(top.amount/r.expenses)*100:0;
    const secondPct=second&&r.expenses>0?(second.amount/r.expenses)*100:0;
    const collectedLifetime=r.contract-r.uncollected;
    let description="During "+periodLabel+", "+esc(r.project||"this project")+" recorded "+money(r.payments)+" in client payments";
    if(r.otherIncome) description+=", "+money(r.otherIncome)+" in other income";
    if(r.refunds) description+=", and "+money(r.refunds)+" in refunds";
    description+=". Total expenses for the period were "+money(r.expenses)+", resulting in a net cash movement of "+money(r.netMovement)+". ";
    description+="Client payments represent "+pct(r.collectionPct)+" of the contract value";
    if(r.contract) description+=", with "+money(r.uncollected)+" remaining uncollected";
    description+=". ";
    if(r.payments) description+=money(r.expenses)+" or "+pct(r.spentPctOfReceived)+" of the payments received during the period was spent, leaving "+money(remainingForProject(r))+" or "+pct(r.remainingPctOfReceived)+" of period payments before other income/refund adjustments. ";
    else description+="No client payment was recorded during this period, so spending as a percentage of period payments is not applicable. ";
    if(top.amount>0){
      description+=esc(top.name)+" was the largest expense category at "+money(top.amount)+" ("+pct(topPct)+" of project expenses).";
      if(second&&second.amount>0) description+=" "+esc(second.name)+" followed at "+money(second.amount)+" ("+pct(secondPct)+").";
    }else{
      description+="No categorized expense movement was recorded during this period.";
    }

    const cats=r.categoryBreakdown||[];
    const catRows=cats.map(function(c){
      return "<tr><td>"+esc(c.name)+"</td><td>"+money(c.amount)+"</td><td>"+pct(r.expenses>0?(c.amount/r.expenses)*100:0)+"</td></tr>";
    }).join("")||'<tr><td colspan="3" class="empty-row">No expense categories recorded.</td></tr>';

    return '<section class="project-card">'+
      '<div class="project-heading"><div><div class="eyebrow">PROJECT FINANCIAL REVIEW</div><h2>'+esc(r.project||"Unnamed Project")+'</h2>'+(r.reference?'<div class="reference">Reference · '+esc(r.reference)+'</div>':'')+'</div>'+
      '<div class="project-status">'+(r.expenses>r.payments?"ATTENTION":"IN GOOD STANDING")+'</div></div>'+
      '<div class="mini-grid">'+
      '<div class="mini-card"><small>Contract Value</small><b>'+money(r.contract)+'</b></div>'+
      '<div class="mini-card"><small>Period Payments</small><b>'+money(r.payments)+'</b></div>'+
      '<div class="mini-card"><small>Period Expenses</small><b>'+money(r.expenses)+'</b></div>'+
      '<div class="mini-card"><small>Net Movement</small><b>'+money(r.netMovement)+'</b></div>'+
      '<div class="mini-card"><small>Contract Collected</small><b>'+pct(r.collectionPct)+'</b></div>'+
      '<div class="mini-card"><small>Funds Spent</small><b>'+pct(r.spentPctOfReceived)+'</b></div>'+
      '<div class="mini-card"><small>Funds Remaining</small><b>'+pct(r.remainingPctOfReceived)+'</b></div>'+
      '<div class="mini-card"><small>Outstanding Contract</small><b>'+money(r.uncollected)+'</b></div>'+
      '</div>'+
      '<div class="section-title">Expense Category Movement</div>'+
      '<table><thead><tr><th>Category</th><th>Amount</th><th>% of Project Expenses</th></tr></thead><tbody>'+catRows+'</tbody></table>'+
      '<div class="overview"><div class="overview-label">FINANCIAL COMMENTARY</div><p>'+description+'</p></div>'+
      '</section>';
  }).join("");

  function remainingForProject(r){
    return r.payments-r.expenses;
  }

  const body=rows.map(function(r){
    return "<tr><td><strong>"+esc(r.project||"—")+"</strong></td><td>"+esc(r.reference||"—")+"</td>"+
      "<td>"+money(r.contract)+"</td><td>"+money(r.payments)+"</td><td>"+money(r.expenses)+"</td>"+
      "<td>"+money(r.otherIncome||0)+"</td><td>"+money(r.refunds||0)+"</td><td>"+money(r.netMovement)+"</td>"+
      "<td>"+pct(r.collectionPct)+"</td><td>"+pct(r.spentPctOfReceived)+"</td></tr>";
  }).join("")||'<tr><td colspan="10" class="empty-row">No project records were available.</td></tr>';

  const title=reportType+" Financial Report";
  const generated=Utilities.formatDate(new Date(),Session.getScriptTimeZone()||"Asia/Manila","MMMM d, yyyy · h:mm a");
  const css=
    '@page{size:A4;margin:12mm 11mm 14mm}'+
    'body{margin:0;background:#eef2f5;color:#18344a;font-family:Arial,Helvetica,sans-serif;font-size:10px;line-height:1.45}'+
    '.page{max-width:900px;margin:0 auto;background:#fff;padding:34px 38px;box-sizing:border-box}'+
    '.topbar{display:flex;justify-content:space-between;align-items:flex-start;padding-bottom:22px;border-bottom:3px solid #123451;margin-bottom:22px}'+
    '.brand{letter-spacing:2px;font-size:10px;font-weight:800;color:#123451}.brand-name{font-size:20px;letter-spacing:.3px;margin-top:6px;font-weight:800;color:#123451}.brand-sub{font-size:10px;color:#718395;margin-top:4px}'+
    '.report-meta{text-align:right}.report-meta .type{display:inline-block;background:#123451;color:#fff;padding:6px 9px;border-radius:3px;font-size:8px;font-weight:800;letter-spacing:1px}.report-meta .date{display:block;margin-top:8px;color:#718395;font-size:9px}'+
    'h1{font-size:28px;line-height:1.15;margin:0;color:#123451}.sub{font-size:11px;color:#66798a;margin-top:7px}.intro{margin:20px 0 18px;padding:13px 15px;background:#f6f8fa;border-left:4px solid #c88722;color:#52697b;font-size:10px}'+
    '.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin:0 0 24px}.kpi{border:1px solid #dce4ea;border-radius:5px;padding:12px 12px;background:#fff;min-height:52px;box-sizing:border-box}.kpi span{display:block;text-transform:uppercase;letter-spacing:.5px;font-size:7.5px;color:#718395;font-weight:700}.kpi b{display:block;color:#123451;font-size:13px;margin-top:7px}.kpi.accent{background:#123451;border-color:#123451}.kpi.accent span,.kpi.accent b{color:#fff}'+
    '.section-head{display:flex;justify-content:space-between;align-items:end;border-bottom:1px solid #dce4ea;padding-bottom:8px;margin:18px 0 9px}.section-head h2{margin:0;font-size:14px;color:#123451}.section-head span{font-size:8px;color:#718395}'+
    'table{width:100%;border-collapse:collapse;font-size:8.2px;margin:0 0 17px;page-break-inside:auto}thead{display:table-header-group}tr{page-break-inside:avoid}th{background:#123451;color:#fff;text-transform:uppercase;letter-spacing:.35px;font-size:7.2px;padding:8px 6px;text-align:left;border:1px solid #123451}td{padding:7px 6px;border:1px solid #dce4ea;vertical-align:top}tbody tr:nth-child(even){background:#f8fafb}.empty-row{color:#718395;text-align:center;font-style:italic}'+
    '.project-card{page-break-inside:avoid;border:1px solid #dce4ea;border-radius:6px;margin:18px 0;padding:17px;background:#fff}.project-heading{display:flex;justify-content:space-between;gap:18px;align-items:flex-start;border-bottom:1px solid #e2e8ed;padding-bottom:12px}.eyebrow{font-size:7px;letter-spacing:1.2px;font-weight:800;color:#c88722}.project-card h2{font-size:16px;color:#123451;margin:4px 0 2px}.reference{font-size:8px;color:#718395}.project-status{font-size:7px;letter-spacing:.6px;font-weight:800;padding:5px 7px;border:1px solid #d6e0e7;border-radius:3px;color:#52697b;white-space:nowrap}.mini-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:6px;margin:13px 0 16px}.mini-card{background:#f7f9fa;border:1px solid #e2e8ed;border-radius:4px;padding:8px}.mini-card small{display:block;font-size:7px;color:#718395;text-transform:uppercase;letter-spacing:.25px}.mini-card b{display:block;font-size:10px;color:#17364f;margin-top:4px}.section-title{font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.7px;color:#123451;margin:5px 0 7px}.overview{background:#f7f9fa;border:1px solid #e0e7ec;border-radius:5px;padding:12px;margin-top:10px}.overview-label{font-size:7px;letter-spacing:1px;font-weight:800;color:#c88722;margin-bottom:5px}.overview p{margin:0;color:#52697b;font-size:9px;line-height:1.55}'+
    '.notes{font-size:8px;color:#718395;margin-top:17px;padding-top:10px;border-top:1px solid #dce4ea}.signatures{page-break-inside:avoid;margin-top:28px;padding-top:17px;border-top:2px solid #123451}.signatures h3{font-size:9px;text-transform:uppercase;letter-spacing:.8px;color:#123451;margin:0 0 20px}.signature-grid{display:grid;grid-template-columns:repeat(2,1fr);gap:30px}.signature{min-height:70px;display:flex;flex-direction:column;justify-content:flex-end}.signature-line{width:210px;border-top:1px solid #273b4a;margin-bottom:7px}.signature strong{font-size:8.5px;color:#17364f}.signature span{font-size:7.5px;color:#718395;margin-top:2px}.footer{display:flex;justify-content:space-between;gap:20px;margin-top:25px;padding-top:10px;border-top:1px solid #dce4ea;font-size:7.5px;color:#81909c}.confidential{font-weight:800;letter-spacing:.4px;color:#52697b}'+
    '@media print{body{background:#fff}.page{max-width:none;margin:0;padding:0}.project-card{break-inside:avoid}.kpi,.mini-card{break-inside:avoid}}';

  return '<!doctype html><html><head><meta charset="utf-8"><title>'+esc(title)+'</title><style>'+css+'</style></head><body>'+
    '<main class="page">'+
    '<header class="topbar"><div><div class="brand">LIWO</div><div class="brand-name">LIWO ENGINEERING CONSULTANCY</div><div class="brand-sub">Finance & Project Management</div></div>'+
    '<div class="report-meta"><span class="type">'+esc(reportType.toUpperCase())+' REPORT</span><span class="date">Generated '+esc(generated)+'</span></div></header>'+
    '<h1>'+esc(title)+'</h1><div class="sub">'+esc(periodLabel)+'</div>'+
    '<div class="intro">This report provides a management-level view of project financial activity for the stated reporting period. The attached document is intended for internal financial review, project monitoring, and record-keeping.</div>'+
    '<div class="kpis">'+
    '<div class="kpi"><span>Contract Value</span><b>'+money(totalContract)+'</b></div>'+
    '<div class="kpi"><span>Period Money In</span><b>'+money(totalPayments+totalOtherIncome)+'</b></div>'+
    '<div class="kpi"><span>Period Money Out</span><b>'+money(totalExpenses+totalRefunds)+'</b></div>'+
    '<div class="kpi accent"><span>Net Movement</span><b>'+money(netMovement)+'</b></div>'+
    '<div class="kpi"><span>Collected</span><b>'+pct(collectionPct)+'</b></div>'+
    '<div class="kpi"><span>Spent of Payments</span><b>'+pct(spentPct)+'</b></div>'+
    '<div class="kpi"><span>Remaining of Payments</span><b>'+pct(remainingPct)+'</b></div>'+
    '<div class="kpi"><span>Outstanding</span><b>'+money(totalUncollected)+'</b></div>'+
    '</div>'+
    '<div class="section-head"><h2>Portfolio Financial Summary</h2><span>'+rows.length+' project(s) included</span></div>'+
    '<table><thead><tr><th>Project</th><th>Reference</th><th>Contract</th><th>Payments</th><th>Expenses</th><th>Other Income</th><th>Refunds</th><th>Net Movement</th><th>Collected %</th><th>Spent %</th></tr></thead><tbody>'+body+'</tbody></table>'+
    '<div class="section-head"><h2>Detailed Financial Review by Project</h2><span>'+esc(reportType)+' reporting period</span></div>'+
    projectSections+
    '<div class="notes">Reporting figures are restricted to the stated reporting period. Contract collection and outstanding balance use cumulative client payments so the project balance remains accurate, while money movement and expense analysis use only transactions within the stated period.</div>'+
    '<section class="signatures"><h3>Reviewed and Approved by</h3><div class="signature-grid">'+
    '<div class="signature"><div class="signature-line"></div><strong>ENGR. GODWIN TRISTAN VELASCO, SO2</strong><span>Registered Civil Engineer</span></div>'+
    '<div class="signature"><div class="signature-line"></div><strong>ENGR. CHRISTIAN VALEZA, SO2</strong><span>Registered Civil Engineer</span></div>'+
    '<div class="signature"><div class="signature-line"></div><strong>VISMARK NAVARRO</strong><span>Architectural Designer / BIM Modeler</span></div>'+
    '</div></section>'+
    '<footer class="footer"><span class="confidential">LIWO FINANCE · CONFIDENTIAL FINANCIAL REPORT</span><span>Generated from LIWO Finance</span></footer>'+
    '</main></body></html>';
}

function buildAutomatedReportEmailHtml_(type,label){
  const isMonthly=String(type||"").toLowerCase()==="monthly";
  const reportType=isMonthly?"MONTHLY":"WEEKLY";
  const accent="#123451";
  const gold="#c88722";
  const safeLabel=String(label||"Financial Report").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  return '<!doctype html><html><body style="margin:0;padding:0;background:#eef2f5;font-family:Arial,Helvetica,sans-serif;color:#17364f;">'
    +'<div style="padding:32px 12px;">'
    +'<div style="max-width:700px;margin:0 auto;background:#ffffff;border:1px solid #dbe3e9;border-radius:14px;overflow:hidden;">'
    +'<div style="background:'+accent+';padding:28px 32px 25px;">'
    +'<div style="font-size:11px;letter-spacing:2.2px;font-weight:800;color:#d9e5ec;">LIWO</div>'
    +'<div style="margin-top:5px;font-size:21px;line-height:1.25;font-weight:800;color:#ffffff;">LIWO ENGINEERING CONSULTANCY</div>'
    +'<div style="margin-top:7px;font-size:12px;color:#cbdbe5;">Finance &amp; Project Management</div>'
    +'</div>'
    +'<div style="height:4px;background:'+gold+';"></div>'
    +'<div style="padding:30px 32px;">'
    +'<div style="display:inline-block;padding:6px 10px;border-radius:3px;background:#f7efe2;color:'+gold+';font-size:9px;font-weight:800;letter-spacing:1px;">AUTOMATED '+reportType+' REPORT</div>'
    +'<h1 style="margin:15px 0 7px;font-size:23px;line-height:1.3;color:'+accent+';">Financial Report Ready</h1>'
    +'<div style="font-size:13px;font-weight:700;color:#52697b;line-height:1.5;">'+safeLabel+'</div>'
    +'<p style="margin:17px 0 22px;font-size:13px;line-height:1.7;color:#607487;">Your scheduled LIWO Finance report has been generated successfully. The complete premium financial report is attached to this email as a PDF for review and record-keeping.</p>'
    +'<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;margin:0 0 22px;">'
    +'<tr><td style="padding:0 5px 0 0;width:50%;"><div style="border:1px solid #dfe6eb;border-radius:8px;padding:15px;background:#f8fafb;"><div style="font-size:9px;font-weight:800;letter-spacing:.7px;color:'+gold+';">REPORT TYPE</div><div style="margin-top:5px;font-size:13px;font-weight:800;color:'+accent+';">'+reportType+' FINANCIAL REPORT</div></div></td>'
    +'<td style="padding:0 0 0 5px;width:50%;"><div style="border:1px solid #dfe6eb;border-radius:8px;padding:15px;background:#f8fafb;"><div style="font-size:9px;font-weight:800;letter-spacing:.7px;color:'+gold+';">ATTACHMENT</div><div style="margin-top:5px;font-size:13px;font-weight:800;color:'+accent+';">Premium PDF Report</div></div></td></tr></table>'
    +'<div style="border-left:4px solid '+gold+';background:#f7f9fa;padding:14px 16px;margin-bottom:24px;">'
    +'<div style="font-size:11px;font-weight:800;color:'+accent+';">WHAT’S INSIDE</div>'
    +'<div style="margin-top:5px;font-size:12px;line-height:1.65;color:#607487;">Portfolio financial summary, project-level financial review, expense movement, management commentary, reporting-period notes, and executive review/signature section.</div>'
    +'</div>'
    +'<div style="padding-top:18px;border-top:1px solid #e2e8ed;font-size:11px;line-height:1.7;color:#718395;">'
    +'<strong style="color:'+accent+';">LIWO Finance</strong><br>'
    +'Automated Financial Reporting System<br>'
    +'This is an automatically generated message. Please retain the attached report for your records.'
    +'</div>'
    +'</div>'
    +'<div style="background:#f7f9fa;border-top:1px solid #e2e8ed;padding:15px 32px;text-align:center;font-size:9px;color:#7b8b97;">LIWO ENGINEERING CONSULTANCY · Generated from LIWO Finance</div>'
    +'</div></div></body></html>';
}

function generateWeeklyFinancialReport_(u){
  adminOnly(u);
  const tz=Session.getScriptTimeZone()||"Asia/Manila";
  const now=new Date();
  // Rolling last 7 calendar days, including today.
  const endDate=new Date(now);
  const startDate=new Date(now.getTime()-6*24*60*60*1000);
  const startKey=Utilities.formatDate(startDate,tz,"yyyy-MM-dd");
  const endKey=Utilities.formatDate(endDate,tz,"yyyy-MM-dd");
  const start=Utilities.formatDate(startDate,tz,"MMM d, yyyy");
  const end=Utilities.formatDate(endDate,tz,"MMM d, yyyy");
  const label="Weekly Financial Report: "+start+" – "+end;
  const rows=getProjectFinancialSummary_(startDate,endDate);
  const html=buildFinancialReportHtml_(label,rows);

  const folder=getOrCreateReceiptFolder_();
  const base="LIWO Weekly Financial Report "+endKey;
  const file=folder.createFile(Utilities.newBlob(html,"text/html",base+".html"));
  const pdfFile=folder.createFile(file.getAs(MimeType.PDF).setName(base+".pdf"));
  try{file.setTrashed(true)}catch(_){}

  const recipients=getReportRecipients_();
  if(recipients.length){
    MailApp.sendEmail({
      to:recipients.join(","),
      subject:"LIWO Finance — "+label,
      htmlBody:buildAutomatedReportEmailHtml_("Weekly",label),
      attachments:[pdfFile.getBlob()]
    });
  }
  audit_(u,"WEEKLY_FINANCIAL_REPORT","SYSTEM",label,JSON.stringify({pdfUrl:pdfFile.getUrl(),recipients:recipients,start:startKey,end:endKey}));
  return {ok:true,period:label,pdfUrl:pdfFile.getUrl(),recipients:recipients,projects:rows};
}
function generateMonthlyFinancialReport_(u){
  adminOnly(u);
  const tz=Session.getScriptTimeZone()||"Asia/Manila";
  const now=new Date();
  const startDate=new Date(now.getFullYear(),now.getMonth()-1,1);
  const endDate=new Date(now.getFullYear(),now.getMonth(),0);
  const start=Utilities.formatDate(startDate,tz,"MMMM d, yyyy");
  const end=Utilities.formatDate(endDate,tz,"MMMM d, yyyy");
  const label="Monthly Financial Report: "+start+" – "+end;
  const rows=getProjectFinancialSummary_(startDate,endDate);
  const html=buildFinancialReportHtml_(label,rows,"Monthly");
  const folder=getOrCreateReceiptFolder_();
  const base="LIWO Monthly Financial Report "+Utilities.formatDate(startDate,tz,"yyyy-MM");
  const file=folder.createFile(Utilities.newBlob(html,"text/html",base+".html"));
  const pdfFile=folder.createFile(file.getAs(MimeType.PDF).setName(base+".pdf"));
  try{file.setTrashed(true)}catch(_){}
  const recipients=getReportRecipients_();
  if(recipients.length){
    MailApp.sendEmail({
      to:recipients.join(","),
      subject:"LIWO Finance — "+label,
      htmlBody:buildAutomatedReportEmailHtml_("Monthly",label),
      attachments:[pdfFile.getBlob()]
    });
  }
  audit_(u,"MONTHLY_FINANCIAL_REPORT","SYSTEM",label,JSON.stringify({pdfUrl:pdfFile.getUrl(),recipients:recipients,start:Utilities.formatDate(startDate,tz,"yyyy-MM-dd"),end:Utilities.formatDate(endDate,tz,"yyyy-MM-dd")}));
  return {ok:true,period:label,pdfUrl:pdfFile.getUrl(),recipients:recipients,projects:rows};
}

function installWeeklyReportTrigger_(u){
  adminOnly(u);
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(t.getHandlerFunction()==="runScheduledWeeklyReport") ScriptApp.deleteTrigger(t);
  });
  const cfg=getReportConfig_(u).config;
  if(!cfg.weeklyEnabled) return {ok:true,enabled:false};
  const days={
    SUNDAY:ScriptApp.WeekDay.SUNDAY,MONDAY:ScriptApp.WeekDay.MONDAY,TUESDAY:ScriptApp.WeekDay.TUESDAY,
    WEDNESDAY:ScriptApp.WeekDay.WEDNESDAY,THURSDAY:ScriptApp.WeekDay.THURSDAY,FRIDAY:ScriptApp.WeekDay.FRIDAY,SATURDAY:ScriptApp.WeekDay.SATURDAY
  };
  ScriptApp.newTrigger("runScheduledWeeklyReport")
    .timeBased().onWeekDay(days[cfg.weeklyDay]||ScriptApp.WeekDay.SUNDAY)
    .atHour(Number(cfg.weeklyHour)||17).create();
  return {ok:true,enabled:true,day:cfg.weeklyDay,hour:Number(cfg.weeklyHour)||17};
}

function runScheduledWeeklyReport(){
  const recipients=getReportRecipients_();
  if(!recipients.length) throw Error("No automated report recipients configured.");
  const tz=Session.getScriptTimeZone()||"Asia/Manila";
  const now=new Date();
  const endDate=new Date(now.getFullYear(),now.getMonth(),now.getDate());
  const startDate=new Date(now.getFullYear(),now.getMonth(),now.getDate()-6);
  const start=Utilities.formatDate(startDate,tz,"MMMM d, yyyy");
  const end=Utilities.formatDate(endDate,tz,"MMMM d, yyyy");
  const startKey=Utilities.formatDate(startDate,tz,"yyyy-MM-dd");
  const endKey=Utilities.formatDate(endDate,tz,"yyyy-MM-dd");
  const label="Weekly Financial Report: "+start+" – "+end;
  const rows=getProjectFinancialSummary_(startDate,endDate);
  const html=buildFinancialReportHtml_(label,rows,"Weekly");
  const folder=getOrCreateReceiptFolder_();
  const base="LIWO Weekly Financial Report "+endKey;
  const file=folder.createFile(Utilities.newBlob(html,"text/html",base+".html"));
  const pdfFile=folder.createFile(file.getAs(MimeType.PDF).setName(base+".pdf"));
  try{file.setTrashed(true)}catch(_){}
  MailApp.sendEmail({
    to:recipients.join(","),
    subject:"LIWO Finance — "+label,
    htmlBody:buildAutomatedReportEmailHtml_("Weekly",label),
    attachments:[pdfFile.getBlob()]
  });
  audit_({username:"SYSTEM",name:"SYSTEM",role:"Admin"},"WEEKLY_FINANCIAL_REPORT","SYSTEM",label,JSON.stringify({pdfUrl:pdfFile.getUrl(),recipients:recipients,start:startKey,end:endKey}));
  return {ok:true,period:label,pdfUrl:pdfFile.getUrl(),recipients:recipients,projects:rows};
}

function setupAutomatedReports(){
  const recipients=getReportRecipients_();
  if(!recipients.length) throw Error("No automated report recipients configured.");

  const setupNow=new Date();

  // Store the recipients in the sheet too, so the Admin UI can display them.
  const sh=ensureSheet_("report_recipients",["Email","Enabled","UpdatedAt","UpdatedBy"]);
  // Only seed the sheet when it has no recipients yet. Do not overwrite
  // recipients later selected by the Administrator from the web app.
  if(sh.getLastRow()<=1){
    sh.getRange(2,1,recipients.length,4).setValues(
      recipients.map(function(e){return [e,true,setupNow,"SYSTEM"];})
    );
  }

  // Seed the weekly settings only when they do not already exist.
  // This prevents running setupAutomatedReports() from silently resetting
  // a schedule chosen by the Administrator in the web app.
  const cfg=ensureSheet_("weekly_report_config",["Key","Value","UpdatedAt","UpdatedBy"]);
  if(cfg.getLastRow()<=1){
    cfg.getRange(2,1,3,4).setValues([
      ["weeklyEnabled",true,setupNow,"SYSTEM"],
      ["weeklyDay",CONFIG.REPORT_WEEKLY_DAY,setupNow,"SYSTEM"],
      ["weeklyHour",CONFIG.REPORT_WEEKLY_HOUR,setupNow,"SYSTEM"]
    ]);
  }

  const days={
    SUNDAY:ScriptApp.WeekDay.SUNDAY,MONDAY:ScriptApp.WeekDay.MONDAY,
    TUESDAY:ScriptApp.WeekDay.TUESDAY,WEDNESDAY:ScriptApp.WeekDay.WEDNESDAY,
    THURSDAY:ScriptApp.WeekDay.THURSDAY,FRIDAY:ScriptApp.WeekDay.FRIDAY,
    SATURDAY:ScriptApp.WeekDay.SATURDAY
  };

  // Remove duplicate LIWO report triggers.
  ScriptApp.getProjectTriggers().forEach(function(trigger){
    const fn=trigger.getHandlerFunction();
    if(fn==="runScheduledWeeklyReport" || fn==="sendMonthlyFinancialReport"){
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const savedCfg=getReportConfig_({role:"Admin",username:"SYSTEM",name:"SYSTEM"}).config;
  if(savedCfg.weeklyEnabled){
    ScriptApp.newTrigger("runScheduledWeeklyReport")
      .timeBased()
      .onWeekDay(days[savedCfg.weeklyDay]||ScriptApp.WeekDay.SUNDAY)
      .atHour(Number(savedCfg.weeklyHour)||17)
      .create();
  }

  ScriptApp.newTrigger("sendMonthlyFinancialReport")
    .timeBased()
    .onMonthDay(1)
    .atHour(8)
    .create();

  return {
    ok:true,
    recipients:recipients,
    weeklyDay:savedCfg.weeklyDay,
    weeklyHour:savedCfg.weeklyHour,
    monthlyDay:1,
    monthlyHour:8
  };
}

function sendAutomatedReportTestEmail(u) {
  // Direct executions from the Apps Script editor do not provide
  // the web-app user object. Treat those executions as a system/admin test.
  if (u && typeof u === "object") {
    adminOnly(u);
  } else {
    u = {
      username: "SYSTEM",
      name: "LIWO Administrator",
      role: "Admin"
    };
  }

  const recipients = getReportRecipients_();

  if (!recipients || !recipients.length) {
    throw new Error("No automated report recipients configured.");
  }

  const subject = "LIWO Finance - Automated Report Test";
  const body =
    "This is a test email from the LIWO Finance automated reporting system.\n\n" +
    "Recipients:\n" +
    recipients.join("\n") +
    "\n\nThe automated email system is working correctly.";

  MailApp.sendEmail({
    to: recipients.join(","),
    subject: subject,
    body: body,
    htmlBody: buildAutomatedReportEmailHtml_("Weekly","Automated Report System Test")
  });

  audit(
    "AUTOMATED_REPORT_TEST",
    u,
    "Automated financial report test email sent to: " + recipients.join(", ")
  );

  return {
    ok: true,
    recipients: recipients
  };
}

function generateProjectFinancialReport_(u, clientId){
  adminOnly(u);
  const client=requireClient_(clientId,true);
  const id=client.id;

  const payments=readSheet_("payments").filter(function(p){return String(p.ClientID||"").trim()===id;});
  const expenses=readSheet_("expenses").filter(function(e){return String(e.ClientID||"").trim()===id;});
  const milestones=readSheet_("milestones").filter(function(m){return String(m.ClientID||"").trim()===id;});
  const tools=readSheet_("tools").filter(function(t){return String(t.ClientID||"").trim()===id;});

  const contract=Number(client.ContractValue||client.Budget||0)||0;
  const paid=payments.reduce(function(s,p){return s+(Number(p.Amount||0)||0)},0);
  const spent=expenses.reduce(function(s,e){return s+(Number(e.Amount||0)||0)},0);
  const profit=paid-spent;
  const margin=paid>0?(profit/paid)*100:0;
  const uncollected=Math.max(0,contract-paid);
  const budgetUsed=contract>0?(spent/contract)*100:0;

  const money=function(n){return "₱"+Number(n||0).toLocaleString("en-PH",{minimumFractionDigits:2,maximumFractionDigits:2});};
  const esc=function(s){return String(s??"").replace(/[&<>"']/g,function(m){return({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#039;"})[m]});};

  const paymentRows=payments.map(function(p){
    return "<tr><td>"+esc(p.Date||p.Timestamp||"")+"</td><td>"+esc(p.Description||p.Particulars||"Payment")+"</td><td>"+money(p.Amount)+"</td><td>"+esc(p.Method||p.PaymentMethod||"")+"</td></tr>";
  }).join("") || '<tr><td colspan="4">No payments recorded.</td></tr>';

  const expenseRows=expenses.map(function(e){
    return "<tr><td>"+esc(e.Date||e.Timestamp||"")+"</td><td>"+esc(e.Category||"")+"</td><td>"+esc(e.Description||e.Particulars||"")+"</td><td>"+money(e.Amount)+"</td></tr>";
  }).join("") || '<tr><td colspan="4">No transactions recorded.</td></tr>';

  const milestoneRows=milestones.map(function(m){
    return "<tr><td>"+esc(m.Milestone||m.Name||"")+"</td><td>"+esc(m.DueDate||"")+"</td><td>"+money(m.Amount)+"</td><td>"+esc(m.Status||"")+"</td></tr>";
  }).join("") || '<tr><td colspan="4">No milestones recorded.</td></tr>';

  const toolRows=tools.map(function(t){
    return "<tr><td>"+esc(t.Tool||t.ToolName||"")+"</td><td>"+esc(t.ToolId||t.ToolID||"")+"</td><td>"+esc(t.BorrowedBy||"")+"</td><td>"+esc(t.Status||"")+"</td></tr>";
  }).join("") || '<tr><td colspan="4">No tool records.</td></tr>';

  const now=new Date(), tz=Session.getScriptTimeZone()||"Asia/Manila";
  const projectName=String(client.ProjectName||client.ClientName||"").trim();
  const ref=String(client.Reference||"").trim();

  const html='<!doctype html><html><head><meta charset="utf-8"><style>'+
    'body{font-family:Arial,sans-serif;color:#17324a;padding:34px;font-size:12px}'+
    'h1{margin:0 0 5px;font-size:24px}.sub{color:#718395;margin-bottom:20px}'+
    '.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:9px;margin-bottom:20px}'+
    '.kpi{border:1px solid #d9e2ea;border-radius:9px;padding:12px}.kpi span{display:block;color:#718395;font-size:10px}.kpi b{display:block;font-size:16px;margin-top:4px}'+
    'h2{font-size:16px;margin:24px 0 8px;border-bottom:1px solid #d9e2ea;padding-bottom:6px}'+
    'table{width:100%;border-collapse:collapse;margin-bottom:12px}th,td{border:1px solid #d9e2ea;padding:7px;text-align:left}th{background:#eef4f8;font-size:10px}'+
    '.health{font-weight:bold}.foot{margin-top:25px;color:#718395;font-size:9px}'+
    '</style></head><body>'+
    '<h1>LIWO Finance — Complete Project Financial Report</h1>'+
    '<div class="sub">'+esc(projectName)+(ref?" • "+esc(ref):"")+'<br>Generated '+Utilities.formatDate(now,tz,"MMM d, yyyy h:mm a")+'</div>'+
    '<div class="kpis">'+
      '<div class="kpi"><span>Contract / Budget</span><b>'+money(contract)+'</b></div>'+
      '<div class="kpi"><span>Client Payments</span><b>'+money(paid)+'</b></div>'+
      '<div class="kpi"><span>Total Expenses</span><b>'+money(spent)+'</b></div>'+
      '<div class="kpi"><span>Cash Profit</span><b>'+money(profit)+'</b></div>'+
      '<div class="kpi"><span>Uncollected</span><b>'+money(uncollected)+'</b></div>'+
      '<div class="kpi"><span>Profit Margin</span><b>'+margin.toFixed(1)+'%</b></div>'+
      '<div class="kpi"><span>Budget Used</span><b>'+budgetUsed.toFixed(1)+'%</b></div>'+
      '<div class="kpi"><span>Financial Health</span><b class="health">'+esc(budgetUsed>=100||profit<0?"Critical":(budgetUsed>=80||contract>0&&paid/contract<.5?"Attention":"Healthy"))+'</b></div>'+
    '</div>'+
    '<h2>Client Payments</h2><table><thead><tr><th>Date</th><th>Description</th><th>Amount</th><th>Method</th></tr></thead><tbody>'+paymentRows+'</tbody></table>'+
    '<h2>Fund Transactions / Expenses</h2><table><thead><tr><th>Date</th><th>Category</th><th>Description</th><th>Amount</th></tr></thead><tbody>'+expenseRows+'</tbody></table>'+
    '<h2>Payment Milestones</h2><table><thead><tr><th>Milestone</th><th>Due Date</th><th>Amount</th><th>Status</th></tr></thead><tbody>'+milestoneRows+'</tbody></table>'+
    '<h2>Construction Tools</h2><table><thead><tr><th>Tool</th><th>ID</th><th>Borrowed By</th><th>Status</th></tr></thead><tbody>'+toolRows+'</tbody></table>'+
    '<div class="foot">LIWO Finance Tracker • Project-specific report. Financial figures are isolated by ClientID/Project.</div>'+
    '</body></html>';

  const folder=getOrCreateReceiptFolder_();
  const safe=projectName.replace(/[^\w\-]+/g,"_").slice(0,70)||id;
  const stamp=Utilities.formatDate(now,tz,"yyyy-MM-dd_HHmm");
  const file=folder.createFile(Utilities.newBlob(html,"text/html","LIWO Project Report - "+safe+" - "+stamp+".html"));
  const pdf=folder.createFile(file.getAs(MimeType.PDF).setName("LIWO Project Financial Report - "+safe+" - "+stamp+".pdf"));
  try{file.setTrashed(true)}catch(_){}
  audit_(u,"PROJECT_FINANCIAL_REPORT",id,"Complete project financial report generated",pdf.getUrl());
  return {ok:true,pdfUrl:pdf.getUrl(),projectId:id,project:projectName,reportType:"project"};
}

function doPost(e){try{let r=JSON.parse(e.postData.contents||"{}");switch(r.action){case"health":return json({ok:true,service:"LIWO Finance Tracker",version:"2026-08-29"});
case"login":return json(login(r));case"registerFinance":return json(registerFinance(r));case"listClients":return json(withAuth(r,listClients));case"dashboard":return json(withAuth(r,dashboard));case"executiveDashboard":return json(withAuth(r,executiveDashboard));case"getExecutiveDashboard":return json(withAuth(r,executiveDashboard));case"saveReportSettings":return json(withAuth(r,saveReportSettings));case"getReportSettings":return json(withAuth(r,getReportSettings));
case"setReportRecipients":return json(withAuth(r,setReportRecipients_));
case"getReportConfig":return json(withAuth(r,getReportConfig_));
case"saveReportConfig":return json(withAuth(r,saveReportConfig_));
case"installWeeklyReportTrigger":return json(withAuth(r,installWeeklyReportTrigger_));
case"generateWeeklyReport":return json(withAuth(r,generateWeeklyFinancialReport_));
case"generateMonthlyReport":return json(withAuth(r,generateMonthlyFinancialReport_));
case"sendAutomatedReportTestEmail":return json(withAuth(r,sendAutomatedReportTestEmail));
case"generateProjectFinancialReport":return json(withAuth(r,function(rr,uu){return generateProjectFinancialReport_(uu,rr.clientId);}));
case"generateFinancialReport":return json(withAuth(r,generateFinancialReport));case"installMonthlyReportTrigger":return json(withAuth(r,installMonthlyReportTrigger));case"sendMonthlyFinancialReport":return json(withAuth(r,sendMonthlyFinancialReport));case"projectFinancialDashboard":return json(withAuth(r,projectFinancialDashboard));case"projectDashboard":return json(withAuth(r,projectFinancialDashboard));case"projectWorkspace":return json(withAuth(r,projectFinancialDashboard));
case"addPayment":return json(withAuth(r,addPayment));case"updatePayment":return json(withAuth(r,updatePayment));case"editPayment":return json(withAuth(r,editPayment));case"addExpense":return json(withAuth(r,addExpense));case"updateExpense":return json(withAuth(r,updateExpense));case"editExpense":return json(withAuth(r,editExpense));case"listUsers":return json(withAuth(r,listUsers));
case"upsertUser":return json(withAuth(r,upsertUser));case"listTools":return json(withAuth(r,listTools));case"tools":return json(withAuth(r,listTools));case"constructionTools":return json(withAuth(r,listTools));case"listConstructionTools":return json(withAuth(r,listTools));case"getTools":return json(withAuth(r,listTools));case"addTool":return json(withAuth(r,addTool));case"updateTool":return json(withAuth(r,updateTool));case"cashBalances":return json(withAuth(r,cashBalances));case"cashPosition":return json(withAuth(r,cashBalances));case"getCashPosition":return json(withAuth(r,cashBalances));case"getCashBalances":return json(withAuth(r,cashBalances));case"updateCashBalance":return json(withAuth(r,updateCashBalance));case"updateCashBalances":return json(withAuth(r,updateCashBalances));case"changeInvite":return json(withAuth(r,changeInvite));case"reopenRegistration":return json(withAuth(r,reopenRegistration));case"upsertClient":return json(withAuth(r,upsertClient));case"archiveClient":return json(withAuth(r,archiveClient));case"restoreClient":return json(withAuth(r,restoreClient));case"deleteClient":return json(withAuth(r,deleteClient));case"notifications":return json(withAuth(r,notifications));case"markNotificationsRead":return json(withAuth(r,markNotificationsRead));case"listMilestones":return json(withAuth(r,listMilestones));case"addMilestone":return json(withAuth(r,addMilestone));case"updateMilestone":return json(withAuth(r,updateMilestone));case"verifyReceipt":return json(withAuth(r,verifyReceipt));case"reconcileCash":return json(withAuth(r,reconcileCash));case"getProjectBudget":return json(withAuth(r,getProjectBudget));case"saveProjectBudget":return json(withAuth(r,saveProjectBudget));case"returnTool":return json(withAuth(r,returnTool));case"deletePayment":return json(withAuth(r,deletePayment));case"deleteExpense":return json(withAuth(r,deleteExpense));case"listBills":return json(withAuth(r,listBills));case"addBill":return json(withAuth(r,addBill));case"updateBill":return json(withAuth(r,updateBill));case"editBill":return json(withAuth(r,editBill));case"deleteBill":return json(withAuth(r,deleteBill));case"saveCalculation":return json(withAuth(r,saveCalculation));case"listCalculations":return json(withAuth(r,listCalculations));
case"deleteUser":return json(withAuth(r,deleteUser));case"deactivateUser":return json(withAuth(r,deactivateUser));case"reactivateUser":return json(withAuth(r,reactivateUser));case"setUserActive":return json(withAuth(r,setUserActive));case"listReceipts":return json(withAuth(r,listReceipts));case"listReceipt":return json(withAuth(r,listReceipts));case"getReceipts":return json(withAuth(r,listReceipts));case"getReceiptList":return json(withAuth(r,listReceipts));case"receipts":return json(withAuth(r,listReceipts));case"receiptGallery":return json(withAuth(r,listReceipts));case"getReceiptGallery":return json(withAuth(r,listReceipts));case"loadReceipts":return json(withAuth(r,listReceipts));case"getReceiptLibrary":return json(withAuth(r,listReceipts));case"migrateLegacyToolClientIds":return json(withAuth(r,migrateLegacyToolClientIds_));default:return json({ok:false,error:"Unknown action: "+String(r.action||"")})}}catch(x){return json({ok:false,error:String(x.message||x)})}}
function json(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function ensureHeaders_(sh,headers){if(sh.getLastRow()===0){sh.getRange(1,1,1,headers.length).setValues([headers]);return}let existing=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),1)).getValues()[0].map(String);if(existing.length<headers.length){sh.getRange(1,existing.length+1,1,headers.length-existing.length).setValues([headers.slice(existing.length)]);}}
function receiptRootFolder_(){
  let folders=DriveApp.getFoldersByName(CONFIG.RECEIPT_FOLDER_NAME);
  return folders.hasNext()?folders.next():DriveApp.createFolder(CONFIG.RECEIPT_FOLDER_NAME);
}
function safeFolderName_(name){
  return String(name||"Unassigned").replace(/[\\/:*?"<>|#%{}~]/g,"_").replace(/\s+/g," ").trim().slice(0,120)||"Unassigned";
}
function projectReceiptFolder_(client){
  if(!client||!client.id)throw Error("A valid client/project is required.");
  let root=receiptRootFolder_();
  // ClientID is part of the folder name so two projects with the same
  // display name can never share a receipt folder.
  let folderName=safeFolderName_(client.name||"Unassigned")+" ["+String(client.id)+"]";
  let folders=root.getFoldersByName(folderName);
  return folders.hasNext()?folders.next():root.createFolder(folderName);
}
function saveReceipt_(dataUrl,fileName,u,client){
  if(!dataUrl)return null;
  if(!client||!client.id)throw Error("A client/project is required for receipt upload.");
  let m=String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if(!m)throw Error("Invalid receipt file data.");
  let bytes=Utilities.base64Decode(m[2]);
  if(bytes.length>CONFIG.MAX_RECEIPT_BYTES)throw Error("Receipt image is too large. Please use a file under 5 MB.");
  let mime=m[1].toLowerCase();
  if(!/^image\/(jpeg|png|webp)$/.test(mime)&&mime!=="application/pdf")throw Error("Receipt must be JPG, PNG, WEBP, or PDF.");
  let safe=String(fileName||"receipt").replace(/[^a-zA-Z0-9._-]/g,"_");
  let stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd_HHmmss");
  let blob=Utilities.newBlob(bytes,mime,stamp+"_"+safe);
  let file=projectReceiptFolder_(client).createFile(blob);
  try{file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW)}catch(e){}
  return{id:file.getId(),url:file.getUrl(),name:file.getName(),uploadedBy:u.name,projectFolder:client.name};
}
function ss(){return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)}
function num(v){return Number(v)||0}
function hash_(s){return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s),Utilities.Charset.UTF_8).map(b=>{let x=(b<0?b+256:b).toString(16);return x.length===1?"0"+x:x}).join("")}
function rows(n){let sh=ss().getSheetByName(n);if(!sh)return[];let v=sh.getDataRange().getValues();return v.length>1?v.slice(1):[]}
function ensureSheet_(name,headers){let s=ss(),sh=s.getSheetByName(name);if(!sh){sh=s.insertSheet(name);ensureHeaders_(sh,headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight("bold");}else ensureHeaders_(sh,headers);return sh;}
function settingsMap(){let v=ss().getSheetByName("settings").getDataRange().getValues(),m={};v.slice(1).forEach(r=>m[r[0]]=r[1]);return m}
function login(r){let v=ss().getSheetByName("users").getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===String(r.username||"")&&String(x[2])===hash_(r.password||"")&&String(x[4]).toLowerCase()!=="false");if(i<0)return{ok:false,error:"Invalid username or password."};let row=v[i+1],u={username:String(row[0]),name:String(row[1]),role:String(row[3])},t=Utilities.getUuid();CacheService.getScriptCache().put("session_"+t,JSON.stringify(u),CONFIG.SESSION_HOURS*3600);audit("LOGIN",u,"Successful login");return{ok:true,token:t,user:u}}
function withAuth(r,f){let raw=CacheService.getScriptCache().get("session_"+String(r.token||""));if(!raw)return{ok:false,error:"Session expired. Please sign in again."};return f(r,JSON.parse(raw))}
function adminOnly(u){if(u.role!=="Admin")throw Error("Administrator access required.")}
function registerFinance(r){
 let sm=settingsMap();if(String(sm.RegistrationOpen).toLowerCase()==="false")return{ok:false,error:"LIWO Executive registration is closed. Ask the administrator to reopen it."};
 if(!sm.InviteCodeHash||hash_(r.inviteCode||"")!==String(sm.InviteCodeHash))return{ok:false,error:"Invalid invitation code."};
 if(!r.name||!r.username||!r.password)throw Error("All fields are required.");if(String(r.password).length<8)throw Error("Password must be at least 8 characters.");
 let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues();if(v.slice(1).some(x=>String(x[0]).toLowerCase()===String(r.username).toLowerCase()))throw Error("Username already exists.");
 let count=v.slice(1).filter(x=>String(x[3])==="Finance"&&String(x[4]).toLowerCase()!=="false").length;if(count>=3){setSetting_("RegistrationOpen",false);return{ok:false,error:"The 3 LIWO Executive accounts have already been registered. Registration is now closed."}}
 sh.appendRow([String(r.username),String(r.name),hash_(r.password),"Finance",true,new Date(),new Date()]);
 count++;if(count>=3)setSetting_("RegistrationOpen",false);return{ok:true};
}

function listReceipts(r,u){
  /*
   * Receipt gallery endpoint.
   * Reads the LIWO Finance Receipts Drive folder, including:
   *  - receipt files inside project subfolders
   *  - receipt files accidentally/directly stored in the root folder
   *  - nested folders, so older uploads are not lost
   *
   * The project map is also read from the Expenses sheet so root-level
   * files that were attached to a transaction can still show their project.
   */
  let project=String(r.project||"").trim();
  let clientId=String(r.clientId||"").trim();
  if(clientId){
    const c0=requireClient_(clientId,true);
    project=c0.name;
  }else if(project){
    const c0=requireUniqueClientByName_(project,true);
    clientId=c0.id;
    project=c0.name;
  }
  let rootIt=DriveApp.getFoldersByName(CONFIG.RECEIPT_FOLDER_NAME);
  if(!rootIt.hasNext()){
    let root=DriveApp.createFolder(CONFIG.RECEIPT_FOLDER_NAME);
    return{ok:true,receipts:[],projects:[]};
  }

  let root=rootIt.next();

  // Map stored Drive file IDs to their project/client from the spreadsheet.
  let fileClientMap={};
  try{
    rows("expenses").forEach(x=>{
      let fileId=String(x[15]||"").trim();
      let expenseClientId=String(x[2]||"").trim();
      if(fileId && expenseClientId)fileClientMap[fileId]=expenseClientId;
    });
  }catch(e){}

  // Top-level project folders.
  let projects=[];
  let projectFolderNames={};
  let rootFolders=root.getFolders();
  while(rootFolders.hasNext()){
    let folder=rootFolders.next();
    let name=folder.getName();
    projects.push(name);
    projectFolderNames[name]=true;
  }
  projects.sort();

  let receipts=[];
  let seen={};

  function addFile_(f,folderProject){
    let id=f.getId();
    if(seen[id])return;
    seen[id]=true;

    let mime=String(f.getMimeType()||"");
    let preview="";

    // Return an embedded preview for normal-sized receipt photos.
    // This makes the gallery display the actual image instead of a generic icon.
    if(/^image\/(jpeg|png|webp)$/i.test(mime) && f.getSize()<=2500000){
      try{
        preview="data:"+mime+";base64,"+Utilities.base64Encode(f.getBlob().getBytes());
      }catch(e){}
    }

    // Keep existing Drive files viewable when possible.
    try{
      if(f.getSharingAccess()===DriveApp.Access.PRIVATE){
        f.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
      }
    }catch(e){}

    const mappedClientId=fileClientMap[id]||"";
    let mappedProject="";
    if(mappedClientId){
      const mappedClient=clientNameMap()[mappedClientId];
      mappedProject=mappedClient?mappedClient.name:"";
    }
    if(!mappedProject) mappedProject=folderProject||"Unassigned";

    // For a project-scoped receipt request, ONLY an exact ClientID match
    // is accepted. Folder-name matching is intentionally not used as a
    // security boundary because project names can be duplicated/renamed.
    if(clientId){
      if(mappedClientId!==clientId)return;
      if(!mappedClientId)return;
    }

    receipts.push({
      id:id,
      name:f.getName(),
      url:f.getUrl(),
      viewUrl:"https://drive.google.com/uc?export=view&id="+encodeURIComponent(id),
      mimeType:mime,
      size:f.getSize(),
      createdAt:f.getDateCreated(),
      updatedAt:f.getLastUpdated(),
      project:mappedProject,
      preview:preview
    });
  }

  function walkFolder_(folder,topProject){
    let files=folder.getFiles();
    while(files.hasNext())addFile_(files.next(),topProject);

    let subs=folder.getFolders();
    while(subs.hasNext())walkFolder_(subs.next(),topProject);
  }

  // Include files directly in the main receipt folder.
  // Their project is recovered from the expense record when possible.
  let rootFiles=root.getFiles();
  while(rootFiles.hasNext())addFile_(rootFiles.next(),"");

  // Include all project folders recursively.
  let projectFolders=root.getFolders();
  while(projectFolders.hasNext()){
    let folder=projectFolders.next();
    walkFolder_(folder,folder.getName());
  }

  let vm=receiptVerificationMap_();receipts.forEach(x=>{let v=vm[String(x.id)]||{};x.verificationStatus=v.status||"Pending";x.verifiedBy=v.verifiedBy||"";x.verifiedAt=v.verifiedAt||null;});
  receipts.sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  return{ok:true,receipts:receipts,projects:projects};
}

/* Optional one-time migration helper for older tool rows that predate ClientID.
   It assigns a legacy tool row to a project only when the project name matches
   exactly one active client/project. No automatic cross-project guess is made. */
function migrateLegacyToolClientIds_(u){
  adminOnly(u);
  const sh=ss().getSheetByName("tools");
  if(!sh||sh.getLastRow()<2)return{ok:true,updated:0,skipped:0};
  const clients=clientObjects(false);
  const byName={};
  clients.forEach(function(c){
    const k=c.name.trim().toLowerCase();
    if(!byName[k])byName[k]=[];
    byName[k].push(c);
  });
  let updated=0,skipped=0;
  const v=sh.getDataRange().getValues();
  for(let i=1;i<v.length;i++){
    const cid=String(v[i][12]||"").trim();
    if(cid)continue;
    const name=String(v[i][4]||"").trim().toLowerCase();
    const matches=byName[name]||[];
    if(matches.length===1){
      sh.getRange(i+1,13).setValue(matches[0].id);
      updated++;
    }else{
      skipped++;
    }
  }
  audit("MIGRATE_LEGACY_TOOL_CLIENT_IDS",u,"Updated: "+updated+" | Skipped: "+skipped);
  return{ok:true,updated,skipped};
}

function notifications(r,u){
  let limit=Math.min(Math.max(Number(r.limit)||100,1),100);
  let sh=ss().getSheetByName("audit"),v=sh.getDataRange().getValues();
  let readSh=ensureSheet_("notification_reads",SHEETS.notification_reads),rv=readSh.getDataRange().getValues();
  let rr=rv.slice(1).find(x=>String(x[0])===String(u.username));
  let lastRead=rr?new Date(rr[1]).getTime():0;
  let items=v.slice(1).filter(x=>String(x[1]||"")!=="LOGIN").slice(-limit).reverse().map(x=>({
    time:x[0],action:String(x[1]||"Finance activity"),username:String(x[2]||""),name:String(x[3]||""),details:String(x[4]||""),
    unread:new Date(x[0]).getTime()>lastRead
  }));
  return{ok:true,notifications:items,unreadCount:items.filter(x=>x.unread).length,lastReadAt:lastRead?new Date(lastRead):null};
}
function markNotificationsRead(r,u){
  let sh=ensureSheet_("notification_reads",SHEETS.notification_reads),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===String(u.username)),now=new Date();
  if(i<0)sh.appendRow([u.username,now]);else sh.getRange(i+2,2).setValue(now);
  return{ok:true,unreadCount:0,lastReadAt:now};
}

function clientObjects(activeOnly){return rows("clients").filter(x=>!activeOnly||String(x[4]).toLowerCase()!=="false").map(x=>({id:String(x[0]),name:String(x[1]),reference:String(x[2]||""),budget:num(x[3]),active:String(x[4]).toLowerCase()!=="false"}))}
function clientNameMap(){let m={};clientObjects(false).forEach(c=>m[c.id]=c);return m}

/* ============================================================
   PROJECT-SCOPE SECURITY
   ------------------------------------------------------------
   All project-scoped reads/writes must validate the ClientID
   on the server. The frontend selection is NOT trusted.
   ============================================================ */

function requireClient_(clientId, requireActive){
  const id=String(clientId||"").trim();
  if(!id) throw Error("Client / project ID is required.");
  const cm=clientNameMap();
  const c=cm[id];
  if(!c) throw Error("Invalid client/project.");
  if(requireActive!==false && !c.active) throw Error("Client/project is inactive.");
  return c;
}

function requireClientFromRequest_(r, requireActive){
  return requireClient_(String(r&&r.clientId||"").trim(), requireActive);
}

function requireUniqueClientByName_(name, requireActive){
  const n=String(name||"").trim();
  if(!n) throw Error("Client / project is required.");
  const matches=clientObjects(false).filter(function(c){
    return c.name.trim().toLowerCase()===n.toLowerCase() &&
      (requireActive===false || c.active);
  });
  if(matches.length===0) throw Error("Invalid client/project.");
  if(matches.length>1) throw Error("Project name is not unique. Use the ClientID.");
  return matches[0];
}

/*
 * Verify that a spreadsheet row belongs to the project requested
 * by the caller. This prevents a caller from supplying an arbitrary
 * row number and editing/deleting another project's transaction.
 */
function requireRowClient_(sheetName, row, clientId, clientColumn){
  const sh=ss().getSheetByName(sheetName);
  if(!sh) throw Error("Sheet not found: "+sheetName);
  const n=Number(row);
  if(!Number.isInteger(n) || n<2 || n>sh.getLastRow())
    throw Error("Record not found.");
  const col=Number(clientColumn||3);
  const stored=String(sh.getRange(n,col).getValue()||"").trim();
  const requested=String(clientId||"").trim();
  if(!requested) throw Error("Client / project ID is required.");
  if(!stored || stored!==requested)
    throw Error("Project scope violation: this record does not belong to the selected client/project.");
  return {sheet:sh,row:n,clientId:stored};
}

function requireProjectScopedRead_(r){
  return requireClientFromRequest_(r,true);
}

function assertToolProject_(r, oldRow){
  const requested=String(r&&r.clientId||"").trim();
  if(!requested) throw Error("Client / project ID is required.");
  const c=requireClient_(requested,true);
  if(oldRow){
    const stored=String(oldRow[12]||"").trim();
    if(!stored || stored!==requested)
      throw Error("Project scope violation: this tool record belongs to another project or has no project ID.");
  }
  return c;
}

function listClients(r,u){
  return {ok:true,clients:clientObjects(true)};
}

function dashboard(r,u){
  const c=requireProjectScopedRead_(r);
  const cid=c.id;
  const cm=clientNameMap();

  const p=rows("payments")
    .filter(x=>String(x[2]||"").trim()===cid)
    .map((x,i)=>({row:rows("payments").indexOf(x)+2,date:x[1],clientName:c.name,clientId:cid,reference:x[3],description:x[4],amount:num(x[6]),method:x[7],enteredBy:x[9]}));

  const ex=rows("expenses")
    .filter(x=>String(x[2]||"").trim()===cid)
    .map((x,i)=>({row:rows("expenses").indexOf(x)+2,date:x[1],clientName:c.name,clientId:cid,type:x[3],category:x[4],payee:x[5],description:x[6],amount:num(x[7]),method:x[8],reference:x[9],approvedBy:x[10]||"",enteredBy:x[12],username:x[13],receiptUrl:x[14]||"",receiptId:x[15]||"",recordStatus:String(x[16]||"Recorded")}));

  const approvedEx=ex;
  const totalPayments=p.reduce((a,x)=>a+x.amount,0);
  const totalExpenses=approvedEx.filter(x=>x.type==="Expense").reduce((a,x)=>a+x.amount,0);
  const refunds=approvedEx.filter(x=>x.type==="Refund").reduce((a,x)=>a+x.amount,0);
  const income=approvedEx.filter(x=>x.type==="Other Income").reduce((a,x)=>a+x.amount,0);
  const budgetRemaining=Math.max(0,c.budget-totalExpenses);
  const cashPosition=totalPayments-totalExpenses+refunds+income;

  const budget=getProjectBudgetData_(cid).map(x=>{
    const actual=approvedEx
      .filter(y=>y.category===x.category&&y.type==="Expense")
      .reduce((a,y)=>a+y.amount,0);
    const b=num(x.budget);
    return{category:x.category,budget:b,actual,variance:b-actual,used:b?actual/b:0,remaining:Math.max(0,b-actual)};
  }).filter(x=>x.budget||x.actual);

  const recent=[...p.map(x=>({...x,type:"Client Payment"})),...ex.map(x=>({...x,type:x.type}))]
    .sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,12);

  // Activity is global by design; transaction data above is strictly project-scoped.
  const av=rows("audit").slice(-100).reverse().map(x=>({time:x[0],action:x[1],username:x[2],name:x[3],details:x[4]}));
  const cb=cashBalances({},u);
  const tools=listToolsData_().filter(x=>String(x.clientId||"").trim()===cid);
  const pendingExpenses=0;
  const milestones=listMilestones({clientId:cid},u).milestones||[];
  const overdueMilestones=milestones.filter(x=>x.status==="Overdue").length;
  const overdueTools=tools.filter(x=>x.status==="Overdue").length;
  const usedPct=c.budget?totalExpenses/c.budget*100:0;
  const health=usedPct>=100||Math.max(0,c.budget-totalPayments)>c.budget*.5
    ?"Critical"
    :(usedPct>=75||Math.max(0,c.budget-totalPayments)>c.budget*.3||pendingExpenses||overdueMilestones||overdueTools)
      ?"Attention":"Healthy";

  const alerts=[];
  if(overdueMilestones)alerts.push({level:"critical",title:"Overdue client milestone",detail:overdueMilestones+" payment milestone(s) are overdue."});
  if(overdueTools)alerts.push({level:"attention",title:"Overdue construction tool",detail:overdueTools+" borrowed tool(s) are past their expected return date."});
  if(usedPct>=75)alerts.push({level:usedPct>=100?"critical":"attention",title:"Budget utilization warning",detail:"Project spending has reached "+usedPct.toFixed(1)+"% of the contract budget."});
  if(Math.max(0,c.budget-totalPayments)>c.budget*.3)alerts.push({level:"attention",title:"Uncollected balance",detail:"More than 30% of the contract value remains uncollected."});

  const o={
    ok:true,
    clients:clientObjects(true),
    selectedClientId:cid,
    summary:{clientName:c.name,reference:c.reference,contractAmount:c.budget,totalPayments,totalExpenses,budgetRemaining,outstandingBalance:Math.max(0,c.budget-totalPayments),cashPosition,bankBalance:cb.bankBalance,cashOnHand:cb.cashOnHand,totalCash:cb.totalCash},
    smart:{health,alerts,pendingExpenses,overdueMilestones,overdueTools},
    cashBalances:cb,
    tools:tools,
    payments:p.slice(-100).reverse(),
    expenses:ex.slice(-100).reverse(),
    budget,
    recent,
    activity:av
  };

  if(u.role==="Admin"){
    let uv=rows("users"),fc=uv.filter(x=>String(x[3])==="Finance"&&String(x[4]).toLowerCase()!=="false").length;
    o.users=uv.map(x=>({username:x[0],name:x[1],role:x[3],active:String(x[4]).toLowerCase()!=="false"}));
    o.financeCount=fc;
    o.financeRegistrationOpen=String(settingsMap().RegistrationOpen).toLowerCase()!=="false";
  }
  return o;
}

function getProjectBudgetData_(cid){
  requireClient_(cid,true);
  // IMPORTANT: never fall back to the global "budget" sheet.
  // A project may only read its own project_budgets rows.
  return rows("project_budgets")
    .filter(x=>String(x[0]||"").trim()===String(cid).trim())
    .map(x=>({category:String(x[1]||""),budget:num(x[2])}))
    .filter(x=>x.category);
}
function getProjectBudget(r,u){const c=requireProjectScopedRead_(r);return{ok:true,clientId:c.id,budget:getProjectBudgetData_(c.id)}}
function saveProjectBudget(r,u){
  adminOrFinance_(u);
  const c=requireProjectScopedRead_(r);
  const items=Array.isArray(r.items)?r.items:[];
  if(!items.length)throw Error("Add at least one budget category.");
  const clean=[],seen={};
  items.forEach(x=>{
    const cat=String(x.category||"").trim();
    const b=Number(x.budget);
    if(!cat)return;
    if(!Number.isFinite(b)||b<0)throw Error("Invalid budget for "+cat+".");
    const key=cat.toLowerCase();
    if(seen[key])throw Error("Duplicate budget category: "+cat);
    seen[key]=true;
    clean.push([c.id,cat,b,new Date(),u.name]);
  });
  const sh=ensureSheet_("project_budgets",SHEETS.project_budgets);
  for(let i=sh.getLastRow();i>=2;i--){
    if(String(sh.getRange(i,1).getValue()||"").trim()===c.id)sh.deleteRow(i);
  }
  if(clean.length)sh.getRange(sh.getLastRow()+1,1,clean.length,5).setValues(clean);
  audit("SAVE_PROJECT_BUDGET",u,c.name+" | "+clean.length+" categories");
  return{ok:true,clientId:c.id,budget:clean.map(x=>({category:x[1],budget:x[2]}))};
}

function projectFinancialDashboard(r,u){
  const c=requireProjectScopedRead_(r);
  const cid=c.id;
  const cm=clientNameMap();

  const payments=rows("payments").map((x,i)=>({
    row:i+2,date:x[1],clientId:String(x[2]||"").trim(),clientName:(cm[String(x[2]||"")]||{}).name||"",
    reference:String(x[3]||""),description:String(x[4]||""),dueAmount:num(x[5]),amount:num(x[6]),method:String(x[7]||""),
    notes:String(x[8]||""),enteredBy:String(x[9]||""),username:String(x[10]||"")
  })).filter(x=>x.clientId===cid);

  const expenses=rows("expenses").map((x,i)=>({
    row:i+2,date:x[1],clientId:String(x[2]||"").trim(),clientName:(cm[String(x[2]||"")]||{}).name||"",
    type:String(x[3]||"Expense"),category:String(x[4]||"Other"),payee:String(x[5]||""),description:String(x[6]||""),
    amount:num(x[7]),method:String(x[8]||""),reference:String(x[9]||""),approvedBy:String(x[10]||""),
    notes:String(x[11]||""),enteredBy:String(x[12]||""),username:String(x[13]||""),receiptUrl:String(x[14]||""),receiptId:String(x[15]||""),
    recordStatus:String(x[16]||"Recorded")
  })).filter(x=>x.clientId===cid);

  const approved=expenses;
  const totalPayments=payments.reduce((a,x)=>a+x.amount,0);
  const totalExpenses=approved.filter(x=>x.type==="Expense").reduce((a,x)=>a+x.amount,0);
  const refunds=approved.filter(x=>x.type==="Refund").reduce((a,x)=>a+x.amount,0);
  const income=approved.filter(x=>x.type==="Other Income").reduce((a,x)=>a+x.amount,0);
  const projectCash=totalPayments-totalExpenses+refunds+income;
  const budgetRemaining=Math.max(0,c.budget-totalExpenses);
  const outstanding=Math.max(0,c.budget-totalPayments);
  const collectionPct=c.budget?Math.min(100,totalPayments/c.budget*100):0;
  const usedPct=c.budget?totalExpenses/c.budget*100:0;

  const budget=getProjectBudgetData_(cid).map(x=>{
    const b=num(x.budget);
    const actual=approved.filter(y=>y.type==="Expense"&&String(y.category)===String(x.category)).reduce((a,y)=>a+y.amount,0);
    return{category:String(x.category),budget:b,actual,variance:b-actual,used:b?actual/b:0,remaining:Math.max(0,b-actual)};
  }).filter(x=>x.budget||x.actual);

  const milestones=listMilestones({clientId:cid},u).milestones||[];
  const tools=listToolsData_().filter(x=>String(x.clientId||"").trim()===cid);
  let receipts=[];
  try{receipts=(listReceipts({clientId:cid},u).receipts||[])}catch(e){receipts=[]}

  const overdueTools=tools.filter(x=>x.status==="Overdue").length;
  const overdueMilestones=milestones.filter(x=>x.status==="Overdue").length;
  const pendingExpenses=0;
  const health=(usedPct>=100||outstanding>c.budget*.5)
    ?"Critical"
    :(usedPct>=75||outstanding>c.budget*.3||pendingExpenses||overdueMilestones||overdueTools)
      ?"Attention":"Healthy";

  const alerts=[];
  if(pendingExpenses)alerts.push({level:"critical",title:"Expenses awaiting approval",detail:pendingExpenses+" pending expense transaction(s)."});
  if(overdueMilestones)alerts.push({level:"critical",title:"Overdue payment milestones",detail:overdueMilestones+" milestone(s) need attention."});
  if(overdueTools)alerts.push({level:"attention",title:"Overdue construction tools",detail:overdueTools+" tool(s) are past their expected return date."});
  if(usedPct>=75)alerts.push({level:usedPct>=100?"critical":"attention",title:"Budget utilization",detail:"Approved spending is at "+usedPct.toFixed(1)+"% of contract budget."});
  if(outstanding>c.budget*.3)alerts.push({level:"attention",title:"Client balance outstanding",detail:""+Math.round(100-collectionPct)+"% of the contract remains uncollected."});

  const recent=[...payments.map(x=>({...x,recordType:"Client Payment"})),...expenses.map(x=>({...x,recordType:x.type||"Fund Transaction"}))]
    .sort((a,b)=>new Date(b.date||0)-new Date(a.date||0)).slice(0,15);

  return{
    ok:true,
    project:{id:c.id,name:c.name,reference:c.reference,budget:c.budget,active:c.active},
    summary:{contractAmount:c.budget,totalPayments,totalExpenses,budgetRemaining,outstandingBalance:outstanding,projectCash,cashProfit:projectCash,profitMargin:totalPayments?projectCash/totalPayments*100:0,collectionPct,budgetUsedPct:usedPct},
    counts:{payments:payments.length,expenses:expenses.length,approvedExpenses:approved.filter(x=>x.type==="Expense").length,pendingExpenses,receipts:receipts.length,tools:tools.length,borrowedTools:tools.filter(x=>x.status==="Borrowed"||x.status==="Overdue").length,overdueTools,milestones:milestones.length,overdueMilestones},
    payments:payments.slice().reverse(),expenses:expenses.slice().reverse(),budget,milestones,tools,recent,receipts
  };
}

function saveCalculation(r,u){
  if(!u || !u.username)throw Error("Authentication required.");
  const calculator=String(r.calculator||"").trim();
  if(!calculator)throw Error("Calculator type is required.");

  const allowed=["concrete","steel","block","paint"];
  if(allowed.indexOf(calculator.toLowerCase())<0)throw Error("Unsupported calculator.");

  let clientId="";
  if(r.clientId){
    const c=requireProjectScopedRead_(r);
    clientId=c.id;
  }

  const inputs=typeof r.inputs==="string"?r.inputs:JSON.stringify(r.inputs||{});
  const result=typeof r.result==="string"?r.result:JSON.stringify(r.result||{});
  const sh=ensureSheet_("calculator_history",SHEETS.calculator_history);
  sh.appendRow([new Date(),calculator,inputs,result,clientId,u.name,u.username]);

  audit("CALCULATOR_USED",u,
    (calculator+" calculator"+(clientId?" | Project: "+clientId:"")),
    "Inputs: "+inputs+" | Result: "+result
  );

  return{ok:true,calculator:calculator,clientId:clientId};
}

function listCalculations(r,u){
  if(!u || !u.username)throw Error("Authentication required.");
  const sh=ensureSheet_("calculator_history",SHEETS.calculator_history);
  const values=sh.getDataRange().getValues();
  if(values.length<2)return{ok:true,calculations:[]};

  const requestedClient=String(r.clientId||"").trim();
  const limit=Math.min(Math.max(Number(r.limit)||20,1),100);

  let items=values.slice(1).map(function(row){
    return{
      timestamp:row[0],
      calculator:String(row[1]||""),
      inputs:String(row[2]||""),
      result:String(row[3]||""),
      clientId:String(row[4]||""),
      createdBy:String(row[5]||""),
      username:String(row[6]||"")
    };
  });

  if(requestedClient){
    items=items.filter(function(x){return x.clientId===requestedClient;});
  }

  items=items.slice(-limit).reverse();
  return{ok:true,calculations:items};
}

function addPayment(r,u){
  adminOrFinance_(u);
  const c=requireProjectScopedRead_(r);
  if(num(r.amount)<=0)throw Error("Amount paid must be greater than zero.");
  ss().getSheetByName("payments").appendRow([
    new Date(),r.date||"",c.id,r.reference||"",r.description||"",num(r.dueAmount),num(r.amount),
    r.method||"",r.notes||"",u.name,u.username
  ]);
  audit("ADD_PAYMENT",u,(c.name+" | "+(r.description||"")+" | "+r.amount));
  return{ok:true,clientId:c.id};
}

/* Added without removing/changing the existing transaction workflow:
   these handlers update the existing spreadsheet row instead of appending
   a second transaction, and write an audit entry for every edit. */
function resolvePaymentRow_(r,clientId){
  const sh=ss().getSheetByName("payments");
  if(!sh) throw Error("Sheet not found: payments");
  const last=sh.getLastRow();
  const requestedRow=Number(r&& (r.row||r.id||r.paymentId));
  if(Number.isInteger(requestedRow) && requestedRow>=2 && requestedRow<=last){
    const stored=String(sh.getRange(requestedRow,3).getValue()||"").trim();
    if(stored===String(clientId||"").trim()) return requestedRow;
  }
  const values=last>1?sh.getRange(2,1,last-1,sh.getLastColumn()).getValues():[];
  const wantDate=String(r&& (r.originalDate||r.date)||"").trim();
  const wantRef=String(r&& (r.originalReference||r.reference)||"").trim();
  const wantDesc=String(r&& (r.originalDescription||r.description)||"").trim();
  const wantAmount=num(r&& (r.originalAmount!=null?r.originalAmount:r.amount));
  const matches=[];
  values.forEach(function(x,i){
    if(String(x[2]||"").trim()!==String(clientId||"").trim()) return;
    if((!wantDate||String(x[1]||"").trim()===wantDate) && (!wantRef||String(x[3]||"").trim()===wantRef) && (!wantDesc||String(x[4]||"").trim()===wantDesc) && num(x[6])===wantAmount) matches.push(i+2);
  });
  if(matches.length===1) return matches[0];
  if(matches.length>1 && matches.indexOf(requestedRow)>=0) return requestedRow;
  throw Error("Record not found.");
}

function updatePayment(r,u){
  adminOrFinance_(u);
  const clientId=String(r.clientId||"").trim();
  requireClient_(clientId,true);
  const paymentRow=resolvePaymentRow_(r,clientId);
  const checked=requireRowClient_("payments",paymentRow,clientId,3);
  const sh=checked.sheet, row=checked.row;
  const old=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  if(num(r.amount)<=0)throw Error("Amount paid must be greater than zero.");
  sh.getRange(row,2,1,8).setValues([[
    r.date||old[1], clientId, r.reference||"", r.description||"", num(r.dueAmount),
    num(r.amount), r.method||"", r.notes||""
  ]]);
  audit("EDIT_PAYMENT",u,
    String(old[4]||r.description||"")+" | "+clientId+" | "+num(old[6])+" -> "+num(r.amount)
  );
  return{ok:true,clientId:clientId,row:row,message:"Payment updated."};
}
function editPayment(r,u){ return updatePayment(r,u); }

function addExpense(r,u){
  adminOrFinance_(u);
  const c=requireProjectScopedRead_(r);
  if(num(r.amount)<=0)throw Error("Amount must be greater than zero.");
  const type=String(r.type||"Expense");
  const receipt=saveReceipt_(r.receiptData,r.receiptName,u,c);
  const approvalStatus="Recorded",approvedBy="";
  ss().getSheetByName("expenses").appendRow([
    new Date(),r.date||"",c.id,type,r.category||"Other",r.payee||"",r.description||"",num(r.amount),
    r.method||"",r.reference||"",approvedBy,r.notes||"",u.name,u.username,
    receipt?receipt.url:"",receipt?receipt.id:"",approvalStatus
  ]);
  const detail=(c.name+" | "+(r.category||"")+" | "+(r.description||"")+" | "+r.amount)+(receipt?" | Receipt attached":" | No receipt attached");
  audit("ADD_EXPENSE",u,detail);
  return{ok:true,clientId:c.id,receipt:receipt,recordStatus:"Recorded"};
}

/* Added without removing/changing the existing transaction workflow:
   these handlers update the existing spreadsheet row instead of appending
   a second transaction, and write an audit entry for every edit. */
function updateExpense(r,u){
  adminOrFinance_(u);
  const clientId=String(r.clientId||"").trim();
  requireClient_(clientId,true);
  const checked=requireRowClient_("expenses",r.row||r.id||r.expenseId,clientId,3);
  const sh=checked.sheet, row=checked.row;
  const old=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  if(num(r.amount)<=0)throw Error("Amount must be greater than zero.");
  const type=String(r.type||old[3]||"Expense");
  sh.getRange(row,2,1,11).setValues([[
    r.date||old[1], clientId, type, r.category||old[4]||"Other", r.payee||"",
    r.description||"", num(r.amount), r.method||"", r.reference||"",
    old[10]||"", r.notes||""
  ]]);
  audit("EDIT_EXPENSE",u,
    String(old[6]||r.description||"")+" | "+clientId+" | "+num(old[7])+" -> "+num(r.amount)
  );
  return{ok:true,clientId:clientId,row:row,message:"Fund transaction updated."};
}
function editExpense(r,u){ return updateExpense(r,u); }

function listBills(r,u){
  adminOrFinance_(u);
  const requested=String(r.clientId||"").trim();
  const cm=clientNameMap();
  const items=rows("bills").map(function(x,i){
    const clientId=String(x[2]||"").trim();
    return {row:i+2,timestamp:x[0],id:String(x[1]||""),clientId,clientName:(cm[clientId]||{}).name||"",name:String(x[3]||""),category:String(x[4]||"Other"),dueDate:x[5],amount:num(x[6]),status:String(x[7]||"Pending"),notes:String(x[8]||""),enteredBy:String(x[9]||""),username:String(x[10]||""),updatedAt:x[11]||x[0]};
  });
  const filtered=requested?items.filter(function(x){return x.clientId===requested;}):items;
  return {ok:true,bills:filtered.reverse()};
}

function addBill(r,u){
  adminOrFinance_(u);
  const clientId=String(r.clientId||"").trim();
  if(clientId)requireClient_(clientId,true);
  const name=String(r.name||"").trim();
  if(!name)throw Error("Bill / Payee is required.");
  const amount=num(r.amount);
  if(amount<0)throw Error("Bill amount cannot be negative.");
  const status=String(r.status||"Pending");
  if(["Pending","Ongoing","Paid"].indexOf(status)<0)throw Error("Invalid bill status.");
  const id=Utilities.getUuid(),now=new Date();
  ensureSheet_("bills",SHEETS.bills).appendRow([now,id,clientId,name,String(r.category||"Other"),r.dueDate||"",amount,status,String(r.notes||""),u.name,u.username,now]);
  audit("ADD_BILL",u,(clientId?((clientNameMap()[clientId]||{}).name+" | "):"")+name+" | "+amount+" | "+status);
  return {ok:true,id,clientId,message:"Bill saved."};
}

function updateBill(r,u){
  adminOrFinance_(u);
  const id=String(r.id||r.billId||"").trim();
  if(!id)throw Error("Bill ID is required.");
  const sh=ensureSheet_("bills",SHEETS.bills),v=sh.getDataRange().getValues();
  const i=v.slice(1).findIndex(function(x){return String(x[1]||"")===id;});
  if(i<0)throw Error("Bill not found.");
  const row=i+2,old=v[i+1],clientId=String(r.clientId??old[2]??"").trim();
  if(clientId)requireClient_(clientId,true);
  const name=String(r.name??old[3]??"").trim();
  if(!name)throw Error("Bill / Payee is required.");
  const amount=num(r.amount??old[6]);
  if(amount<0)throw Error("Bill amount cannot be negative.");
  const status=String(r.status??old[7]??"Pending");
  if(["Pending","Ongoing","Paid"].indexOf(status)<0)throw Error("Invalid bill status.");
  const now=new Date();
  sh.getRange(row,2,1,11).setValues([[id,clientId,name,String(r.category??old[4]??"Other"),r.dueDate??old[5]??"",amount,status,String(r.notes??old[8]??""),old[9]||u.name,old[10]||u.username,now]]);
  audit("EDIT_BILL",u,name+" | "+amount+" | "+status);
  return {ok:true,id,clientId,message:"Bill updated."};
}
function editBill(r,u){return updateBill(r,u);}

function deleteBill(r,u){
  adminOrFinance_(u);
  const id=String(r.id||r.billId||"").trim();
  if(!id)throw Error("Bill ID is required.");
  const sh=ensureSheet_("bills",SHEETS.bills),v=sh.getDataRange().getValues();
  const i=v.slice(1).findIndex(function(x){return String(x[1]||"")===id;});
  if(i<0)throw Error("Bill not found.");
  const row=i+2,old=v[i+1],clientId=String(old[2]||"");
  audit("DELETE_BILL",u,(clientId?((clientNameMap()[clientId]||{}).name+" | "):"")+String(old[3]||"")+" | "+num(old[6])+" | "+String(old[7]||""));
  sh.deleteRow(row);
  return {ok:true,id,clientId,message:"Bill deleted."};
}

function listUsers(r,u){adminOnly(u);let set=settingsMap(),uv=rows("users"),fc=uv.filter(x=>String(x[3])==="Finance"&&String(x[4]).toLowerCase()!=="false").length;return{ok:true,users:uv.map(x=>({username:x[0],name:x[1],role:x[3],active:String(x[4]).toLowerCase()!=="false"})),financeCount:fc,financeRegistrationOpen:String(set.RegistrationOpen).toLowerCase()!=="false",clients:clientObjects(false)}}
function upsertUser(r,u){adminOnly(u);if(!r.username||!r.name||!r.password)throw Error("Username, name and password are required.");if(String(r.password).length<8)throw Error("Password must be at least 8 characters.");let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===String(r.username)),now=new Date(),active=String(r.active)!=="FALSE";if(i<0)sh.appendRow([r.username,r.name,hash_(r.password),r.role||"Finance",active,now,now]);else sh.getRange(i+2,1,1,7).setValues([[r.username,r.name,hash_(r.password),r.role||"Finance",active,v[i+1][5]||now,now]]);audit("UPSERT_USER",u,String(r.username));return{ok:true}}
function upsertClient(r,u){if(!r.name)throw Error("Client / Project Name is required.");if(num(r.budget)<0)throw Error("Contract budget cannot be negative.");let sh=ss().getSheetByName("clients"),v=sh.getDataRange().getValues(),id=r.id?String(r.id):Utilities.getUuid(),i=v.slice(1).findIndex(x=>String(x[0])===id),now=new Date(),active=String(r.active)!=="FALSE";if(i<0)sh.appendRow([id,r.name,r.reference||"",num(r.budget),active,now,now]);else sh.getRange(i+2,1,1,7).setValues([[id,r.name,r.reference||"",num(r.budget),active,v[i+1][5]||now,now]]);audit("UPSERT_CLIENT",u,String(r.name)+" | "+r.budget);return{ok:true}}

function archiveClient(r,u){
  if(!r.clientId)throw Error("Client / Project ID is required.");
  let sh=ss().getSheetByName("clients"),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===String(r.clientId));
  if(i<0)throw Error("Client / project not found.");
  sh.getRange(i+2,5).setValue(false);
  sh.getRange(i+2,7).setValue(new Date());
  audit("ARCHIVE_CLIENT",u,String(v[i+1][1]));
  return{ok:true,message:"Client / project archived."};
}
function restoreClient(r,u){
  if(!r.clientId)throw Error("Client / Project ID is required.");
  let sh=ss().getSheetByName("clients"),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===String(r.clientId));
  if(i<0)throw Error("Client / project not found.");
  sh.getRange(i+2,5).setValue(true);
  sh.getRange(i+2,7).setValue(new Date());
  audit("RESTORE_CLIENT",u,String(v[i+1][1]));
  return{ok:true,message:"Client / project restored."};
}
function deleteClient(r,u){
  adminOnly(u);
  if(!r.clientId)throw Error("Client / Project ID is required.");
  let id=String(r.clientId),cm=clientNameMap(),c=cm[id];
  if(!c)throw Error("Client / project not found.");
  let hasPayments=rows("payments").some(x=>String(x[2]||"").trim()===id);
  let hasExpenses=rows("expenses").some(x=>String(x[2]||"").trim()===id);
  let hasTools=rows("tools").some(x=>String(x[12]||"").trim()===id);
  let hasMilestones=rows("milestones").some(x=>String(x[1]||"").trim()===id);
  let hasBudgets=rows("project_budgets").some(x=>String(x[0]||"").trim()===id);
  if(hasPayments||hasExpenses||hasTools||hasMilestones||hasBudgets)
    throw Error("This client/project has linked records. Archive it instead of deleting it.");
  let sh=ss().getSheetByName("clients"),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===id);
  sh.deleteRow(i+2);
  audit("DELETE_CLIENT",u,c.name);
  return{ok:true,message:"Client / project deleted."};
}
function deletePayment(r,u){
  adminOrFinance_(u);
  const clientId=String(r.clientId||"").trim();
  requireClient_(clientId,true);
  const paymentRow=resolvePaymentRow_(r,clientId);
  const checked=requireRowClient_("payments",paymentRow,clientId,3);
  const vals=checked.sheet.getRange(checked.row,1,1,checked.sheet.getLastColumn()).getValues()[0];
  audit("DELETE_PAYMENT",u,clientId+" | "+String(vals[4]||"")+" | "+num(vals[6]));
  checked.sheet.deleteRow(checked.row);
  return{ok:true,clientId,message:"Payment deleted."};
}

function deleteExpense(r,u){
  adminOnly(u);
  const clientId=String(r.clientId||"").trim();
  requireClient_(clientId,true);
  const checked=requireRowClient_("expenses",r.row,clientId,3);
  const vals=checked.sheet.getRange(checked.row,1,1,checked.sheet.getLastColumn()).getValues()[0];
  audit("DELETE_EXPENSE",u,clientId+" | "+String(vals[6]||"")+" | "+num(vals[7]));
  checked.sheet.deleteRow(checked.row);
  return{ok:true,clientId,message:"Expense deleted."};
}

function deactivateUser(r,u){
  adminOnly(u);
  if(!r.username)throw Error("Username is required.");
  if(String(r.username).toLowerCase()===String(u.username).toLowerCase())
    throw Error("You cannot deactivate your own administrator account.");

  let lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
  let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues();
  let i=v.slice(1).findIndex(x=>String(x[0]).toLowerCase()===String(r.username).toLowerCase());
  if(i<0)throw Error("User not found.");

  let row=v[i+1];
  sh.getRange(i+2,5).setValue(false);
  sh.getRange(i+2,7).setValue(new Date());
  audit("DEACTIVATE_USER",u,String(row[0])+" | "+String(row[1]));
  SpreadsheetApp.flush();
  return{ok:true,message:"User deactivated."};
  }finally{lock.releaseLock();}
}


function setUserActive(r,u){
  adminOnly(u);
  if(!r.username)throw Error("Username is required.");
  if(String(r.username).toLowerCase()===String(u.username).toLowerCase())
    throw Error("You cannot change your own administrator account status.");
  let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues();
  let i=v.slice(1).findIndex(x=>String(x[0]).toLowerCase()===String(r.username).toLowerCase());
  if(i<0)throw Error("User not found.");
  let row=v[i+1],active=String(r.active).toLowerCase()==="true";
  sh.getRange(i+2,5).setValue(active);
  sh.getRange(i+2,7).setValue(new Date());
  audit(active?"REACTIVATE_USER":"DEACTIVATE_USER",u,String(row[0])+" | "+String(row[1]));
  return{ok:true};
}

function reactivateUser(r,u){
  adminOnly(u);
  if(!r.username)throw Error("Username is required.");

  let lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
  let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues();
  let i=v.slice(1).findIndex(x=>String(x[0]).toLowerCase()===String(r.username).toLowerCase());
  if(i<0)throw Error("User not found.");

  let row=v[i+1];
  sh.getRange(i+2,5).setValue(true);
  sh.getRange(i+2,7).setValue(new Date());
  audit("REACTIVATE_USER",u,String(row[0])+" | "+String(row[1]));
  SpreadsheetApp.flush();
  return{ok:true,message:"User reactivated."};
  }finally{lock.releaseLock();}
}

function deleteUser(r,u){
  adminOnly(u);
  if(!r.username)throw Error("Username is required.");
  if(String(r.username).toLowerCase()===String(u.username).toLowerCase())
    throw Error("You cannot delete your own administrator account.");

  let lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
  let sh=ss().getSheetByName("users"),v=sh.getDataRange().getValues();
  let i=v.slice(1).findIndex(x=>String(x[0]).toLowerCase()===String(r.username).toLowerCase());
  if(i<0)throw Error("User not found.");

  let row=v[i+1];

  // Keep financial/payment/expense/audit history intact.
  // Only the login-account row is removed.
  sh.deleteRow(i+2);

  audit("DELETE_USER",u,String(row[0])+" | "+String(row[1])+" | Role: "+String(row[3]));
  SpreadsheetApp.flush();
  return{ok:true,message:"User deleted. Existing transaction history was preserved."};
  }finally{lock.releaseLock();}
}



function listToolsData_(){
  return rows("tools").map((x,i)=>({
    row:i+2,
    timestamp:x[0],
    tool:String(x[1]||""),
    toolId:String(x[2]||""),
    borrower:String(x[3]||""),
    project:String(x[4]||""),
    borrowed:String(x[5]||""),
    expectedReturn:String(x[6]||""),
    returned:String(x[7]||""),
    status:(String(x[8]||"Borrowed")==="Borrowed"&&x[6]&&new Date(x[6])<new Date()&& !x[7])?"Overdue":String(x[8]||"Borrowed"),
    notes:String(x[9]||""),
    enteredBy:String(x[10]||""),
    username:String(x[11]||""),
    clientId:String(x[12]||"")
  }));
}

function listTools(r,u){
  const c=requireProjectScopedRead_(r);
  const tools=listToolsData_().filter(x=>String(x.clientId||"").trim()===c.id);
  return{ok:true,clientId:c.id,tools};
}

function addTool(r,u){
  adminOrFinance_(u);
  const c=requireProjectScopedRead_(r);
  if(!r.tool)throw Error("Tool / Equipment name is required.");
  if(!r.borrower)throw Error("Borrowed By is required.");

  ensureSheet_("tools",SHEETS.tools).appendRow([
    new Date(),String(r.tool),String(r.toolId||""),String(r.borrower),c.name,String(r.borrowed||""),
    String(r.expectedReturn||""),String(r.returned||""),String(r.status||"Borrowed"),String(r.notes||""),
    u.name,u.username,c.id
  ]);
  audit("ADD_TOOL",u,c.name+" | "+String(r.tool)+" | Borrowed by: "+String(r.borrower));
  return{ok:true,clientId:c.id};
}

function updateTool(r,u){
  adminOrFinance_(u);
  const row=Number(r.row);
  if(!row||row<2)throw Error("Invalid tool record.");
  const sh=ss().getSheetByName("tools");
  if(row>sh.getLastRow())throw Error("Tool record not found.");

  const old=sh.getRange(row,1,1,13).getValues()[0];
  const c=assertToolProject_(r,old);

  sh.getRange(row,1,1,13).setValues([[
    old[0],
    String(r.tool||old[1]),
    String(r.toolId||old[2]),
    String(r.borrower||old[3]),
    c.name,
    String(r.borrowed||old[5]),
    String(r.expectedReturn||old[6]),
    String(r.returned||old[7]),
    String(r.status||old[8]),
    String(r.notes||old[9]),
    u.name,u.username,c.id
  ]]);

  audit("UPDATE_TOOL",u,c.name+" | "+String(r.tool||old[1])+" | Status: "+String(r.status||old[8]));
  return{ok:true,clientId:c.id};
}

function cashBalances(r,u){
  let sh=ss().getSheetByName("cash_balances");
  let v=sh.getDataRange().getValues();
  let map={bank:0,onhand:0};
  v.slice(1).forEach(x=>{
    let k=String(x[0]||"").toLowerCase().replace(/\s+/g,"");
    if(k==="bank")map.bank=num(x[1]);
    if(k==="onhand"||k==="cashonhand")map.onhand=num(x[1]);
  });
  return{
    ok:true,
    bankBalance:map.bank,
    cashOnHand:map.onhand,
    totalCash:map.bank+map.onhand
  };
}

function updateCashBalance(r,u){
  adminOrFinance_(u);
  let account=String(r.account||"").toLowerCase().replace(/\s+/g,"");
  if(account!=="bank"&&account!=="onhand")throw Error("Account must be Bank or On Hand.");
  let balance=num(r.balance);
  if(balance<0)throw Error("Cash balance cannot be negative.");

  let sh=ss().getSheetByName("cash_balances");
  let v=sh.getDataRange().getValues();
  let i=v.slice(1).findIndex(x=>String(x[0]||"").toLowerCase().replace(/\s+/g,"")===account);
  let display=account==="bank"?"Bank":"On Hand";
  let notes=String(r.notes||"");

  if(i<0){
    sh.appendRow([display,balance,new Date(),u.name,notes]);
  }else{
    sh.getRange(i+2,1,1,5).setValues([[display,balance,new Date(),u.name,notes]]);
  }

  audit("UPDATE_CASH_BALANCE",u,display+" | "+balance);
  return cashBalances(r,u);
}

function updateCashBalances(r,u){
  adminOrFinance_(u);
  let bank=Number(r.bankBalance),onHand=Number(r.cashOnHand);
  if(!Number.isFinite(bank)||bank<0)throw Error("Enter a valid bank balance.");
  if(!Number.isFinite(onHand)||onHand<0)throw Error("Enter a valid cash-on-hand balance.");
  let lock=LockService.getScriptLock();
  lock.waitLock(10000);
  try{
    let sh=ss().getSheetByName("cash_balances"),v=sh.getDataRange().getValues(),now=new Date();
    function save_(account,label,balance){
      let row=v.slice(1).findIndex(x=>String(x[0]||"").toLowerCase().replace(/\s+/g,"")===account);
      if(row<0)sh.appendRow([label,balance,now,u.name,""]);
      else sh.getRange(row+2,1,1,5).setValues([[label,balance,now,u.name,""]]);
    }
    save_("bank","Bank",bank);
    save_("onhand","On Hand",onHand);
    audit("UPDATE_CASH_BALANCES",u,"Bank: "+bank+" | On Hand: "+onHand);
    SpreadsheetApp.flush();
    return cashBalances({},u);
  }finally{lock.releaseLock();}
}

function changeInvite(r,u){adminOnly(u);if(!r.inviteCode||String(r.inviteCode).length<8)throw Error("Invitation code must be at least 8 characters.");setSetting_("InviteCodeHash",hash_(r.inviteCode));setSetting_("RegistrationOpen",true);audit("CHANGE_INVITE",u,"Invitation code changed and registration reopened");return{ok:true}}
function reopenRegistration(r,u){adminOnly(u);let uv=rows("users"),fc=uv.filter(x=>String(x[3])==="Finance"&&String(x[4]).toLowerCase()!=="false").length;if(fc>=3)throw Error("There are already 3 active LIWO Executive accounts. Deactivate a Finance account first.");setSetting_("RegistrationOpen",true);audit("REOPEN_REGISTRATION",u,"LIWO Executive registration reopened");return{ok:true}}
function setSetting_(key,value){let sh=ss().getSheetByName("settings"),v=sh.getDataRange().getValues(),i=v.findIndex(x=>x[0]===key);if(i<0)sh.appendRow([key,value]);else sh.getRange(i+1,2).setValue(value)}

function adminOrFinance_(u){if(!u || (u.role!=="Admin" && u.role!=="Finance"))throw Error("Authorized LIWO Finance users only.");}
function listPendingExpenses(r,u){
  const c=requireProjectScopedRead_(r);
  const cm=clientNameMap(),out=[];
  rows("expenses").forEach(function(x,i){
    const clientId=String(x[2]||"").trim();
    if(clientId!==c.id)return;
    const status=String(x[16]||((String(x[3])==="Expense")?"Approved":"Not Required"));
    if(String(x[3])==="Expense"&&status==="Pending")
      out.push({row:i+2,date:x[1],clientId,clientName:(cm[clientId]||{}).name||"",category:x[4],payee:x[5],description:x[6],amount:num(x[7]),enteredBy:x[12],username:x[13],receiptUrl:x[14]||""});
  });
  return{ok:true,clientId:c.id,expenses:out.reverse()};
}

function setExpenseApproval_(r,u,status){
  adminOnly(u);
  const clientId=String(r.clientId||"").trim();
  requireClient_(clientId,true);
  const checked=requireRowClient_("expenses",r.row,clientId,3);
  const old=checked.sheet.getRange(checked.row,1,1,17).getValues()[0];
  if(String(old[3])!=="Expense")throw Error("Only Expense transactions require approval.");
  checked.sheet.getRange(checked.row,11).setValue(status==="Approved"?u.name:"");
  checked.sheet.getRange(checked.row,17).setValue(status);
  audit(status==="Approved"?"APPROVE_EXPENSE":"REJECT_EXPENSE",u,clientId+" | "+String(old[6]||old[5]||"")+" | "+old[7]);
  return{ok:true,clientId,status};
}
function approveExpense(r,u){return setExpenseApproval_(r,u,"Approved")}
function rejectExpense(r,u){return setExpenseApproval_(r,u,"Rejected")}

function listMilestones(r,u){
  const c=requireProjectScopedRead_(r);
  const cm=clientNameMap();
  const items=rows("milestones")
    .map((x,i)=>({row:i+2,id:String(x[0]),clientId:String(x[1]||"").trim(),clientName:(cm[String(x[1]||"")]||{}).name||"",milestone:String(x[2]||""),dueDate:x[3],amount:num(x[4]),status:String(x[5]||"Upcoming"),paidAmount:num(x[6]),notes:String(x[7]||""),createdAt:x[8],updatedAt:x[9],createdBy:String(x[10]||"")}))
    .filter(x=>x.clientId===c.id);
  items.forEach(x=>{
    if(x.status!=="Paid"&&x.dueDate&&new Date(x.dueDate)<new Date()&&x.status!=="Cancelled")x.status="Overdue";
  });
  return{ok:true,clientId:c.id,milestones:items};
}

function addMilestone(r,u){
  adminOrFinance_(u);
  const c=requireProjectScopedRead_(r);
  if(!r.milestone)throw Error("Client/project and milestone are required.");
  if(num(r.amount)<0)throw Error("Milestone amount cannot be negative.");
  const id=Utilities.getUuid(),now=new Date();
  ensureSheet_("milestones",SHEETS.milestones).appendRow([
    id,c.id,String(r.milestone),r.dueDate||"",num(r.amount),String(r.status||"Upcoming"),
    num(r.paidAmount),String(r.notes||""),now,now,u.name
  ]);
  audit("ADD_MILESTONE",u,c.name+" | "+String(r.milestone)+" | "+r.amount);
  return{ok:true,id,clientId:c.id};
}

function updateMilestone(r,u){
  adminOrFinance_(u);
  const clientId=String(r.clientId||"").trim();
  const c=requireClient_(clientId,true);
  const row=Number(r.row);
  if(!row||row<2)throw Error("Invalid milestone record.");
  const sh=ensureSheet_("milestones",SHEETS.milestones);
  if(row>sh.getLastRow())throw Error("Milestone not found.");
  const old=sh.getRange(row,1,1,11).getValues()[0];
  if(String(old[1]||"").trim()!==c.id)
    throw Error("Project scope violation: this milestone belongs to another project.");
  const now=new Date();
  sh.getRange(row,1,1,11).setValues([[
    old[0],c.id,String(r.milestone||old[2]),r.dueDate||old[3],num(r.amount??old[4]),
    String(r.status||old[5]),num(r.paidAmount??old[6]),String(r.notes??old[7]),old[8],now,u.name
  ]]);
  audit("UPDATE_MILESTONE",u,c.name+" | "+String(r.milestone||old[2]));
  return{ok:true,clientId:c.id};
}

function verifyReceipt(r,u){adminOrFinance_(u);let id=String(r.receiptId||"");if(!id)throw Error("Receipt ID is required.");let sh=ensureSheet_("receipt_verification",SHEETS.receipt_verification),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===id),now=new Date(),status=String(r.status||"Verified");if(!["Pending","Verified","Rejected"].includes(status))throw Error("Invalid receipt status.");let row=[id,status,u.name,now,String(r.notes||"")];if(i<0)sh.appendRow(row);else sh.getRange(i+2,1,1,5).setValues([row]);audit("VERIFY_RECEIPT",u,status+" | "+id);return{ok:true,status};}
function receiptVerificationMap_(){let m={};rows("receipt_verification").forEach(x=>m[String(x[0])]={status:String(x[1]||"Pending"),verifiedBy:String(x[2]||""),verifiedAt:x[3],notes:String(x[4]||"")});return m;}
function reconcileCash(r,u){adminOrFinance_(u);let expected=cashBalances({},u),ab=Number(r.actualBank),ao=Number(r.actualCashOnHand);if(!Number.isFinite(ab)||ab<0||!Number.isFinite(ao)||ao<0)throw Error("Enter valid actual bank and cash-on-hand balances.");let variance=(ab+ao)-expected.totalCash,sh=ensureSheet_("cash_reconciliation",SHEETS.cash_reconciliation),now=new Date();sh.appendRow([now,expected.bankBalance,expected.cashOnHand,ab,ao,variance,String(r.notes||""),u.name]);audit("RECONCILE_CASH",u,"Variance: "+variance);return{ok:true,expectedBank:expected.bankBalance,expectedCashOnHand:expected.cashOnHand,actualBank:ab,actualCashOnHand:ao,variance};}
function returnTool(r,u){
  adminOrFinance_(u);
  const row=Number(r.row);
  if(!row||row<2)throw Error("Invalid tool record.");
  const sh=ss().getSheetByName("tools");
  if(row>sh.getLastRow())throw Error("Tool record not found.");
  const old=sh.getRange(row,1,1,13).getValues()[0];
  const c=assertToolProject_(r,old);
  const returned=r.returned||Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyy-MM-dd");
  sh.getRange(row,5).setValue(c.name);
  sh.getRange(row,8).setValue(returned);
  sh.getRange(row,9).setValue("Returned");
  sh.getRange(row,10).setValue(String(r.notes||old[9]||""));
  sh.getRange(row,11).setValue(u.name);
  sh.getRange(row,12).setValue(u.username);
  sh.getRange(row,13).setValue(c.id);
  audit("RETURN_TOOL",u,c.name+" | "+String(old[1])+" | Returned: "+returned);
  return{ok:true,clientId:c.id};
}

function audit(a,u,d){ss().getSheetByName("audit").appendRow([new Date(),a,u.username,u.name,d])}

/* ===================== LIWO EXECUTIVE + AUTOMATED REPORTING ===================== */
function executiveDashboard(r,u){
  let clients=clientObjects(true),cm=clientNameMap();
  let payments=rows("payments"),expenses=rows("expenses");
  let projects=clients.map(c=>{
    let p=payments.filter(x=>String(x[2])===String(c.id));
    let e=expenses.filter(x=>String(x[2])===String(c.id));
    let totalPayments=p.reduce((a,x)=>a+num(x[6]),0);
    let totalExpenses=e.filter(x=>String(x[3]||"Expense")==="Expense").reduce((a,x)=>a+num(x[7]),0);
    let refunds=e.filter(x=>String(x[3])==="Refund").reduce((a,x)=>a+num(x[7]),0);
    let income=e.filter(x=>String(x[3])==="Other Income").reduce((a,x)=>a+num(x[7]),0);
    let cashProfit=totalPayments+income+refunds-totalExpenses;
    let margin=totalPayments?cashProfit/totalPayments*100:0;
    let used=c.budget?totalExpenses/c.budget*100:0;
    let outstanding=Math.max(0,c.budget-totalPayments);
    let status=used>=100||outstanding>c.budget*.5?"Critical":(used>=75||outstanding>c.budget*.3)?"Attention":"Healthy";
    return {id:c.id,name:c.name,reference:c.reference,budget:c.budget,totalPayments,totalExpenses,otherIncome:income,refunds,cashProfit,margin,usedPct:used,collectionPct:c.budget?Math.min(100,totalPayments/c.budget*100):0,outstanding,status};
  });
  let totals=projects.reduce((a,p)=>{a.contract+=p.budget;a.payments+=p.totalPayments;a.expenses+=p.totalExpenses;a.profit+=p.cashProfit;a.outstanding+=p.outstanding;return a},{contract:0,payments:0,expenses:0,profit:0,outstanding:0});
  totals.margin=totals.payments?totals.profit/totals.payments*100:0;
  return {ok:true,generatedAt:new Date(),totals,projects,activeProjects:projects.length,healthy:projects.filter(p=>p.status==="Healthy").length,attention:projects.filter(p=>p.status==="Attention").length,critical:projects.filter(p=>p.status==="Critical").length};
}
function getReportSettings(r,u){if(u.role!=="Admin")throw Error("Administrator access required.");let s=settingsMap();return{ok:true,recipients:String(s.ReportRecipients||"")};}
function saveReportSettings(r,u){adminOnly(u);let emails=String(r.recipients||"").split(/[;,\s]+/).map(x=>x.trim()).filter(Boolean);emails.forEach(e=>{if(!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e))throw Error("Invalid report email: "+e)});setSetting_("ReportRecipients",emails.join(","));audit("SAVE_REPORT_SETTINGS",u,emails.join(", "));return{ok:true,recipients:emails.join(",")};}
function reportDataForPeriod_(start,end){
  let clients=clientObjects(true),cm=clientNameMap(),payments=rows("payments"),expenses=rows("expenses");
  let projects=clients.map(c=>{
    let p=payments.filter(x=>String(x[2])===String(c.id)&&(!start||new Date(x[1])>=start)&&(!end||new Date(x[1])<end));
    let e=expenses.filter(x=>String(x[2])===String(c.id)&&(!start||new Date(x[1])>=start)&&(!end||new Date(x[1])<end));
    let pay=p.reduce((a,x)=>a+num(x[6]),0),cost=e.filter(x=>String(x[3]||"Expense")==="Expense").reduce((a,x)=>a+num(x[7]),0),income=e.filter(x=>String(x[3])==="Other Income").reduce((a,x)=>a+num(x[7]),0),refund=e.filter(x=>String(x[3])==="Refund").reduce((a,x)=>a+num(x[7]),0),profit=pay+income+refund-cost;
    return{name:c.name,reference:c.reference,budget:c.budget,payments:pay,expenses:cost,profit,margin:pay?profit/pay*100:0,outstanding:Math.max(0,c.budget-pay)};
  });
  return projects;
}
function generateFinancialReport(r,u){
  adminOnly(u);
  let period=String(r.period||"month"),now=new Date(),start=null,end=null,label="Complete Financial Report";
  if(period==="month"){start=new Date(now.getFullYear(),now.getMonth(),1);end=new Date(now.getFullYear(),now.getMonth()+1,1);label=Utilities.formatDate(start,Session.getScriptTimeZone(),"MMMM yyyy")+" Financial Report";}
  if(period==="week"){let d=new Date(now);let day=d.getDay();let diff=day===0?-6:1-day;start=new Date(d);start.setHours(0,0,0,0);start.setDate(d.getDate()+diff);end=new Date(start);end.setDate(start.getDate()+7);label="Weekly Financial Report";}
  let projects=reportDataForPeriod_(start,end),tot=projects.reduce((a,p)=>{a.budget+=p.budget;a.payments+=p.payments;a.expenses+=p.expenses;a.profit+=p.profit;a.outstanding+=p.outstanding;return a},{budget:0,payments:0,expenses:0,profit:0,outstanding:0});
  tot.margin=tot.payments?tot.profit/tot.payments*100:0;
  let html='<html><head><style>body{font-family:Arial,sans-serif;color:#16324f}h1{margin-bottom:4px}table{border-collapse:collapse;width:100%;margin-top:20px}th,td{border:1px solid #d8e1e7;padding:8px;font-size:11px}th{background:#eef5f8;text-align:left}.num{text-align:right}.summary{display:grid;grid-template-columns:repeat(5,1fr);gap:8px}.box{border:1px solid #d8e1e7;padding:10px}.box small{display:block;color:#667} </style></head><body><h1>LIWO ENGINEERING CONSULTANCY</h1><h2>'+label+'</h2><p>Generated '+Utilities.formatDate(now,Session.getScriptTimeZone(),"MMM d, yyyy HH:mm")+'</p><div class="summary">'+[['Contract Budget',tot.budget],['Client Payments',tot.payments],['Expenses',tot.expenses],['Cash Profit',tot.profit],['Uncollected',tot.outstanding]].map(x=>'<div class="box"><small>'+x[0]+'</small><b>₱'+Number(x[1]).toLocaleString('en-PH',{minimumFractionDigits:2})+'</b></div>').join('')+'</div><table><tr><th>Project</th><th>Reference</th><th>Budget</th><th>Payments</th><th>Expenses</th><th>Cash Profit</th><th>Margin</th><th>Uncollected</th></tr>'+projects.map(p=>'<tr><td>'+escReport_(p.name)+'</td><td>'+escReport_(p.reference)+'</td><td class="num">₱'+p.budget.toLocaleString('en-PH',{minimumFractionDigits:2})+'</td><td class="num">₱'+p.payments.toLocaleString('en-PH',{minimumFractionDigits:2})+'</td><td class="num">₱'+p.expenses.toLocaleString('en-PH',{minimumFractionDigits:2})+'</td><td class="num">₱'+p.profit.toLocaleString('en-PH',{minimumFractionDigits:2})+'</td><td class="num">'+p.margin.toFixed(1)+'%</td><td class="num">₱'+p.outstanding.toLocaleString('en-PH',{minimumFractionDigits:2})+'</td></tr>').join('')+'</table></body></html>';
  let folderIt=DriveApp.getFoldersByName(CONFIG.REPORT_FOLDER_NAME),folder=folderIt.hasNext()?folderIt.next():DriveApp.createFolder(CONFIG.REPORT_FOLDER_NAME);
  let doc=DocumentApp.create('LIWO '+label+' '+Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy-MM-dd'));let body=doc.getBody();body.appendParagraph('LIWO ENGINEERING CONSULTANCY').setHeading(DocumentApp.ParagraphHeading.HEADING1);body.appendParagraph(label).setHeading(DocumentApp.ParagraphHeading.HEADING2);body.appendParagraph('Generated '+Utilities.formatDate(now,Session.getScriptTimeZone(),'MMM d, yyyy HH:mm'));body.appendParagraph('Contract Budget: ₱'+tot.budget.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Payments: ₱'+tot.payments.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Expenses: ₱'+tot.expenses.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Cash Profit: ₱'+tot.profit.toLocaleString('en-PH',{minimumFractionDigits:2}));projects.forEach(p=>body.appendParagraph(p.name+' ('+p.reference+') — Budget ₱'+p.budget.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Payments ₱'+p.payments.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Expenses ₱'+p.expenses.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Profit ₱'+p.profit.toLocaleString('en-PH',{minimumFractionDigits:2})+' | Margin '+p.margin.toFixed(1)+'%'));doc.saveAndClose();let docFile=DriveApp.getFileById(doc.getId());let pdf=folder.createFile(docFile.getAs(MimeType.PDF)).setName(docFile.getName()+'.pdf');docFile.setTrashed(true);ensureSheet_("report_runs",SHEETS.report_runs).appendRow([now,period,label,pdf.getId(),pdf.getUrl(),u.name]);audit("GENERATE_REPORT",u,label);return{ok:true,fileId:pdf.getId(),fileUrl:pdf.getUrl(),label};
}
function escReport_(s){return String(s||"").replace(/[&<>"']/g,m=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':'&quot;',"'":"&#039;"}[m]));}
function sendMonthlyFinancialReport(){
  const recipients=getReportRecipients_();
  if(!recipients.length) throw Error("No automated report recipients configured.");
  const tz=Session.getScriptTimeZone()||"Asia/Manila";
  const now=new Date();
  const startDate=new Date(now.getFullYear(),now.getMonth()-1,1);
  const endDate=new Date(now.getFullYear(),now.getMonth(),0);
  const start=Utilities.formatDate(startDate,tz,"MMMM d, yyyy");
  const end=Utilities.formatDate(endDate,tz,"MMMM d, yyyy");
  const label="Monthly Financial Report: "+start+" – "+end;
  const rows=getProjectFinancialSummary_(startDate,endDate);
  const html=buildFinancialReportHtml_(label,rows,"Monthly");
  const folder=getOrCreateReceiptFolder_();
  const base="LIWO Monthly Financial Report "+Utilities.formatDate(startDate,tz,"yyyy-MM");
  const file=folder.createFile(Utilities.newBlob(html,"text/html",base+".html"));
  const pdfFile=folder.createFile(file.getAs(MimeType.PDF).setName(base+".pdf"));
  try{file.setTrashed(true)}catch(_){}
  MailApp.sendEmail({
    to:recipients.join(","),
    subject:"LIWO Finance — "+label,
    htmlBody:buildAutomatedReportEmailHtml_("Monthly",label),
    attachments:[pdfFile.getBlob()]
  });
  audit_({username:"SYSTEM",name:"SYSTEM",role:"Admin"},"MONTHLY_FINANCIAL_REPORT","SYSTEM",label,JSON.stringify({pdfUrl:pdfFile.getUrl(),recipients:recipients}));
  return {ok:true,period:label,pdfUrl:pdfFile.getUrl(),recipients:recipients,projects:rows};
}
function installMonthlyReportTrigger(r,u){adminOnly(u);ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='sendMonthlyFinancialReport').forEach(t=>ScriptApp.deleteTrigger(t));ScriptApp.newTrigger('sendMonthlyFinancialReport').timeBased().onMonthDay(1).atHour(8).create();audit('INSTALL_MONTHLY_REPORT',u,'Monthly automated report trigger installed');return{ok:true};}
