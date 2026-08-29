const CONFIG = {
  SPREADSHEET_ID: "19G1z2Yl56XwQpQxPNLzV4JUhxWy7f_hXohS8wDyeIZ8",
  PROJECT_NAME: "LIWO FINANCE TRACKER",
  SESSION_HOURS: 12,

  BOOTSTRAP_ADMIN_USERNAME: "admin",
  BOOTSTRAP_ADMIN_NAME: "LIWO Administrator",
  BOOTSTRAP_ADMIN_PASSWORD: "Althea0610",

  DEFAULT_INVITE_CODE: "liwoecfinance",

  RECEIPT_FOLDER_NAME: "LIWO Finance Receipts",
  MAX_RECEIPT_BYTES: 5 * 1024 * 1024
};
const SHEETS={
 settings:["Key","Value"],
 users:["Username","Name","PasswordHash","Role","Active","CreatedAt","UpdatedAt"],
 clients:["ClientID","Client / Project Name","Reference","Contract Budget","Active","CreatedAt","UpdatedAt"],
 payments:["Timestamp","Date","ClientID","Payment Ref.","Description / Milestone","Due Amount","Amount Paid","Payment Method","Notes","Entered By","Username"],
 expenses:["Timestamp","Date","ClientID","Type","Category","Payee / Supplier","Description","Amount","Payment Method","Receipt / Ref.","Approved By","Notes","Entered By","Username","Receipt File URL","Receipt File ID"],
 budget:["Category","Budget"],tools:["Timestamp","Tool","Tool ID","Borrowed By","Project","Date Borrowed","Expected Return","Date Returned","Status","Notes","Entered By","Username"],cash_balances:["Account","Balance","UpdatedAt","UpdatedBy","Notes"],audit:["Timestamp","Action","Username","Name","Details"]
};
function setup(){
 const s=ss();Object.keys(SHEETS).forEach(n=>{let sh=s.getSheetByName(n)||s.insertSheet(n);ensureHeaders_(sh,SHEETS[n]);sh.setFrozenRows(1);sh.getRange(1,1,1,SHEETS[n].length).setFontWeight("bold")});
 let sm=settingsMap(),st=s.getSheetByName("settings");
 if(sm.ProjectName===undefined)st.appendRow(["ProjectName",CONFIG.PROJECT_NAME]);
 if(sm.InviteCodeHash===undefined){if(CONFIG.DEFAULT_INVITE_CODE==="CHANGE_THIS_INVITE_CODE")throw Error("Change DEFAULT_INVITE_CODE before setup().");st.appendRow(["InviteCodeHash",hash_(CONFIG.DEFAULT_INVITE_CODE)])}
 if(sm.RegistrationOpen===undefined)st.appendRow(["RegistrationOpen",true]);
 let u=s.getSheetByName("users");if(u.getLastRow()===1){if(CONFIG.BOOTSTRAP_ADMIN_PASSWORD==="CHANGE_THIS_BEFORE_SETUP")throw Error("Change BOOTSTRAP_ADMIN_PASSWORD before setup().");u.appendRow([CONFIG.BOOTSTRAP_ADMIN_USERNAME,CONFIG.BOOTSTRAP_ADMIN_NAME,hash_(CONFIG.BOOTSTRAP_ADMIN_PASSWORD),"Admin",true,new Date(),new Date()])}
 return"Setup complete";
}
function doGet(){return json({ok:true,service:"LIWO Finance Tracker"})}
function doPost(e){try{let r=JSON.parse(e.postData.contents||"{}");switch(r.action){case"health":return json({ok:true,service:"LIWO Finance Tracker",version:"2026-08-29"});
case"login":return json(login(r));case"registerFinance":return json(registerFinance(r));case"dashboard":return json(withAuth(r,dashboard));
case"addPayment":return json(withAuth(r,addPayment));case"addExpense":return json(withAuth(r,addExpense));case"listUsers":return json(withAuth(r,listUsers));
case"upsertUser":return json(withAuth(r,upsertUser));case"listTools":return json(withAuth(r,listTools));case"tools":return json(withAuth(r,listTools));case"constructionTools":return json(withAuth(r,listTools));case"listConstructionTools":return json(withAuth(r,listTools));case"getTools":return json(withAuth(r,listTools));case"addTool":return json(withAuth(r,addTool));case"updateTool":return json(withAuth(r,updateTool));case"cashBalances":return json(withAuth(r,cashBalances));case"cashPosition":return json(withAuth(r,cashBalances));case"getCashPosition":return json(withAuth(r,cashBalances));case"getCashBalances":return json(withAuth(r,cashBalances));case"updateCashBalance":return json(withAuth(r,updateCashBalance));case"updateCashBalances":return json(withAuth(r,updateCashBalances));case"changeInvite":return json(withAuth(r,changeInvite));case"reopenRegistration":return json(withAuth(r,reopenRegistration));case"upsertClient":return json(withAuth(r,upsertClient));case"archiveClient":return json(withAuth(r,archiveClient));case"restoreClient":return json(withAuth(r,restoreClient));case"deleteClient":return json(withAuth(r,deleteClient));case"notifications":return json(withAuth(r,notifications));case"deletePayment":return json(withAuth(r,deletePayment));case"deleteExpense":return json(withAuth(r,deleteExpense));
case"deleteUser":return json(withAuth(r,deleteUser));case"deactivateUser":return json(withAuth(r,deactivateUser));case"reactivateUser":return json(withAuth(r,reactivateUser));case"setUserActive":return json(withAuth(r,setUserActive));case"listReceipts":return json(withAuth(r,listReceipts));case"listReceipt":return json(withAuth(r,listReceipts));case"getReceipts":return json(withAuth(r,listReceipts));case"getReceiptList":return json(withAuth(r,listReceipts));case"receipts":return json(withAuth(r,listReceipts));case"receiptGallery":return json(withAuth(r,listReceipts));case"getReceiptGallery":return json(withAuth(r,listReceipts));case"loadReceipts":return json(withAuth(r,listReceipts));case"getReceiptLibrary":return json(withAuth(r,listReceipts));default:return json({ok:false,error:"Unknown action: "+String(r.action||"")})}}catch(x){return json({ok:false,error:String(x.message||x)})}}
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
  let root=receiptRootFolder_();
  let folderName=safeFolderName_(client&&client.name?client.name:"Unassigned");
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
function rows(n){let v=ss().getSheetByName(n).getDataRange().getValues();return v.length>1?v.slice(1):[]}
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
  let rootIt=DriveApp.getFoldersByName(CONFIG.RECEIPT_FOLDER_NAME);
  if(!rootIt.hasNext()){
    let root=DriveApp.createFolder(CONFIG.RECEIPT_FOLDER_NAME);
    return{ok:true,receipts:[],projects:[]};
  }

  let root=rootIt.next();

  // Map stored Drive file IDs to their project/client from the spreadsheet.
  let fileProjectMap={};
  try{
    let cm=clientNameMap();
    rows("expenses").forEach(x=>{
      let fileId=String(x[15]||"").trim();
      let clientId=String(x[2]||"").trim();
      if(fileId && cm[clientId])fileProjectMap[fileId]=cm[clientId].name;
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

    let mappedProject=fileProjectMap[id]||folderProject||"Unassigned";
    if(project && safeFolderName_(mappedProject)!==safeFolderName_(project) && mappedProject!==project)return;

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

  receipts.sort((a,b)=>new Date(b.createdAt)-new Date(a.createdAt));
  return{ok:true,receipts:receipts,projects:projects};
}
function notifications(r,u){
  let limit=Math.min(Math.max(Number(r.limit)||100,1),100);
  let v=ss().getSheetByName("audit").getDataRange().getValues();
  let items=v.slice(1).filter(x=>String(x[1]||"")!=="LOGIN").slice(-limit).reverse().map(x=>({
    time:x[0],action:String(x[1]||"Finance activity"),username:String(x[2]||""),
    name:String(x[3]||""),details:String(x[4]||"")
  }));
  return{ok:true,notifications:items};
}

function clientObjects(activeOnly){return rows("clients").filter(x=>!activeOnly||String(x[4]).toLowerCase()!=="false").map(x=>({id:String(x[0]),name:String(x[1]),reference:String(x[2]||""),budget:num(x[3]),active:String(x[4]).toLowerCase()!=="false"}))}
function clientNameMap(){let m={};clientObjects(false).forEach(c=>m[c.id]=c);return m}
function dashboard(r,u){
 let clients=clientObjects(true),cid=String(r.clientId||"");if(!cid&&clients.length)cid=clients[0].id;let cm=clientNameMap(),c=cm[cid]||{id:cid,name:"",reference:"",budget:0};
 let p=rows("payments").filter(x=>String(x[2])===cid).map(x=>({date:x[1],clientName:c.name,reference:x[3],description:x[4],amount:num(x[6]),method:x[7],enteredBy:x[9]}));
 let ex=rows("expenses").filter(x=>String(x[2])===cid).map(x=>({date:x[1],clientName:c.name,clientId:x[2],type:x[3],category:x[4],payee:x[5],description:x[6],amount:num(x[7]),method:x[8],reference:x[9],enteredBy:x[12],username:x[13],receiptUrl:x[14]||"",receiptId:x[15]||""}));
 let totalPayments=p.reduce((a,x)=>a+x.amount,0),totalExpenses=ex.filter(x=>x.type==="Expense").reduce((a,x)=>a+x.amount,0),refunds=ex.filter(x=>x.type==="Refund").reduce((a,x)=>a+x.amount,0),income=ex.filter(x=>x.type==="Other Income").reduce((a,x)=>a+x.amount,0);
 let budgetRemaining=Math.max(0,c.budget-totalExpenses),cashPosition=totalPayments-totalExpenses+refunds+income;
 let budget=rows("budget").map(x=>{let actual=ex.filter(y=>y.category===x[0]&&y.type==="Expense").reduce((a,y)=>a+y.amount,0),b=num(x[1]);return{category:x[0],budget:b,actual,variance:b-actual,used:b?actual/b:0,remaining:Math.max(0,b-actual)}}).filter(x=>x.budget||x.actual);
 let recent=[...p.map(x=>({...x,type:"Client Payment"})),...ex.map(x=>({...x,type:x.type}))].sort((a,b)=>String(b.date).localeCompare(String(a.date))).slice(0,12);
 let av=rows("audit").slice(-100).reverse().map(x=>({time:x[0],action:x[1],username:x[2],name:x[3],details:x[4]}));
 let cb=cashBalances({},u);let tools=listToolsData_();let o={ok:true,clients,summary:{clientName:c.name,reference:c.reference,contractAmount:c.budget,totalPayments,totalExpenses,budgetRemaining,outstandingBalance:Math.max(0,c.budget-totalPayments),cashPosition,bankBalance:cb.bankBalance,cashOnHand:cb.cashOnHand,totalCash:cb.totalCash},cashBalances:cb,tools:tools,payments:p.slice(-100).reverse(),expenses:ex.slice(-100).reverse(),budget,recent,activity:av};
 if(u.role==="Admin"){let uv=rows("users"),fc=uv.filter(x=>String(x[3])==="Finance"&&String(x[4]).toLowerCase()!=="false").length;o.users=uv.map(x=>({username:x[0],name:x[1],role:x[3],active:String(x[4]).toLowerCase()!=="false"}));o.financeCount=fc;o.financeRegistrationOpen=String(settingsMap().RegistrationOpen).toLowerCase()!=="false"}return o;
}
function addPayment(r,u){if(!r.clientId)throw Error("Select a client/project.");if(num(r.amount)<=0)throw Error("Amount paid must be greater than zero.");let cm=clientNameMap();if(!cm[String(r.clientId)]||!cm[String(r.clientId)].active)throw Error("Invalid or inactive client/project.");ss().getSheetByName("payments").appendRow([new Date(),r.date||"",r.clientId,r.reference||"",r.description||"",num(r.dueAmount),num(r.amount),r.method||"",r.notes||"",u.name,u.username]);audit("ADD_PAYMENT",u,(cm[String(r.clientId)].name+" | "+(r.description||"")+" | "+r.amount));return{ok:true}}
function addExpense(r,u){if(!r.clientId)throw Error("Select a client/project.");if(num(r.amount)<=0)throw Error("Amount must be greater than zero.");let cm=clientNameMap();if(!cm[String(r.clientId)]||!cm[String(r.clientId)].active)throw Error("Invalid or inactive client/project.");let receipt=saveReceipt_(r.receiptData,r.receiptName,u,cm[String(r.clientId)]);ss().getSheetByName("expenses").appendRow([new Date(),r.date||"",r.clientId,r.type||"Expense",r.category||"Other",r.payee||"",r.description||"",num(r.amount),r.method||"",r.reference||"",r.approvedBy||"",r.notes||"",u.name,u.username,receipt?receipt.url:"",receipt?receipt.id:""]);let detail=(cm[String(r.clientId)].name+" | "+(r.category||"")+" | "+(r.description||"")+" | "+r.amount)+(receipt?" | Receipt attached":" | No receipt attached");audit("ADD_EXPENSE",u,detail);return{ok:true,receipt:receipt}}
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
  let hasPayments=rows("payments").some(x=>String(x[2])===id);
  let hasExpenses=rows("expenses").some(x=>String(x[2])===id);
  if(hasPayments||hasExpenses)throw Error("This client/project has financial records. Archive it instead of deleting it.");
  let sh=ss().getSheetByName("clients"),v=sh.getDataRange().getValues(),i=v.slice(1).findIndex(x=>String(x[0])===id);
  sh.deleteRow(i+2);
  audit("DELETE_CLIENT",u,c.name);
  return{ok:true,message:"Client / project deleted."};
}
function deletePayment(r,u){
  adminOnly(u);
  let row=Number(r.row);if(!Number.isInteger(row)||row<2)throw Error("Valid payment row is required.");
  let sh=ss().getSheetByName("payments");if(row>sh.getLastRow())throw Error("Payment record not found.");
  let vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  audit("DELETE_PAYMENT",u,String(vals[2]||"")+" | "+String(vals[4]||"")+" | "+num(vals[6]));
  sh.deleteRow(row);
  return{ok:true,message:"Payment deleted."};
}
function deleteExpense(r,u){
  adminOnly(u);
  let row=Number(r.row);if(!Number.isInteger(row)||row<2)throw Error("Valid expense row is required.");
  let sh=ss().getSheetByName("expenses");if(row>sh.getLastRow())throw Error("Expense record not found.");
  let vals=sh.getRange(row,1,1,sh.getLastColumn()).getValues()[0];
  audit("DELETE_EXPENSE",u,String(vals[2]||"")+" | "+String(vals[6]||"")+" | "+num(vals[7]));
  sh.deleteRow(row);
  return{ok:true,message:"Expense deleted."};
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
    status:String(x[8]||"Borrowed"),
    notes:String(x[9]||""),
    enteredBy:String(x[10]||""),
    username:String(x[11]||"")
  }));
}

function listTools(r,u){
  return{ok:true,tools:listToolsData_()};
}

function addTool(r,u){
  if(!r.tool)throw Error("Tool / Equipment name is required.");
  if(!r.borrower)throw Error("Borrowed By is required.");

  ss().getSheetByName("tools").appendRow([
    new Date(),
    String(r.tool),
    String(r.toolId||""),
    String(r.borrower),
    String(r.project||""),
    String(r.borrowed||""),
    String(r.expectedReturn||""),
    String(r.returned||""),
    String(r.status||"Borrowed"),
    String(r.notes||""),
    u.name,
    u.username
  ]);

  audit("ADD_TOOL",u,String(r.tool)+" | Borrowed by: "+String(r.borrower));
  return{ok:true};
}

function updateTool(r,u){
  adminOrFinance_(u);
  let row=Number(r.row);
  if(!row||row<2)throw Error("Invalid tool record.");

  let sh=ss().getSheetByName("tools");
  if(row>sh.getLastRow())throw Error("Tool record not found.");

  let old=sh.getRange(row,1,1,12).getValues()[0];
  sh.getRange(row,1,1,12).setValues([[
    old[0],
    String(r.tool||old[1]),
    String(r.toolId||old[2]),
    String(r.borrower||old[3]),
    String(r.project||old[4]),
    String(r.borrowed||old[5]),
    String(r.expectedReturn||old[6]),
    String(r.returned||old[7]),
    String(r.status||old[8]),
    String(r.notes||old[9]),
    u.name,
    u.username
  ]]);

  audit("UPDATE_TOOL",u,String(r.tool||old[1])+" | Status: "+String(r.status||old[8]));
  return{ok:true};
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
function audit(a,u,d){ss().getSheetByName("audit").appendRow([new Date(),a,u.username,u.name,d])}