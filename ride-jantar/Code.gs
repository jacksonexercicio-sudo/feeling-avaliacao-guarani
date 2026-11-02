/** Feeling World Ride — Votação do jantar | Apps Script vinculado à planilha */
const PAISES=['Espanha','Canadá','França','Ruanda','África do Sul','Emirados Árabes Unidos','Japão','Colômbia','Taiwan','Nova Zelândia'];
function doGet(e){
 const p=(e&&e.parameter)||{},cb=String(p.callback||'');
 if(!/^ridecb_[0-9]{13}_[a-z0-9]{7}$/.test(cb)) return ContentService.createTextOutput('Callback inválido');
 try{
  const ss=SpreadsheetApp.getActiveSpreadsheet(),controle=ss.getSheetByName('Controle');
  if(!controle)return jsonp(cb,{ok:false,code:'error',message:'Falta a aba Controle.'});
  const aberta=String(controle.getRange('B2').getDisplayValue()).trim().toUpperCase()==='ABERTA';
  if(p.action==='status')return jsonp(cb,{ok:true,open:aberta});
  if(p.action!=='vote')return jsonp(cb,{ok:false,code:'invalid',message:'Ação inválida.'});
  if(!aberta)return jsonp(cb,{ok:false,code:'closed',message:'A votação foi encerrada.'});
  const nome=String(p.name||'').trim().replace(/\s+/g,' '),ultimos=String(p.last4||'').trim(),pais=String(p.country||'').trim();
  if(nome.length<5||nome.length>80||nome.split(' ').length<2||!/^[0-9]{4}$/.test(ultimos)||!PAISES.includes(pais))
   return jsonp(cb,{ok:false,code:'invalid',message:'Confira o nome, sobrenome, 4 dígitos e país.'});
  const texto=nome.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'')+'|'+ultimos;
  const identificador=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,texto,Utilities.Charset.UTF_8).map(x=>(x&255).toString(16).padStart(2,'0')).join('');
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
   if(String(controle.getRange('B2').getDisplayValue()).trim().toUpperCase()!=='ABERTA')return jsonp(cb,{ok:false,code:'closed',message:'A votação foi encerrada.'});
   const votos=ss.getSheetByName('Votos');
   if(!votos)return jsonp(cb,{ok:false,code:'error',message:'Falta a aba Votos.'});
   const n=votos.getLastRow();
   if(n>1&&votos.getRange(2,3,n-1,1).getValues().some(l=>String(l[0])===identificador))
    return jsonp(cb,{ok:false,code:'duplicate',message:'Você já registrou seu voto.'});
   votos.appendRow([new Date(),nome.replace(/^[=+@\-]/,''),identificador,pais]);
   return jsonp(cb,{ok:true,code:'saved',message:'Voto registrado!'});
  }finally{lock.releaseLock()}
 }catch(err){console.error(err);return jsonp(cb,{ok:false,code:'error',message:'Erro ao registrar o voto. Tente novamente.'})}
}
function jsonp(cb,data){return ContentService.createTextOutput(cb+'('+JSON.stringify(data)+');').setMimeType(ContentService.MimeType.JAVASCRIPT)}