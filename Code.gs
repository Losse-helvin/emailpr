/** UBE Pro — emailing léger depuis Google Sheets. */
const ONGLET_PROSPECTS = 'Prospects';
const ONGLET_CONFIG = 'Config';
const ENTETES = ['#','Entreprise','Secteur','Ville / Pays','Contact','Email','Type d’email','Objet','Email personnalisé','id','statut','date_envoi','nb_relances','date_relance','date_clic','erreur'];
const STATUTS = ['À envoyer','Envoyé','Intéressé','Désabonné','Relancé','Erreur'];
const CONFIG_DEFAUT = {lp_url:'',logo_url:'',signature:'',nom_expediteur:'UBE Pro',jours_avant_relance:'3',max_relances:'2',limite_quotidienne:'40',token:''};

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(ONGLET_PROSPECTS);
  if (!sh) {
    const ancien = ss.getSheetByName('Sheet1');
    sh = ancien || ss.insertSheet(ONGLET_PROSPECTS);
    if (sh.getName() !== ONGLET_PROSPECTS) sh.setName(ONGLET_PROSPECTS);
  }
  if (sh.getMaxColumns() < ENTETES.length) sh.insertColumnsAfter(sh.getMaxColumns(), ENTETES.length - sh.getMaxColumns());
  const existants = sh.getRange(1,1,1,Math.max(ENTETES.length, sh.getLastColumn())).getDisplayValues()[0];
  ENTETES.forEach((h,i) => { if (!existants[i]) sh.getRange(1,i+1).setValue(h); });
  let cfg = ss.getSheetByName(ONGLET_CONFIG);
  if (!cfg) { cfg = ss.insertSheet(ONGLET_CONFIG); cfg.getRange(1,1,1,2).setValues([['clé','valeur']]); }
  const vals = cfg.getLastRow() > 1 ? cfg.getRange(2,1,cfg.getLastRow()-1,2).getDisplayValues() : [];
  const map = {}; vals.forEach(r => { if (r[0]) map[r[0]] = r[1]; });
  if (!map.token) map.token = Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'');
  Object.keys(CONFIG_DEFAUT).forEach(k => { if (map[k] === undefined || map[k] === '') map[k] = CONFIG_DEFAUT[k] || map[k]; });
  cfg.clearContents();
  cfg.getRange(1,1,1,2).setValues([['clé','valeur']]);
  cfg.getRange(2,1,Object.keys(CONFIG_DEFAUT).length,2).setValues(Object.keys(CONFIG_DEFAUT).map(k => [k,map[k] || '']));
  cfg.setFrozenRows(1); sh.setFrozenRows(1);
  const n = Math.max(0, sh.getLastRow()-1);
  if (n) {
    const ids = sh.getRange(2,10,n,1).getValues();
    ids.forEach((r,i) => { if (!r[0]) ids[i][0] = nouvelId_(); });
    sh.getRange(2,10,n,1).setValues(ids);
    const stat = sh.getRange(2,11,n,1).getValues();
    stat.forEach(r => { if (!r[0]) r[0] = 'À envoyer'; });
    sh.getRange(2,11,n,1).setValues(stat);
    const rel = sh.getRange(2,13,n,1).getValues(); rel.forEach(r => { if (r[0] === '') r[0] = 0; }); sh.getRange(2,13,n,1).setValues(rel);
  }
  SpreadsheetApp.flush();
  Logger.log('Installation terminée. Le token initial est dans Config!B' + (Object.keys(CONFIG_DEFAUT).indexOf('token')+2));
}

function doGet(e) {
  const q = (e && e.parameter) || {}, action = q.action || '';
  if (action === 'list') {
    const cfg = lireConfig_();
    if (!cfg.token || !constantEqual_(q.token || '', cfg.token)) return json_({ok:false,error:'Accès refusé : token invalide.'});
    return json_({ok:true,prospects:lireProspects_(),config:cfg});
  }
  if (action === 'devis' || action === 'stop') {
    const id = String(q.id || '');
    if (!id || id.length > 80) return html_('<h2>Lien invalide.</h2>');
    const sh = feuille_(), rows = sh.getDataRange().getValues(), idx = rows.findIndex((r,i) => i && String(r[9]) === id);
    if (idx < 1) return html_('<h2>Lien introuvable ou expiré.</h2>');
    const row = idx+1;
    if (action === 'stop') {
      sh.getRange(row,11).setValue('Désabonné');
      return html_('<!doctype html><html lang="fr"><meta charset="utf-8"><meta name="viewport" content="width=device-width"><title>Désabonnement</title><body style="font:16px Arial,sans-serif;max-width:560px;margin:12vh auto;padding:24px;color:#263238"><h2>Désabonnement confirmé</h2><p>Vous ne recevrez plus de message de notre part.</p></body></html>');
    }
    if (String(rows[idx][10]) !== 'Désabonné') { sh.getRange(row,11).setValue('Intéressé'); sh.getRange(row,15).setValue(new Date()); }
    const lp = lireConfig_().lp_url;
    if (!lp) return html_('<!doctype html><html lang="fr"><meta charset="utf-8"><body style="font:16px Arial;margin:12vh auto;max-width:560px;padding:24px"><h2>Merci pour votre intérêt</h2><p>La page de présentation n’est pas encore configurée.</p></body></html>');
    const target = lp + (lp.indexOf('?') === -1 ? '?' : '&') + 'id=' + encodeURIComponent(id);
    return html_(
      '<!doctype html><html lang="fr">' +
      '<meta charset="utf-8">' +
      '<meta name="viewport" content="width=device-width">' +
      '<title>Continuer</title>' +
      '<body style="font:16px Arial,sans-serif;max-width:560px;margin:12vh auto;padding:24px;color:#263238">' +
      '<p>Merci pour votre intérêt.</p>' +
      '<p><a href="' + escapeAttr_(target) + '" target="_top" rel="noopener" ' +
      'style="display:inline-block;background:#238636;color:#fff;text-decoration:none;padding:12px 18px;border-radius:5px">' +
      'Continuer vers la page de présentation</a></p>' +
      '</body></html>'
    );
  }
  return html_('<!doctype html><html lang="fr"><meta charset="utf-8"><body style="font:16px Arial;padding:30px">Service d’e-mailing UBE Pro opérationnel.</body></html>');
}

function doPost(e) {
  try {
    const p = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const cfg = lireConfig_();
    if (!cfg.token || !constantEqual_(String(p.token || ''), cfg.token)) return json_({ok:false,error:'Accès refusé : token invalide.'});
    if (p.action === 'sauver_config') return sauverConfig_(p.config || {});
    if (p.action !== 'envoyer' && p.action !== 'relancer') return json_({ok:false,error:'Action inconnue.'});
    return traiterEnvois_(p.action, p.ids);
  } catch (err) { return json_({ok:false,error:String(err && err.message || err)}); }
}

function traiterEnvois_(action, ids) {
  const sh = feuille_(), cfg = lireConfig_(), data = sh.getDataRange().getValues();
  if (data.length < 2) return json_({ok:true,sent:0,errors:0,skipped:0,message:'Aucun prospect.'});
  const maintenant = new Date(), limites = Math.max(0, parseInt(cfg.limite_quotidienne,10) || 40);
  const aujourdHui = new Date(maintenant.getFullYear(),maintenant.getMonth(),maintenant.getDate());
  const envoisAujourdHui = data.slice(1).filter(r => r[11] instanceof Date && r[11] >= aujourdHui).length;
  let restant = Math.max(0, limites-envoisAujourdHui), sent=0, errors=0, skipped=0;
  const selection = Array.isArray(ids) ? new Set(ids.map(String)) : null;
  // Toute adresse déjà envoyée (ou désabonnée/intéressée) est exclue, y compris les lignes dupliquées.
  const emailsDejaTraites = new Set(data.slice(1).filter(r => r[11] || ['Envoyé','Relancé','Intéressé','Désabonné'].includes(String(r[10]))).map(r => normaliserEmail_(r[5])).filter(Boolean));
  const dejaDansLot = new Set();
  for (let i=1; i<data.length; i++) {
    const r=data[i], statut=String(r[10]||'À envoyer'), id=String(r[9]||''), email=normaliserEmail_(r[5]);
    if (action === 'envoyer') {
      if (selection ? !selection.has(id) : statut !== 'À envoyer') continue;
      if (!selection && statut !== 'À envoyer') continue;
      if (selection && statut !== 'À envoyer') { skipped++; continue; }
    } else {
      if (!['Envoyé','Relancé'].includes(statut) || Number(r[12]||0) >= (parseInt(cfg.max_relances,10)||2)) continue;
      const derniereDate = statut === 'Relancé' && r[13] instanceof Date ? r[13] : r[11];
      if (!(derniereDate instanceof Date)) continue;
      const diff = (maintenant-derniereDate)/86400000;
      if (diff < (parseInt(cfg.jours_avant_relance,10)||3)) continue;
    }
    if (['Intéressé','Désabonné','Erreur'].includes(statut) || !email || (action === 'envoyer' && emailsDejaTraites.has(email)) || dejaDansLot.has(email)) { skipped++; continue; }
    if (restant <= 0) { skipped++; continue; }
    try {
      if (action === 'relancer') envoyerUn_(r,cfg,true); else envoyerUn_(r,cfg,false);
      sh.getRange(i+1,11).setValue(action === 'relancer' ? 'Relancé' : 'Envoyé');
      sh.getRange(i+1,12).setValue(maintenant);
      sh.getRange(i+1,16).clearContent();
      if (action === 'relancer') { sh.getRange(i+1,13).setValue(Number(r[12]||0)+1); sh.getRange(i+1,14).setValue(maintenant); }
      sent++; restant--; emailsDejaTraites.add(email); dejaDansLot.add(email);
      if (restant > 0) Utilities.sleep(1000 + Math.floor(Math.random()*2001));
    } catch (err) {
      sh.getRange(i+1,11).setValue('Erreur'); sh.getRange(i+1,16).setValue(String(err && err.message || err)); errors++;
      console.error('Échec ligne '+(i+1)+': '+err);
    }
  }
  return json_({ok:true,sent:sent,errors:errors,skipped:skipped,remainingToday:restant,message:sent+' e-mail(s) envoyé(s), '+errors+' erreur(s), '+skipped+' ignoré(s).'});
}

function envoyerUn_(r,cfg,relance) {
  if (!cfg.lp_url) throw new Error('Configurez d’abord l’URL de la page de présentation.');
  const email=normaliserEmail_(r[5]); if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('Adresse e-mail invalide.');
  const id=String(r[9]||''); if (!id) throw new Error('Identifiant manquant.');
  const entreprise=String(r[1]||''), contact=String(r[4]||'');
  const objetBase=String(r[7]||'');
  const objet=relance ? (objetBase ? (/^re:/i.test(objetBase) ? objetBase : 'Re: '+objetBase) : 'Suite à mon message') : objetBase;
  const texte=relance ? (String(cfg.texte_relance||'').trim() || 'Je me permets de revenir vers vous concernant mon précédent message. Un échange rapide permettrait de vérifier si l’automatisation de certains flux internes pourrait vous être utile.') : String(r[8]||'');
  if (!texte.trim()) throw new Error('Objet ou message vide.');
  const webapp=ScriptApp.getService().getUrl(); if (!webapp) throw new Error('Déployez le script en application Web avant l’envoi.');
  const devis=webapp+'?action=devis&id='+encodeURIComponent(id), stop=webapp+'?action=stop&id='+encodeURIComponent(id);
  const corps=texte.replace(/\{\{\s*contact\s*\}\}/gi,contact).replace(/\{\{\s*entreprise\s*\}\}/gi,entreprise);
  const logo=cfg.logo_url ? '<tr><td style="padding:0 0 20px"><img src="'+escapeAttr_(cfg.logo_url)+'" alt="UBE Pro" style="max-width:180px;height:auto"></td></tr>' : '';
  const html='<div style="font-family:Arial,sans-serif;color:#263238;font-size:15px;line-height:1.6"><table role="presentation" style="width:100%;max-width:620px;border-collapse:collapse"><tbody>'+logo+'<tr><td style="white-space:pre-line">'+escapeHtml_(corps).replace(/\n/g,'<br>')+'</td></tr><tr><td style="padding:24px 0 10px"><a href="'+escapeAttr_(devis)+'" style="display:inline-block;background:#238636;color:#fff;text-decoration:none;padding:12px 20px;border-radius:5px;font-weight:bold">Demander un devis</a>&nbsp;&nbsp;<a href="'+escapeAttr_(stop)+'" style="display:inline-block;background:#e9ecef;color:#263238;text-decoration:none;padding:12px 20px;border-radius:5px">Ça ne m’intéresse pas</a></td></tr><tr><td style="border-top:1px solid #e5e7eb;padding-top:16px;color:#59636e;font-size:13px;white-space:pre-line">'+escapeHtml_(cfg.signature||'')+(cfg.signature?'\n':'')+escapeHtml_(cfg.nom_expediteur||'UBE Pro')+' — UBE Pro<br><br><a href="'+escapeAttr_(stop)+'" style="color:#59636e">Ne plus recevoir nos messages : se désabonner</a></td></tr></tbody></table></div>';
  GmailApp.sendEmail(email,objet,'Votre messagerie ne prend pas en charge le HTML. '+corps,{htmlBody:html,name:cfg.nom_expediteur||'UBE Pro'});
}

function sauverConfig_(incoming) {
  const allowed=['lp_url','logo_url','signature','nom_expediteur','jours_avant_relance','max_relances','limite_quotidienne','texte_relance'];
  const cfg=lireConfig_();
  allowed.forEach(k => { if (incoming[k] !== undefined) cfg[k]=String(incoming[k]).trim(); });
  if (!cfg.lp_url || !/^https:\/\//i.test(cfg.lp_url)) return json_({ok:false,error:'L’URL de la page de présentation doit commencer par https://'});
  ['jours_avant_relance','max_relances','limite_quotidienne'].forEach(k => { if (!/^\d+$/.test(String(cfg[k])) || Number(cfg[k]) < 0) throw new Error('Valeur invalide pour '+k+'.'); });
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_CONFIG), keys=Object.keys(cfg);
  sh.clearContents(); sh.getRange(1,1,1,2).setValues([['clé','valeur']]);
  sh.getRange(2,1,keys.length,2).setValues(keys.map(k=>[k,cfg[k]]));
  return json_({ok:true,message:'Configuration enregistrée.'});
}

function lireConfig_() {
  const sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_CONFIG);
  if (!sh) throw new Error('Onglet Config introuvable. Lancez setup().');
  const out=Object.assign({},CONFIG_DEFAUT,{texte_relance:''});
  if (sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,2).getDisplayValues().forEach(r=>{if(r[0])out[r[0]]=r[1];});
  return out;
}
function lireProspects_() {
  const sh=feuille_(), vals=sh.getDataRange().getValues();
  return vals.slice(1).filter(r=>r.some(v=>v!=='')).map((r,i)=>({ligne:i+2,numero:r[0],entreprise:r[1]||'',secteur:r[2]||'',ville:r[3]||'',contact:r[4]||'',email:r[5]||'',typeEmail:r[6]||'',objet:r[7]||'',message:r[8]||'',id:r[9]||'',statut:r[10]||'À envoyer',dateEnvoi:dateIso_(r[11]),nbRelances:Number(r[12]||0),dateRelance:dateIso_(r[13]),dateClic:dateIso_(r[14]),erreur:r[15]||''}));
}
function feuille_(){const s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(ONGLET_PROSPECTS);if(!s)throw new Error('Onglet Prospects introuvable. Lancez setup().');return s;}
function nouvelId_(){return Utilities.getUuid().replace(/-/g,'').slice(0,16)+Utilities.getUuid().replace(/-/g,'').slice(0,8);}
function dateIso_(v){return v instanceof Date ? v.toISOString() : (v||'');}
function normaliserEmail_(v){return String(v||'').trim().toLowerCase();}
function constantEqual_(a,b){a=String(a);b=String(b);if(a.length!==b.length)return false;let x=0;for(let i=0;i<a.length;i++)x|=a.charCodeAt(i)^b.charCodeAt(i);return x===0;}
function json_(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);}
function html_(s){return HtmlService.createHtmlOutput(s).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}
function escapeHtml_(s){return String(s||'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));}
function escapeAttr_(s){return escapeHtml_(s);}
