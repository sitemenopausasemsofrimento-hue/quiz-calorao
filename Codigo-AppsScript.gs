// Apps Script da planilha "quiz calorão"
// Extensões > Apps Script > cole este código > Implantar > Nova implantação > App da Web
// (Executar como: eu · Quem pode acessar: qualquer pessoa) > copie a URL /exec para APPS_SCRIPT_URL no index.html
const PERGUNTAS = ['momento','cafe','jantar','nome','idade','curiosidade_tentou','estresse','gatilho_emocao','curiosidade_natural','sono','outros_sinais'];

function doPost(e){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let d = {};
  try { d = JSON.parse(e.postData.contents); } catch(err) { return ContentService.createTextOutput('erro json'); }
  const quando = d.data || new Date().toISOString();
  if (d.tipo === 'lead') {
    const aba = ss.getSheetByName('leads') || ss.insertSheet('leads');
    if (aba.getLastRow() === 0) aba.appendRow(['data','sessao','nome','idade','whatsapp','gatilho','sinal'].concat(PERGUNTAS));
    const r = d.respostas || {};
    aba.appendRow([quando, d.sessao||'', d.nome||'', d.idade||'', "'"+(d.whatsapp||''), d.gatilho||'', d.sinal||''].concat(PERGUNTAS.map(p => r[p] || '')));
  } else {
    const aba = ss.getSheetByName('respostas') || ss.insertSheet('respostas');
    if (aba.getLastRow() === 0) aba.appendRow(['data','sessao','tipo','passo','pergunta','resposta','gatilho']);
    aba.appendRow([quando, d.sessao||'', d.tipo||'', d.passo||'', d.pergunta||'', d.resposta||'', d.gatilho||'']);
  }
  return ContentService.createTextOutput('ok');
}

// Opcional: menu para montar a aba "resumo" (desistência por passo e opção mais marcada por pergunta)
function onOpen(){ SpreadsheetApp.getUi().createMenu('Quiz').addItem('Atualizar resumo','montarResumo').addToUi(); }
function montarResumo(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const resp = ss.getSheetByName('respostas'); if(!resp) return;
  const dados = resp.getDataRange().getValues().slice(1);
  const out = ss.getSheetByName('resumo') || ss.insertSheet('resumo'); out.clear();
  const inicios = new Set(dados.filter(r=>r[2]==='inicio').map(r=>r[1])).size;
  out.appendRow(['Funil', 'Pessoas']); out.appendRow(['Iniciaram o quiz', inicios]);
  for (let p=1;p<=PERGUNTAS.length;p++){ const n = new Set(dados.filter(r=>r[2]==='resposta' && r[3]==p).map(r=>r[1])).size; out.appendRow(['Responderam passo '+p+' ('+PERGUNTAS[p-1]+')', n]); }
  const leads = ss.getSheetByName('leads'); out.appendRow(['Deixaram WhatsApp', leads ? Math.max(0, leads.getLastRow()-1) : 0]);
  out.appendRow([]); out.appendRow(['Pergunta','Resposta','Cliques']);
  const cont = {};
  dados.filter(r=>r[2]==='resposta' && r[4]!=='nome').forEach(r=>{ const k=r[4]+'||'+r[5]; cont[k]=(cont[k]||0)+1; });
  Object.keys(cont).sort().forEach(k=>{ const [p,r]=k.split('||'); out.appendRow([p,r,cont[k]]); });
  if (leads && leads.getLastRow()>1){
    out.appendRow([]); out.appendRow(['Gatilho no resultado','Leads']);
    const g = {}; leads.getDataRange().getValues().slice(1).forEach(r=>{ g[r[5]]=(g[r[5]]||0)+1; });
    Object.keys(g).forEach(k=>out.appendRow([k,g[k]]));
  }
}
