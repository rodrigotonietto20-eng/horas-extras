/**
 * Backend do app "Horas Extras — Rede Sander".
 *
 * COMO USAR:
 * 1. Crie uma nova planilha Google Sheets (separada da planilha do CRM de Visitas).
 * 2. Renomeie a primeira aba para "Solicitacoes".
 * 3. Na linha 1 dessa aba, cole exatamente estes cabeçalhos, um por coluna:
 *    ID | Filial | Area | Funcionario | Quantidade | Motivo | Data | Status | Historico | CriadoEm
 * 4. Menu Extensões → Apps Script. Apague o conteúdo padrão e cole este arquivo inteiro.
 * 5. Menu Implantar → Nova implantação → tipo "App da Web".
 *    - Executar como: Eu (sua conta)
 *    - Quem pode acessar: Qualquer pessoa
 * 6. Copie a URL gerada (termina em /exec) e cole na variável SCRIPT_URL no topo do index.html.
 * 7. Sempre que reimplantar como NOVA implantação, a URL muda — atualize o index.html de novo.
 *    Reimplantar como "nova versão" na MESMA implantação não muda a URL.
 *
 * Lição aprendida no CRM de Visitas: usar um mapa explícito cabeçalho↔campo (CAMPOS abaixo)
 * em vez de comparar o texto do cabeçalho direto com o nome da propriedade — muito mais
 * resistente a alguém reformatar a planilha no futuro.
 */

var CAMPOS = [
  {header:'ID',           key:'id'},
  {header:'Filial',       key:'fil',          num:true},
  {header:'Area',         key:'area'},
  {header:'Funcionario',  key:'funcionario'},
  {header:'Quantidade',   key:'quantidade',   num:true},
  {header:'Motivo',       key:'motivo'},
  {header:'Data',         key:'data',         date:true},
  {header:'Status',       key:'status'},
  {header:'Historico',    key:'historico',    json:true},
  {header:'CriadoEm',     key:'criadoEm',     num:true}
];

// Catálogo de filiais/postos (2026-07-28) — antes só existia hardcoded no index.html (objeto
// POSTOS), sem jeito de cadastrar filial nova sem editar código. Aba nova, auto-criada aqui na
// primeira chamada (mesmo padrão já usado no app de Abertura de Postos pra abas novas).
var CAMPOS_POSTO = [
  {header:'Area',   key:'area'},
  {header:'Nome',   key:'nome'},
  {header:'Filial', key:'fil', num:true}
];

function getSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Solicitacoes');
}

function getSheetPostos(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('Postos');
  if(!sh){
    sh = ss.insertSheet('Postos');
    sh.getRange(1,1,1,CAMPOS_POSTO.length).setValues([CAMPOS_POSTO.map(function(c){ return c.header; })]);
  }
  return sh;
}

// Normaliza texto de cabeçalho (trim + minúsculas) antes de comparar — sem isso, alguém
// retypar "funcionario" em vez de "Funcionario" na planilha (foi exatamente o que aconteceu
// e deixou o nome do funcionário sumir de todas as telas de alçada) faz o indexOf falhar e
// cair no fallback posicional; melhor nem precisar do fallback pra essa causa específica.
function normalizarHeader(h){ return String(h||'').trim().toLowerCase(); }

function lerPostos(){
  var sheet = getSheetPostos();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(normalizarHeader);
  var out = [];
  for(var i=1;i<rows.length;i++){
    var row = rows[i];
    if(!row[0]) continue;
    var obj = {};
    CAMPOS_POSTO.forEach(function(c, idx){
      var col = headers.indexOf(normalizarHeader(c.header));
      if(col<0) col = idx;
      var val = row[col];
      if(c.num) val = val===''||val==null ? 0 : Number(val);
      obj[c.key] = val;
    });
    out.push(obj);
  }
  return out;
}

function doGet(e){
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0].map(normalizarHeader);
  var solicitacoes = [];
  for(var i=1;i<rows.length;i++){
    var row = rows[i];
    if(!row[0]) continue; // sem ID, linha vazia
    var obj = {};
    CAMPOS.forEach(function(c, idx){
      var col = headers.indexOf(normalizarHeader(c.header));
      // Se o texto do cabeçalho não bate (célula vazia, renomeada, acento diferente...),
      // cai pra posição fixa da coluna — CAMPOS sempre está na mesma ordem que doPost
      // escreve, então a coluna certa continua sendo lida mesmo com o cabeçalho quebrado.
      if(col<0) col = idx;
      var val = row[col];
      if(c.json){ try{ val = val ? JSON.parse(val) : []; }catch(err){ val = []; } }
      else if(c.num){ val = val===''||val==null ? 0 : Number(val); }
      else if(c.date && Object.prototype.toString.call(val)==='[object Date]'){
        // A célula "Data" recebe uma string tipo "2026-07-24", mas o Sheets detecta o
        // formato e converte pra um valor de data de verdade — sem isso, getValues()
        // devolve um objeto Date que vira timestamp completo ao serializar (vira lixo
        // tipo "2026-07-24T03:00:00.000Z" em vez de "2026-07-24" na tela do app).
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      }
      obj[c.key] = val;
    });
    solicitacoes.push(obj);
  }
  return ContentService.createTextOutput(JSON.stringify({ok:true, solicitacoes:solicitacoes, postos:lerPostos()}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  // Trava a execução: sem isso, duas sincronizações simultâneas (dois aparelhos, ou um
  // retry do mesmo aparelho) rodam o "apaga as linhas do ID, depois insere de novo" ao
  // mesmo tempo. Cada execução lê a planilha ANTES da outra terminar de apagar, então
  // nenhuma das duas vê a linha nova que a outra acabou de inserir — sobra linha
  // duplicada com o mesmo ID. Foi isso que causou solicitações já aprovadas aparecendo
  // como pendente de novo: bastava uma sincronização futura tocar só numa das cópias.
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try{
    return doPostSemTrava(e);
  } finally {
    lock.releaseLock();
  }
}

function doPostSemTrava(e){
  var body = JSON.parse(e.postData.contents);

  if(body.postos){
    var novosP = body.postos || [];
    if(novosP.length){
      var sheetP = getSheetPostos();
      var idsNovosP = {};
      novosP.forEach(function(p){ idsNovosP[p.fil] = true; });
      var dataP = sheetP.getDataRange().getValues();
      for(var i=dataP.length-1;i>=1;i--){ if(idsNovosP[dataP[i][2]]) sheetP.deleteRow(i+1); }
      var linhasP = novosP.map(function(p){
        return CAMPOS_POSTO.map(function(c){ return p[c.key]==null ? '' : p[c.key]; });
      });
      sheetP.getRange(sheetP.getLastRow()+1, 1, linhasP.length, CAMPOS_POSTO.length).setValues(linhasP);
    }
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }

  var novas = body.solicitacoes || [];
  if(!novas.length) return ContentService.createTextOutput(JSON.stringify({ok:true}));

  var sheet = getSheet();
  var headers = sheet.getDataRange().getValues()[0];

  // Upsert por ID: remove as linhas cujo ID bate com algum item recebido, depois insere
  // o lote inteiro de novo. Não depende de quem está sincronizando (funciona igual pra
  // posto, gerente de área ou operações), evita duplicar linha.
  var idsNovos = {};
  novas.forEach(function(s){ idsNovos[s.id] = true; });

  var data = sheet.getDataRange().getValues();
  var idCol = headers.indexOf('ID');
  if(idCol<0) idCol = 0; // mesma rede de segurança do doGet — ID é sempre a 1ª coluna
  for(var i=data.length-1;i>=1;i--){
    if(idsNovos[data[i][idCol]]) sheet.deleteRow(i+1);
  }

  var linhas = novas.map(function(s){
    return CAMPOS.map(function(c){
      var val = s[c.key];
      if(c.json) return JSON.stringify(val||[]);
      return val==null ? '' : val;
    });
  });
  if(linhas.length){
    sheet.getRange(sheet.getLastRow()+1, 1, linhas.length, CAMPOS.length).setValues(linhas);
  }

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
