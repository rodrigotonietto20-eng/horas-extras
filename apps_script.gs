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

function getSheet(){
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Solicitacoes');
}

function doGet(e){
  var sheet = getSheet();
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var solicitacoes = [];
  for(var i=1;i<rows.length;i++){
    var row = rows[i];
    if(!row[0]) continue; // sem ID, linha vazia
    var obj = {};
    CAMPOS.forEach(function(c){
      var col = headers.indexOf(c.header);
      if(col<0) return;
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
  return ContentService.createTextOutput(JSON.stringify({ok:true, solicitacoes:solicitacoes}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  var body = JSON.parse(e.postData.contents);
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
