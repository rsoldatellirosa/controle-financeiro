/**
 * Controle Financeiro — backend do Web App
 *
 * Os dados nunca vão para o navegador sem passar por aqui, e o Web App roda
 * como o dono da planilha — por isso a planilha pode ficar como "Restrito".
 */

// ID da planilha — está na URL dela, entre /d/ e /edit.
// Fica no servidor e nunca chega ao navegador.
// Usar openById em vez de getActive() faz o script funcionar tanto vinculado
// à planilha (Extensões > Apps Script) quanto avulso (script.google.com).
var SHEET_ID = 'COLE_AQUI_O_ID_DA_SUA_PLANILHA';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Controle Financeiro')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Lê uma aba da planilha e devolve as linhas como texto exibido.
 * Usa getDisplayValues() para manter exatamente o mesmo formato que a API REST
 * devolvia antes (valores já formatados), sem quebrar o parsing do front.
 *
 * @param {string} aba Nome da aba, ex.: "agosto_2026"
 * @return {string[][]} Linhas de A até E, incluindo o cabeçalho
 */
function getDados(aba) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(aba);
  if (!sheet) {
    throw new Error('Aba nao encontrada: ' + aba);
  }
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  return sheet.getRange(1, 1, lastRow, 5).getDisplayValues();
}

/**
 * Teste rápido: rode esta função no editor (botão "Executar") antes de implantar.
 * Ela autoriza o script e mostra no log quantas linhas leu do mês atual.
 */
function testarLeitura() {
  var meses = ['janeiro','fevereiro','marco','abril','maio','junho',
               'julho','agosto','setembro','outubro','novembro','dezembro'];
  var hoje = new Date();
  var aba  = meses[hoje.getMonth()] + '_' + hoje.getFullYear();
  var dados = getDados(aba);
  Logger.log('Aba: %s — %s linhas lidas', aba, dados.length);
}
