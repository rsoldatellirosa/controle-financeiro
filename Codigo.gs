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

// Menu na planilha para rodar o resumo sem sair dela.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financeiro')
    .addItem('Atualizar resumo do mes', 'atualizarResumo')
    .addToUi();
}

// Converte o valor exibido em numero. Mesma regra do front:
// virgula = decimal ("1.234,56"->1234.56); ponto sozinho = milhar ("1.000"->1000).
function parseValor_(v) {
  var s = String(v == null ? '' : v).replace(/R\$/g, '').replace(/\s/g, '');
  if (!s) return 0;
  if (s.indexOf(',') !== -1) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) {
    s = s.replace(/\./g, '');
  }
  return parseFloat(s) || 0;
}

/**
 * Escreve o resumo do mes num bloco fixo G1:H7 da aba ativa.
 * A fatura do cartao (Pagamento "Fatura") e a verdade do gasto no cartao — o
 * valor real que o banco cobrou. Os itens de cartao lancados a mao dao so o
 * detalhe e costumam ser incompletos; a diferenca (fatura menos itens) entra
 * como "Cartao nao detalhado" para o Total gasto fechar com o banco sem contar
 * o cartao duas vezes. Total gasto = gastos fora do cartao + fatura.
 * Faturamento fica na celula H1 (editavel); vazia, e semeada com 7600.
 * Rode pelo menu "Financeiro > Atualizar resumo do mes".
 */
function atualizarResumo() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  var lastRow = sheet.getLastRow();
  var itens = 0, fatCartao = 0, cartaoLanc = 0;
  if (lastRow >= 2) {
    var vals = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    vals.forEach(function(r) {
      if (!r[0]) return;                         // sem descricao = linha vazia
      var valor = parseValor_(r[1]);
      var pgto  = String(r[4] || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');
      if (pgto.indexOf('fatura') !== -1) {
        fatCartao += valor;                      // fatura do banco: a verdade do cartao
      } else {
        itens += valor;                          // gastos fora da fatura (PIX + itens cartao)
        if (pgto.indexOf('cartao') !== -1) cartaoLanc += valor;
      }
    });
  }
  var cartaoGap = Math.max(0, fatCartao - cartaoLanc);   // compras no cartao nao lancadas
  var gasto     = itens + cartaoGap;                     // fecha com o banco: PIX + fatura
  // H1 e a celula editavel do faturamento. Preserva o que estiver la; se vazia,
  // semeia com 7600 (Hapolo 6700 + Mutirao 900) para o usuario ajustar depois.
  var fatur = parseValor_(sheet.getRange('H1').getDisplayValue()) || 7600;
  var linhas = [
    ['Faturamento',            fatur],
    ['Total gasto',            gasto],
    ['Total sobra',            fatur - gasto],
    ['Fatura cartao (banco)',  fatCartao],
    ['Cartao lancado (itens)', cartaoLanc],
    ['Cartao nao detalhado',   cartaoGap],
  ];
  sheet.getRange(1, 7, linhas.length, 2).setValues(linhas);          // G1:H6
  sheet.getRange(1, 7, linhas.length, 1).setFontWeight('bold');      // rotulos
  sheet.getRange(1, 8, linhas.length, 1).setNumberFormat('R$ #,##0.00');
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
