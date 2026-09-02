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

// Faturamento usado quando a celula H1 da aba do mes esta vazia.
var FATURAMENTO_PADRAO = 7600;

// Rotulo do balde que representa a parte da fatura do cartao que nao foi
// lancada item a item. Mesmo nome usado no front (index.html).
var CAT_GAP = 'Cartao - nao detalhado';

var MESES       = ['janeiro','fevereiro','marco','abril','maio','junho',
                   'julho','agosto','setembro','outubro','novembro','dezembro'];
var MESES_EXIB  = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                   'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// Menu na planilha para rodar o resumo sem sair dela.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Financeiro')
    .addItem('Atualizar resumo do mes', 'atualizarResumo')
    .addItem('Enviar analise do mes passado agora', 'enviarAnaliseMensal')
    .addItem('Instalar envio automatico (dia 10)', 'criarGatilhoMensal')
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
  var hoje = new Date();
  var aba  = MESES[hoje.getMonth()] + '_' + hoje.getFullYear();
  var dados = getDados(aba);
  Logger.log('Aba: %s — %s linhas lidas', aba, dados.length);
}


/**
 * Correcao pontual dos meses antigos.
 *
 * Antes do modelo atual, a linha da fatura do Nubank ficava com Pagamento
 * "Cartao" — entao ela era somada junto com os itens do cartao e o mes inteiro
 * saia inflado (a soma dupla). Isto reetiqueta essas linhas como "Fatura",
 * que e o que o dashboard e a analise esperam.
 *
 * So mexe em linhas cuja descricao contem "cartao nubank", e ignora as que ja
 * estao como "Fatura" — pode rodar quantas vezes quiser sem estragar nada.
 * Rode pelo editor (escolha a funcao e clique em Executar) e veja o log.
 *
 * @return {string[]} O que foi alterado, uma linha por celula.
 */
function corrigirTagFatura() {
  var mudou = [];
  SpreadsheetApp.openById(SHEET_ID).getSheets().forEach(function(sheet) {
    var nome  = sheet.getName();
    var parte = nome.split('_');
    if (parte.length !== 2 || MESES.indexOf(parte[0]) === -1) return;  // so abas de mes
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    var vals = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
    vals.forEach(function(r, i) {
      if (norm_(r[0]).indexOf('cartao nubank') === -1) return;   // so a linha da fatura
      if (norm_(r[4]).indexOf('fatura') !== -1) return;          // ja corrigida
      sheet.getRange(i + 2, 5).setValue('Fatura');
      mudou.push(nome + ' linha ' + (i + 2) + ': "' + r[4] + '" -> "Fatura"');
    });
  });
  Logger.log(mudou.length ? mudou.join('\n') : 'Nada a corrigir — todos os meses ja estao certos.');
  return mudou;
}


/* ==========================================================================
 * ANALISE MENSAL POR EMAIL
 *
 * Roda dentro do Apps Script, entao le a planilha como dono: sem chave de API
 * e sem precisar deixar a planilha publica. O email sai pelo MailApp, sem
 * OAuth nem token guardado em lugar nenhum.
 * ========================================================================== */

// Tira acento e caixa alta, para comparar "Cartão" com "cartao".
function norm_(s) {
  return String(s == null ? '' : s).toLowerCase()
    .normalize('NFD').replace(/[̀-ͯ]/g, '').trim();
}

function fmtBRL_(v) {
  var neg = v < 0;
  var p   = Math.abs(v).toFixed(2).split('.');
  var int = p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  return (neg ? '-' : '') + 'R$ ' + int + ',' + p[1];
}

function abaDe_(ano, mesIdx) { return MESES[mesIdx] + '_' + ano; }
function nomeDe_(ano, mesIdx) { return MESES_EXIB[mesIdx] + ' de ' + ano; }

/**
 * Le uma aba e devolve os numeros do mes ja no modelo correto do cartao:
 * a linha "Fatura" e a verdade do gasto no cartao (valor real do banco) e nao
 * e somada junto com os itens; a diferenca entre ela e os itens de cartao
 * lancados vira o balde CAT_GAP, para o total fechar com o banco sem contar
 * o cartao duas vezes.
 *
 * @return {Object|null} null se a aba nao existe ou esta vazia.
 */
function analisarAba_(aba) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(aba);
  if (!sheet) return null;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var vals = sheet.getRange(2, 1, lastRow - 1, 5).getDisplayValues();
  var itens = [], fatCartao = 0, cartaoLanc = 0;

  vals.forEach(function(r) {
    if (!r[0]) return;                              // sem descricao = linha vazia
    var valor = parseValor_(r[1]);
    var pgto  = norm_(r[4]);
    if (pgto.indexOf('fatura') !== -1) {
      fatCartao += valor;                           // referencia do banco
      return;
    }
    itens.push({ desc: r[0], valor: valor, cat: r[2] || 'outros', data: r[3], pgto: r[4] });
    if (pgto.indexOf('cartao') !== -1) cartaoLanc += valor;
  });

  if (!itens.length && !fatCartao) return null;

  var cartaoGap = Math.max(0, fatCartao - cartaoLanc);
  var gasto = itens.reduce(function(s, i) { return s + i.valor; }, 0) + cartaoGap;

  var cats = {};
  itens.forEach(function(i) { if (i.valor > 0) cats[i.cat] = (cats[i.cat] || 0) + i.valor; });
  if (cartaoGap > 0) cats[CAT_GAP] = (cats[CAT_GAP] || 0) + cartaoGap;

  // Gasto fixo e marcado pela palavra "fixo" na descricao, igual ao dashboard.
  // O balde do cartao entra como variavel: e compra avulsa, nao compromisso.
  var fixos = 0, variaveis = cartaoGap;
  itens.forEach(function(i) {
    if (norm_(i.desc).indexOf('fixo') !== -1) fixos += i.valor; else variaveis += i.valor;
  });

  var maior = itens.slice().sort(function(a, b) { return b.valor - a.valor; })[0] || null;
  var fatur = parseValor_(sheet.getRange('H1').getDisplayValue()) || FATURAMENTO_PADRAO;

  return {
    aba: aba, gasto: gasto, fatur: fatur, sobra: fatur - gasto,
    fatCartao: fatCartao, cartaoLanc: cartaoLanc, cartaoGap: cartaoGap,
    cats: cats, fixos: fixos, variaveis: variaveis, maior: maior, nItens: itens.length
  };
}

/** Media por categoria nos meses anteriores. Mes sem a categoria conta como 0. */
function mediaHistorica_(anteriores) {
  if (!anteriores.length) return {};
  var soma = {};
  anteriores.forEach(function(a) {
    Object.keys(a.cats).forEach(function(k) { soma[k] = (soma[k] || 0) + a.cats[k]; });
  });
  var media = {};
  Object.keys(soma).forEach(function(k) { media[k] = soma[k] / anteriores.length; });
  return media;
}

// Variacao entre dois valores. Em gasto, subir e ruim (laranja), cair e bom (verde).
function delta_(atual, ant) {
  if (!ant) return { txt: '—', cor: '#888780' };
  var d = atual - ant;
  if (Math.abs(d) < 0.01) return { txt: 'igual', cor: '#888780' };
  var p = Math.round(d / ant * 100);
  return {
    txt: (d > 0 ? '+' : '') + fmtBRL_(d) + '  (' + (d > 0 ? '+' : '') + p + '%)',
    cor: d > 0 ? '#e05a2b' : '#1d9e75'
  };
}

/**
 * Envia por email a analise do mes que acabou. Chamada pelo gatilho mensal
 * (dia 1) e pelo menu "Financeiro".
 */
function enviarAnaliseMensal(e) {
  var hoje = new Date();
  var ref  = new Date(hoje.getFullYear(), hoje.getMonth() - 1, 1);   // mes fechado
  var ano  = ref.getFullYear(), mes = ref.getMonth();

  var atual = analisarAba_(abaDe_(ano, mes));

  // Gatilho de tempo recebe um evento; chamada pelo menu, nao. Serve para o
  // envio automatico ser cauteloso e o manual passar direto.
  var automatico = !!(e && e.triggerUid);

  if (!atual) {
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: 'Analise Financeira — ' + nomeDe_(ano, mes) + ' (sem dados)',
      htmlBody: '<p style="font-family:sans-serif">A aba <code>' + abaDe_(ano, mes) +
                '</code> nao existe ou esta vazia. Nada para analisar neste mes.</p>'
    });
    return;
  }

  // A fatura do cartao so e lancada quando o cartao vira, por volta do dia 10.
  // Sem ela o total do mes sai subestimado, entao o envio automatico manda um
  // lembrete em vez de uma analise errada. Pelo menu, envia mesmo assim.
  if (automatico && atual.fatCartao === 0 && atual.cartaoLanc > 0) {
    MailApp.sendEmail({
      to: Session.getEffectiveUser().getEmail(),
      subject: 'Controle Financeiro — falta a fatura de ' + nomeDe_(ano, mes),
      htmlBody:
        '<div style="font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif;' +
        'max-width:520px;margin:32px auto;padding:24px 28px;border:1px solid #e8e8e3;border-radius:12px">' +
        '<div style="font-size:17px;font-weight:700;color:#1a1a1a">Falta a fatura de ' + esc_(nomeDe_(ano, mes)) + '</div>' +
        '<p style="font-size:14px;color:#1a1a1a;line-height:1.6">A linha com Pagamento <b>Fatura</b> da aba ' +
        '<code>' + esc_(abaDe_(ano, mes)) + '</code> esta sem valor. Sem ela o total do mes sai errado, ' +
        'entao a analise nao foi gerada.</p>' +
        '<p style="font-size:14px;color:#1a1a1a;line-height:1.6">Preenche o valor da fatura e depois roda ' +
        '<b>Financeiro &gt; Enviar analise do mes passado agora</b> no menu da planilha.</p>' +
        '<div style="font-size:12px;color:#888780;margin-top:16px;padding-top:12px;border-top:1px solid #e8e8e3">' +
        'Ja lancados ' + fmtBRL_(atual.cartaoLanc) + ' em itens de cartao neste mes.</div></div>'
    });
    Logger.log('Fatura de %s ainda vazia — lembrete enviado.', nomeDe_(ano, mes));
    return;
  }

  // Ate 6 meses anteriores, do mais recente para o mais antigo.
  var anteriores = [];
  for (var k = 1; k <= 6; k++) {
    var d = new Date(ano, mes - k, 1);
    var a = analisarAba_(abaDe_(d.getFullYear(), d.getMonth()));
    if (a) { a.nome = nomeDe_(d.getFullYear(), d.getMonth()); anteriores.push(a); }
  }

  var html = montarEmailAnalise_(atual, nomeDe_(ano, mes), anteriores);
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Analise Financeira — ' + nomeDe_(ano, mes),
    htmlBody: html
  });
  Logger.log('Analise de %s enviada.', nomeDe_(ano, mes));
}

// Escapa texto vindo da planilha antes de entrar no HTML do email.
function esc_(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function montarEmailAnalise_(a, nomeMes, anteriores) {
  var ant   = anteriores[0] || null;             // mes imediatamente anterior
  var media = mediaHistorica_(anteriores);
  var pctComp = a.fatur > 0 ? Math.round(a.gasto / a.fatur * 100) : 0;

  var C = { borda:'#e8e8e3', texto:'#1a1a1a', fraco:'#888780',
            verde:'#1d9e75', laranja:'#e05a2b', azul:'#2878c8', fundo:'#f5f5f0' };

  function card(rot, val, cor, sub) {
    return '<td style="padding:6px" width="33%" valign="top">' +
      '<div style="background:#fff;border:1px solid ' + C.borda + ';border-radius:10px;padding:14px 16px">' +
      '<div style="font-size:10px;font-weight:700;color:' + C.fraco + ';text-transform:uppercase;letter-spacing:0.06em">' + rot + '</div>' +
      '<div style="font-size:19px;font-weight:700;color:' + cor + ';margin-top:6px">' + val + '</div>' +
      (sub ? '<div style="font-size:11px;color:' + C.fraco + ';margin-top:3px">' + sub + '</div>' : '') +
      '</div></td>';
  }
  function titulo(t) {
    return '<div style="font-size:11px;font-weight:700;color:' + C.fraco + ';text-transform:uppercase;' +
           'letter-spacing:0.08em;margin:28px 0 12px;padding-bottom:6px;border-bottom:1px solid ' + C.borda + '">' + t + '</div>';
  }

  var h = '<div style="background:' + C.fundo + ';padding:32px 0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',Arial,sans-serif">' +
    '<div style="max-width:680px;margin:0 auto;background:#fff;border:1px solid ' + C.borda + ';border-radius:14px;padding:32px 36px">' +
    '<div style="font-size:21px;font-weight:700;color:' + C.texto + '">Controle Financeiro</div>' +
    '<div style="font-size:14px;color:' + C.laranja + ';font-weight:500;margin-top:2px">' + esc_(nomeMes) + '</div>';

  // ---- Resumo -------------------------------------------------------------
  var dGasto = ant ? delta_(a.gasto, ant.gasto) : null;
  h += titulo('Resumo do mes') +
    '<table width="100%" cellpadding="0" cellspacing="0" style="margin:0 -6px"><tr>' +
    card('Total gasto', fmtBRL_(a.gasto), C.laranja, dGasto ? 'vs mes anterior: ' + dGasto.txt : '') +
    card('Sobra', fmtBRL_(a.sobra), a.sobra >= 0 ? C.azul : C.laranja, 'faturamento ' + fmtBRL_(a.fatur)) +
    card('% comprometido', pctComp + '%', C.texto, a.nItens + ' lancamentos') +
    '</tr></table>';

  // ---- Cartao -------------------------------------------------------------
  h += titulo('Cartao') +
    '<table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;color:' + C.texto + '">' +
    '<tr><td style="padding:5px 0">Fatura do banco</td><td align="right" style="padding:5px 0;font-weight:600">' + fmtBRL_(a.fatCartao) + '</td></tr>' +
    '<tr><td style="padding:5px 0">Lancado item a item</td><td align="right" style="padding:5px 0">' + fmtBRL_(a.cartaoLanc) + '</td></tr>' +
    '<tr><td style="padding:5px 0;border-top:1px solid ' + C.borda + '">Nao detalhado</td>' +
    '<td align="right" style="padding:5px 0;border-top:1px solid ' + C.borda + ';font-weight:700;color:' +
    (a.cartaoGap > a.fatCartao * 0.3 ? C.laranja : C.fraco) + '">' + fmtBRL_(a.cartaoGap) + '</td></tr>' +
    '</table>';

  // ---- Categorias: valor, share, variacao e desvio da media ---------------
  var keys = Object.keys(a.cats).sort(function(x, y) { return a.cats[y] - a.cats[x]; });
  h += titulo('Por categoria') +
    '<table width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;border-collapse:collapse">' +
    '<tr style="font-size:10px;color:' + C.fraco + ';text-transform:uppercase;letter-spacing:0.05em">' +
    '<td style="padding:0 0 6px">Categoria</td><td align="right" style="padding:0 0 6px">Total</td>' +
    '<td align="right" style="padding:0 0 6px">%</td><td align="right" style="padding:0 0 6px">vs mes ant.</td>' +
    '<td align="right" style="padding:0 0 6px">media</td></tr>';

  keys.forEach(function(k) {
    var v  = a.cats[k];
    var d  = ant ? delta_(v, ant.cats[k] || 0) : { txt: '—', cor: C.fraco };
    var m  = media[k] || 0;
    var fugiu = m > 0 && v > m * 1.5 && v - m > 100;
    h += '<tr style="border-top:1px solid #f0f0eb">' +
      '<td style="padding:7px 0">' + esc_(k) + (fugiu ? ' <span style="color:' + C.laranja + '">▲</span>' : '') + '</td>' +
      '<td align="right" style="padding:7px 0;font-weight:600">' + fmtBRL_(v) + '</td>' +
      '<td align="right" style="padding:7px 0;color:' + C.fraco + '">' + Math.round(v / a.gasto * 100) + '%</td>' +
      '<td align="right" style="padding:7px 0;color:' + d.cor + ';font-size:12px">' + d.txt + '</td>' +
      '<td align="right" style="padding:7px 0;color:' + C.fraco + ';font-size:12px">' + (m ? fmtBRL_(m) : '—') + '</td>' +
      '</tr>';
  });
  h += '</table>';
  if (anteriores.length) {
    h += '<div style="font-size:11px;color:' + C.fraco + ';margin-top:8px">' +
         'Media dos ' + anteriores.length + ' meses anteriores. ▲ marca categoria 50% acima da media.</div>';
  }

  // ---- Destaques ----------------------------------------------------------
  h += titulo('Destaques');
  var itensD = [];
  if (a.maior) {
    itensD.push('<b>Maior gasto:</b> ' + esc_(a.maior.desc) + ' — ' + fmtBRL_(a.maior.valor) +
                ' <span style="color:' + C.fraco + '">(' + esc_(a.maior.cat) + ')</span>');
  }
  itensD.push('<b>Top 3 categorias:</b> ' + keys.slice(0, 3).map(function(k) {
    return esc_(k) + ' ' + fmtBRL_(a.cats[k]);
  }).join(' · '));
  itensD.push('<b>Fixos vs variaveis:</b> ' + fmtBRL_(a.fixos) + ' fixos (' +
              Math.round(a.fixos / a.gasto * 100) + '%) · ' + fmtBRL_(a.variaveis) + ' variaveis');
  if (ant) {
    var subiu = keys.filter(function(k) { return (ant.cats[k] || 0) > 0; })
      .map(function(k) { return { k: k, d: a.cats[k] - ant.cats[k] }; })
      .sort(function(x, y) { return y.d - x.d; })[0];
    if (subiu && subiu.d > 0) {
      itensD.push('<b>Maior alta vs mes anterior:</b> ' + esc_(subiu.k) + ' +' + fmtBRL_(subiu.d));
    }
  }
  h += '<ul style="font-size:13px;color:' + C.texto + ';line-height:1.8;padding-left:18px;margin:0">' +
       itensD.map(function(x) { return '<li>' + x + '</li>'; }).join('') + '</ul>';

  // ---- Alertas ------------------------------------------------------------
  var alertas = [];
  if (a.sobra < 0) {
    alertas.push('Gasto passou do faturamento em ' + fmtBRL_(-a.sobra) + '.');
  }
  if (a.fatCartao === 0 && a.cartaoLanc > 0) {
    alertas.push('A linha "Fatura" do mes ficou sem valor. O total do cartao saiu so dos itens lancados, ' +
                 'entao o gasto real pode estar subestimado.');
  } else if (a.cartaoGap > a.fatCartao * 0.3 && a.fatCartao > 0) {
    alertas.push(fmtBRL_(a.cartaoGap) + ' da fatura (' + Math.round(a.cartaoGap / a.fatCartao * 100) +
                 '%) nao foi lancado item a item — esse pedaco fica sem categoria.');
  }
  keys.forEach(function(k) {
    var m = media[k] || 0;
    if (m > 0 && a.cats[k] > m * 1.5 && a.cats[k] - m > 100) {
      alertas.push(esc_(k) + ' ficou ' + fmtBRL_(a.cats[k] - m) + ' acima da media (' + fmtBRL_(m) + ').');
    }
  });
  if (alertas.length) {
    h += titulo('Alertas') +
      '<div style="background:#fdf6ec;border:1px solid #f0d9b5;border-radius:10px;padding:14px 18px">' +
      '<ul style="font-size:13px;color:#7a4b12;line-height:1.7;padding-left:18px;margin:0">' +
      alertas.map(function(x) { return '<li>' + x + '</li>'; }).join('') + '</ul></div>';
  }

  h += '<div style="margin-top:28px;padding-top:14px;border-top:1px solid ' + C.borda +
       ';font-size:11px;color:' + C.fraco + '">Gerado pelo Apps Script da propria planilha. ' +
       'Nenhuma chave de API envolvida — a planilha continua restrita.</div>' +
       '</div></div>';
  return h;
}

/**
 * Instala o gatilho mensal. Roda no dia 10 a noite, e nao no dia 1, porque a
 * fatura do cartao so e lancada quando o cartao vira, por volta do dia 10 —
 * antes disso o total do mes sairia sem o cartao. A hora tardia da folga para
 * o lancamento do dia acontecer antes do envio.
 * Rode uma vez.
 */
function criarGatilhoMensal() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'enviarAnaliseMensal') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarAnaliseMensal')
    .timeBased().onMonthDay(10).atHour(20).create();
  var msg = 'Pronto. A analise sera enviada todo dia 10, a noite.';
  Logger.log(msg);
  // getUi() so existe quando a funcao roda pelo menu da planilha; pelo editor, nao.
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) {}
}
