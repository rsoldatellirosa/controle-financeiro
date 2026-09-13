# Controle Financeiro

Dashboard mensal de gastos pessoais. Roda como Web App do Google Apps Script e le uma planilha do Google Sheets. Sem servidor e sem chave de API: o navegador chama `google.script.run` e o backend le a planilha como dono, por isso a planilha fica "Restrito".

Este arquivo e versionado num repo publico. Identificadores (ID da planilha, URL do Web App, ID do projeto, chaves) ficam so na memoria local do Claude, nunca aqui.

## Pastas

- Este repo: versao publica. `SHEET_ID` e placeholder. Nunca commitar o ID real.
- `C:\Users\rsold\controle-financeiro-appsscript\`: fonte real para colar no Apps Script (`Codigo.gs` + `Index.html`), com o `SHEET_ID` verdadeiro. Nao e repo git. Toda edicao feita aqui precisa ser sincronizada para la trocando o placeholder pelo ID real.

## Arquivos

- `Codigo.gs`: `doGet` serve a pagina. `getDados(aba)` le a planilha. `onOpen` cria o menu "Financeiro". `atualizarResumo` escreve o resumo em G1:H6. `enviarAnaliseMensal` e `montarEmailAnalise_` mandam a analise por email. `criarGatilhoMensal` instala o gatilho. `corrigirTagFatura` e `corrigirErrosPontuais` sao utilitarios de limpeza, idempotentes.
- `index.html`: front, Chart.js via CDN. Texto vindo da planilha entra por `textContent`, nunca `innerHTML`.

## Planilha

Uma aba por mes, nome `mes_ano` em minusculo sem acento (`agosto_2026`). Colunas A a E: Descricao, Valor, Categoria, Data, Pagamento. Descricao contendo "fixo" conta como gasto fixo. H1 guarda o faturamento (editavel, padrao 7600).

## Modelo do cartao

Regra central. Nao mudar sem conversar.

A fatura do banco e a verdade do gasto no cartao. Os itens de cartao lancados a mao sao incompletos e servem so para dar categoria.

| Pagamento | Entra no total | Papel |
|---|---|---|
| `PIX` ou qualquer outro | sim | gasto normal |
| `Cartão` (inclui `Cartão (10x) - Parcela N`) | sim, como item | detalhe por categoria |
| `Fatura` | nao, e referencia | valor real cobrado pelo banco |

Total gasto = gastos fora da fatura + fatura. A diferenca entre a fatura e os itens de cartao lancados vira o balde "Cartao - nao detalhado", para as categorias fecharem com o total sem contar o cartao duas vezes. Exatamente uma linha `Fatura` por mes, a `Cartão Nubank - Fixo`. A mesma logica existe em `index.html` e em `Codigo.gs`; mudou num, muda no outro.

## Deploy

Aqui e onde as coisas dao errado.

- O dashboard (`/exec`) serve a versao IMPLANTADA. Depois de editar `index.html`: Implantar > Gerenciar implantacoes > lapis > Versao: Nova versao > Implantar. So salvar nao muda nada no dashboard.
- Menu e gatilhos rodam o codigo SALVO (head). Para mudanca em `Codigo.gs` basta colar e salvar; nao precisa de nova versao.
- Cliques automatizados nao disparam dentro do iframe sandbox do Apps Script (so digitar funciona). Teste de interface e manual, guiado para o usuario clicar.
- Funcoes utilitarias rodam pelo editor: escolher a funcao no seletor, Executar, conferir o Registro de execucao.

## Analise mensal por email

Gatilho no dia 10 as 20h. O cartao vira por volta do dia 10; no dia 1 a fatura ainda nao existe e o total sairia sem o cartao. No envio automatico, fatura vazia gera um lembrete curto em vez de analise errada. Envio manual pelo menu passa direto. Calculo determinista, sem chamada de IA, sem credencial.

## Seguranca

- Nunca deixar a planilha publica. Chave de API do Sheets so le planilha publica, e e por isso que a arquitetura e Apps Script. Se alguma ferramenta pedir para abrir a planilha, a resposta e nao.
- Existe uma chave de API antiga no historico git do repo privado `controle_financeiro`. Deve ser deletada no Google Cloud Console, nao so rotacionada.

## Verificacao antes de entregar

Antes de pedir para o usuario colar codigo: simular a logica em Node com os dados reais (a planilha pode ser lida pelo Drive MCP) e conferir que os totais batem com o dashboard ao vivo. Toda mudanca em `Codigo.gs` que mexe em dados da planilha deve ser idempotente e listar no log exatamente o que alterou.

## Estilo

Sem travessao e sem frase com cara de IA em nada que for entregue: commits, README, emails, este arquivo.
