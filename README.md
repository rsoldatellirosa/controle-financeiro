# Controle Financeiro

Dashboard mensal de gastos pessoais que lê os dados de uma planilha do Google Sheets. Roda como Web App do Google Apps Script, sem servidor e sem custo.

Cards de resumo, gráfico de rosca e de barras por categoria, tabela de percentual e listagem separando gastos fixos de variáveis.

## A decisão de arquitetura

A versão óbvia desse dashboard é uma página estática que chama a API do Sheets direto do navegador com uma chave de API. Funciona, e tem dois problemas que não dá para contornar:

1. **A chave fica visível.** Qualquer coisa em JavaScript no cliente é legível por quem abrir a página. Restrição por HTTP referrer ajuda contra uso casual, mas é um header — quem quiser, forja.
2. **Pior: a planilha precisa ser pública.** Uma chave de API não autentica *quem* está pedindo, só *qual projeto*. Para ela conseguir ler, a planilha tem que estar como "qualquer pessoa com o link". Ou seja: o dado fica aberto para quem descobrir o ID — com ou sem a chave.

O segundo ponto é o que decide. Não existe configuração que deixe uma chave de API ler uma planilha privada.

A solução aqui é um **Web App do Apps Script publicado com _executar como: eu_ e _acesso: somente eu_**:

- o navegador nunca fala com a API do Sheets — chama `google.script.run`, e o backend (`Codigo.gs`) lê a planilha
- o código roda com a identidade do dono, então a planilha fica **Restrito**
- não existe chave de API em lugar nenhum
- o Google exige login na conta autorizada antes de servir a página

O custo: sob `script.google.com` não dá para servir manifest nem service worker, então o app deixa de ser PWA instalável. No celular vira atalho na tela inicial, e não funciona offline.

## Arquivos

```
Codigo.gs     Backend: doGet() serve a página, getDados(aba) lê a planilha
index.html    Front: HTML, CSS e os gráficos (Chart.js via CDN)
```

## Como usar

### 1. A planilha

Uma aba por mês, nomeada `mes_ano` em minúsculas e sem acento — `agosto_2026`, `marco_2026`. Colunas A a E:

| Descrição | Valor | Categoria | Data | Pagamento |
|---|---|---|---|---|
| Mercado | R$ 450,00 | Alimentação | 05/08 | Pix |

Linhas cuja descrição contém "fixo" são agrupadas como gastos fixos; o resto vira variável.

Categorias com cor própria: Saúde, Alimentação, Bem-estar, Lazer/Entret., Fixos/Compromissos, Seguros, Vestuário, Assinat. Tech. Qualquer outra cai em "outros".

### 2. O script

1. Na planilha: **Extensões** → **Apps Script**
2. Cola o `Codigo.gs` no arquivo `.gs` e preenche o `SHEET_ID` (está na URL da planilha, entre `/d/` e `/edit`)
3. Cria um arquivo HTML chamado **`index`** e cola o `index.html`
   (o nome importa: `createHtmlOutputFromFile('index')` diferencia maiúscula de minúscula)
4. Roda a função `testarLeitura` uma vez para autorizar e conferir que lê a aba do mês
5. **Implantar** → **Nova implantação** → **App da Web**
   - *Executar como*: **Eu**
   - *Quem pode acessar*: **Somente eu**
6. Abre a URL `/exec` que aparece no fim

### 3. Fechar a planilha

**Compartilhar** → *Acesso geral* → **Restrito**. O Web App continua lendo normalmente, porque roda como você.

## Detalhe de implementação

Valores vêm de `getDisplayValues()`, ou seja, já formatados pela planilha. A conversão para número trata os dois formatos possíveis:

```js
"R$ 1.234,56"  ->  1234.56    // vírgula decimal, ponto de milhar
"1234.56"      ->  1234.56    // ponto decimal
"1.000"        ->  1000       // ponto de milhar sozinho
```

O caso do meio e o de baixo se distinguem porque separador de milhar sempre vem seguido de exatamente três dígitos.

## Licença

MIT — veja [LICENSE](LICENSE).
