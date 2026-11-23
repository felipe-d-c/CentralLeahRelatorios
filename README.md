# Central de Relatórios Léah

Um site só, protegido por senha, com menu para escolher o relatório.

| Página | O que é | Dados |
|---|---|---|
| Relatório Comercial | metas, comissão, vendedoras | `cache/` (Actions à meia-noite) |
| Relatório de Vendas | vendas por clínica | o mesmo `cache/` |
| Fichas de clientes | ficha de sessões dos planos ativos | API ao clicar em **Gerar**; cache na branch `cache-fichas` |

## Deploy (Streamlit Cloud)
1. Suba esta pasta num repositório GitHub. Main file: `app.py`.
2. **Secrets** do app: copie `.streamlit/secrets.toml.example` e preencha `[tokens]` e `[github]`.
3. Secrets do repositório (Actions): `TOKEN_VIDEIRA`, `TOKEN_JOACABA`, `TOKEN_TREZE_TILIAS`, `TOKEN_CAMPOS_NOVOS`, `TOKEN_ROTA`.
4. Senha: padrão `12qwaszx`; para trocar sem mexer no código, secret `SENHA_CENTRAL`.

## Estrutura
- `app.py` — senha + menu (`st.navigation`). Cada módulo roda como página, sem edição no código dele.
- `modulos/comercial`, `modulos/vendas` — os dois zips originais. Mudanças: cache único em `cache/` (raiz); o de vendas teve `belle_api`/`relatorio` renomeados (`vendas_*`) para não colidir; correção do PDF do mês em andamento.
- `modulos/fichas` — `fichas_core.py` (regra e PDF), `github_cache.py`, `app.py` (tela).
- `belle_core.py` — cliente da API Belle (40 req/min) para os próximos módulos.

## Fichas: como o cache funciona
Só grava quando você clica **Gerar fichas** (nada de Actions). Grava `fichas/<unidade>.json` na branch
`cache-fichas` (criada sozinha). Branch separada porque o Streamlit Cloud reimplanta o app a cada commit
na principal, o que interromperia a geração. "Último cache do GitHub" reabre sem chamar a API.
Sem `[github]` nos secrets, grava numa pasta local temporária.

## Pendente de validar com a API real
O PDF da API não mostra o formato das respostas de `relatorios/sessoes_planos` e `clientes`.
O módulo tenta vários nomes de campo e, se não achar, mostra as chaves recebidas. No 1º uso confira o
"Diagnóstico da API". Plano B: aba **Enviar planilhas**.

## Novas páginas (tudo online, direto da API)
- **Relatórios** (`reports/catalogo.py`, 31 no total): escolha o tipo, unidades e período → tabela + Excel/CSV/PDF.
  Destaques: clientes sem atendimento há +N dias (celular, último serviço, link WhatsApp), aniversariantes,
  planos (saldo / vencendo / sem agendamento), boletos e injetáveis, formas de pagamento, faturamento,
  vendedoras, melhores clientes, NPS, cancelados/faltas, agenda futura, comissão, estoque, vouchers, contas.
- **Pacote completo**: vários relatórios de uma vez num ZIP de planilhas (`ERROS.txt` lista os que falharam).
- **Explorador da API**: consulta GET livre; mostra colunas e resposta bruta (use para conferir nomes de campos).

### Se um relatório der "coluna não encontrada"
A tela pergunta qual coluna é a certa e gera de novo sem baixar tudo outra vez. Se algo falhar, abra o
Explorador com o mesmo endpoint e me mande as colunas.

### Cache
Resultados e a base de clientes (celular/nascimento) são guardados na branch `cache-fichas` do GitHub
**só quando você gera** (checkbox). "Abrir último resultado salvo" lê sem chamar a API.
A base de clientes leva vários minutos na 1ª vez por unidade (100 por página, 40 req/min).

### Não testado com a API real
Caminhos dos endpoints foram lidos do PDF do Belle (OCR); os nomes dos campos de resposta são presumidos.
Testado só com API simulada.
