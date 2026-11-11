# Divino Grão · Sistema

Caixa e gestão do Armazém Divino Grão (Jarinu/SP). Roda no navegador, publicado pelo GitHub Pages, com os dados guardados numa planilha do Google.

**Módulos:** Caixa · Comandas · Produtos · Fiado · Estoque · Fechamentos · Financeiro (Fluxo de caixa e Lançamentos) · Caixa do dia · Ajustes.

## Como funciona

```
navegador (GitHub Pages)  ──login / ler / gravar──▶  Apps Script (App da Web)  ──▶  Google Sheets
```

- `index.html` — o sistema (telas e regras).
- `support.js` — o motor que desenha as telas (React).
- `banco.js` — login, leitura e gravação na planilha. Grava linha a linha só o que mudou e, a cada 15 s, traz o que outros computadores lançaram. Se a internet cair, guarda as alterações no navegador e envia quando voltar.
- `config.js` — o endereço do App da Web. Vazio, o sistema roda só no navegador, sem planilha.
- `apps-script/Code.gs` — o código que fica dentro da planilha.

## Abas da planilha

| Aba | O que guarda |
|---|---|
| Usuarios | quem pode entrar (usuário, nome, perfil, ativo, senha) |
| Produtos | catálogo, preços, custo, código interno, PLU da balança |
| Comandas | mesas e comandas abertas |
| Vendas | uma linha por venda (caixa aberto ou já fechado) |
| ItensVendidos | uma linha por item vendido — base para relatórios |
| Fechamentos | ficha de cada fechamento de caixa |
| Fiado | contas de fiado com saldo |
| Fornecedores | cadastro de fornecedores |
| Estoque_Lotes / Estoque_Movimentos | lotes por produto e movimentações |
| Lancamentos_Despesas / Lancamentos_Entradas | lançamentos do financeiro |
| Despesas_Fixas, Dividas, Contas_a_Pagar, Imobilizado | apoio ao fluxo de caixa |
| Registro | últimas ações feitas no sistema |
| Estado | caixa aberto agora, ajustes e configurações |
| _Excluidos | o que foi apagado, para os outros computadores apagarem também |

Cada linha tem colunas legíveis e a coluna `_json` (oculta), que é o registro completo usado pelo sistema. Para mudar um dado, use o sistema; edições feitas direto nas colunas legíveis não voltam para o sistema.

## Usuários e senhas

Na aba **Usuarios**, escreva a senha na coluna `senha_nova`. No primeiro login ela vira um código (`senha_hash`) e a célula fica vazia. Para trocar a senha, escreva outra em `senha_nova`. Para bloquear alguém, ponha `não` em `ativo`.

## Instalação do Apps Script (uma vez)

1. Na planilha: **Extensões › Apps Script**, cole o conteúdo de `apps-script/Code.gs` e salve.
2. Escolha a função `configurar` e clique em **Executar** (autorize quando pedir). As abas são criadas.
3. **Implantar › Nova implantação › App da Web** · Executar como: **Eu** · Quem pode acessar: **Qualquer pessoa**.
4. Copie a URL que termina em `/exec` para `config.js`.

Ao mudar o `Code.gs`, use **Implantar › Gerenciar implantações › Editar › Nova versão** para manter a mesma URL.
