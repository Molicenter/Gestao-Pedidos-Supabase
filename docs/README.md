# Gestão de Pedidos · app web

Painel e digitação dos pedidos das lojas, em HTML/CSS/JS, publicado pelo GitHub Pages (pasta `docs`).
Grava nas mesmas tabelas do Streamlit (`pedidos_lancamentos`, `pedidos_logs`, `pedidos_sem_pedido_hoje`,
`pedidos_observacoes_lojas`) pelas funções `pedidos_web_*` do Supabase. A central (Separação, Fornecedores,
Catálogo, Log) continua no Streamlit até as etapas 2 e 3.

- `config.js`: endereço do Supabase e chave pública (publishable); `FATOR_ALERTA_LOJA` (10x a média).
- Senhas e o Telegram ficam só no banco (SQL 02), nunca neste repositório.
- Estoque e médias: tabelas `pedidos_estoque_erp` e `pedidos_medias_90d`, atualizadas pela rotina do servidor 192.168.10.69
  (status em `pedidos_sync_status`). Se a rotina parar, a central mostra um aviso e o Telegram recebe um alerta às 8h (seg-sáb).

## Versão 1.5
- **Histórico**: cada *Limpar pedidos* guarda o pedido final do ciclo em `pedidos_historico` (com a média de venda).
  Aba *Histórico* na central (atual × último ciclo, ciclos anteriores, pedido ÷ média por loja) e coluna *Últ.* na digitação da loja.
- **Sem internet**: cada valor digitado fica guardado no aparelho até o banco confirmar; sobe sozinho quando a conexão volta
  (ou na próxima vez que abrir o setor no mesmo aparelho).
- **Monitor do ERP**: aviso na central + Telegram (pg_cron/pg_net) quando estoque ou médias param de atualizar.
- **Tabelas fechadas** (depois de desligar o Streamlit): SQL `pedidos_web_v1.5_fechar_tabelas.sql`.
