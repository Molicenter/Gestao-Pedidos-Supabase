# Gestão de Pedidos · app web (etapa 1)

Painel e digitação dos pedidos das lojas, em HTML/CSS/JS, publicado pelo GitHub Pages (pasta `docs`).
Grava nas mesmas tabelas do Streamlit (`pedidos_lancamentos`, `pedidos_logs`, `pedidos_sem_pedido_hoje`,
`pedidos_observacoes_lojas`) pelas funções `pedidos_web_*` do Supabase. A central (Separação, Fornecedores,
Catálogo, Log) continua no Streamlit até as etapas 2 e 3.

- `config.js`: endereço do Supabase e chave pública (publishable); `FATOR_ALERTA_LOJA` (10x a média).
- Senhas e o Telegram ficam só no banco (SQL 02), nunca neste repositório.
- Estoque: tabela `pedidos_estoque_erp`, atualizada pela rotina do servidor 192.168.10.69 a cada 30 min.
