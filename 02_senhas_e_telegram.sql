-- ============================================================
-- Senhas do app web de Pedidos + Telegram do "Sem Pedido Hoje"
-- Rodar DEPOIS do 01_funcoes_web.sql, no SQL Editor do Supabase.
--
-- 1) Troque os textos entre aspas pelos valores reais
--    (os mesmos do Secrets do Streamlit do Pedidos:
--     SENHA_ADMIN, SENHA_LOJAS e a seção [telegram]: bot_token e chat_id).
-- 2) Clique em Run.
-- 3) NÃO salve esta query com os valores reais, nem suba para o GitHub
--    (o repositório é público). Depois de rodar, feche a aba sem salvar.
--
-- As senhas ficam com hash (bcrypt): nem quem abre a tabela consegue ler.
-- Para trocar uma senha depois, rode só a linha dela com o valor novo.
-- ============================================================

-- senha do "Administrador"
select public.pedidos_web_definir_senha('Administrador', 'moli0000');

-- senha das lojas (Loja 01 a Loja 08), a mesma para todas, como hoje
select public.pedidos_web_definir_senha('Lojas', 'moli1234');

-- (opcional) senha própria para uma loja; ela passa a valer só para essa loja:
-- select public.pedidos_web_definir_senha('Loja 03', 'SENHA_SO_DA_LOJA_03');

-- Telegram do aviso "Sem Pedido Hoje" (fica só no servidor; a página nunca vê)
insert into public.pedidos_web_segredos (chave, valor) values
  ('telegram_bot_token', '8868504104:AAH-IPHAuXYpWZKuWocT7KtdCbwASVpzQek'),
  ('telegram_chat_id',   '-5573584423')
on conflict (chave) do update set valor = excluded.valor;

select chave, case when chave like 'senha_%' then '(hash guardado)' else '(valor guardado)' end as situacao
from public.pedidos_web_segredos order by chave;
