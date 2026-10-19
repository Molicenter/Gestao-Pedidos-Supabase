// Configuração do Supabase (projeto da empresa).
// A chave abaixo é a PÚBLICA (publishable): pode ficar no GitHub sem problema,
// porque o app só chama as funções pedidos_web_* do banco, e todas exigem login.
window.CONFIG = {
  SUPABASE_URL: "https://eemfqkxshqqlhdyofgoz.supabase.co",
  SUPABASE_KEY: "sb_publishable_ZwBOMu1KwPQB2AwNeh08tw_g93KC9d5",
  // aviso na loja quando o pedido passa de N vezes a média (a central vê 3x na Separação)
  FATOR_ALERTA_LOJA: 10,
  VERSAO: "1.7.1",
};
