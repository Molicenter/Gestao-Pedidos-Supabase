import streamlit as st
import pandas as pd
import io
import re
import streamlit.components.v1 as components
from streamlit_gsheets import GSheetsConnection

# ─────────────────────────────────────────────
# IMPORTAÇÕES PARA ESTILIZAÇÃO DO EXCEL
# ─────────────────────────────────────────────
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# ─────────────────────────────────────────────
# FUNÇÃO PARA EXTRAIR APENAS NÚMEROS (TRATAR CX, UND, KG)
# ─────────────────────────────────────────────
def extrair_numero_quantidade(valor):
    """
    Tenta converter o valor para número ignorando textos (ex: '01 cx' -> 1).
    Se o valor for vazio, puramente texto ou erro, retorna 0.
    """
    if pd.isna(valor) or valor == "":
        return 0
    v_str = str(valor)
    match = re.search(r'\d+', v_str)
    if match:
        return int(match.group())
    return 0

# ─────────────────────────────────────────────
# CONSTANTES E PRODUTOS INICIAIS DA MATRIZ
# ─────────────────────────────────────────────
LOJAS = ["Loja 01", "Loja 02", "Loja 03", "Loja 04", "Loja 05", "Loja 06", "Loja 07", "Loja 08"]
MAPA_LOJAS = {l: l for l in LOJAS}

produtos_iniciais = [
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 651017, "Descrição Oficial": "Abacaxi Calda 400g Tozzi Rodelas", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 6846, "Descrição Oficial": "Acucar Cristal 5kg Alto Alegre", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 594624, "Descrição Oficial": "Acucar Mascavo 500g Dacolonia", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 22545, "Descrição Oficial": "Acucar Refinado 5kg Alto Alegre", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 630591, "Descrição Oficial": "Alga Marinha 22g Tokyofoods Yakinori C/10", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 596020, "Descrição Oficial": "Amendoim 400g Caldo Bom Branco", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 34067, "Descrição Oficial": "Arroz 5kg Frias Extra Tp1", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 558057, "Descrição Oficial": "Atum Ralado 140g Pescador Oleo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 558048, "Descrição Oficial": "Atum Solido 140g Pescador Oleo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 188593, "Descrição Oficial": "Azeite Verde Fatiada 1,010kg Atalaia Sache", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 144216, "Descrição Oficial": "Azeite Oliva 500ml Borges Ex Virgem", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 451967, "Descrição Oficial": "Batata Palha 1kg Vo Maria", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 651390, "Descrição Oficial": "Bisc Maizena 300g Marilan", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 291, "Descrição Oficial": "Bombom Wafer Ouro Branco 1kg", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 284, "Descrição Oficial": "Bombom Wafer Sonho Valsa 1kg", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 226200, "Descrição Oficial": "Cafe 500g Bom Jesus Tradicional Vac", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 567107, "Descrição Oficial": "Caldo 1,01kg Apti Carne", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 639334, "Descrição Oficial": "Caldo 1,010kg Soeto Galinha", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 8604, "Descrição Oficial": "Canjiquinha 500g Zaeli", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 52498, "Descrição Oficial": "Cevada 500g Boa Saude Gold", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 217390, "Descrição Oficial": "Cha Mate 250g Leao", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 318839, "Descrição Oficial": "Champignon Fatiado 1,010kg Atalaia Sache", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 77712, "Descrição Oficial": "Choc 33g Nestle Prestigio", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 592316, "Descrição Oficial": "Choc 80g Lacta Laka", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 658269, "Descrição Oficial": "Choc 80g Nestle Alpino", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 611091, "Descrição Oficial": "Choc Wafer Bis 100,8g Lacta C16 Ao Leite", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 609713, "Descrição Oficial": "Choc Wafer Bis 100,8g Lacta C16 Laka", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 6226, "Descrição Oficial": "Conhaque 900ml Presidente", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 204998, "Descrição Oficial": "Ameixa Seca 100g La Violetera S/Caroco", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 450816, "Descrição Oficial": "Far Trigo Integral 1kg Arapongas", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 639370, "Descrição Oficial": "Creme Cebola 1,010kg Soeto", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 28325, "Descrição Oficial": "Creme Leite 200g Lider Tp", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 625058, "Descrição Oficial": "Ervilha 1,7kg Ole Lata", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 633491, "Descrição Oficial": "Extr Tom 1,020kg Ole Sache", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 526508, "Descrição Oficial": "Far Milho 1kg Caldo Bom Amarela", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 6837, "Descrição Oficial": "Far Trigo 5kg Arapongas", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 435242, "Descrição Oficial": "Cons Far Trigo 5kg Coamo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 3292, "Descrição Oficial": "Farofa Mand 500g Zaeli Tradicional", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 516981, "Descrição Oficial": "Feijao Carioca 1kg Caldo Bom", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 516990, "Descrição Oficial": "Feijao Preto 1kg Caldo Bom", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 579430, "Descrição Oficial": "Figo Calda 400g Citral Inteiros", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 3326, "Descrição Oficial": "Fuba 1kg Joia", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584575, "Descrição Oficial": "Gelatina 20g Apti Abacaxi", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584584, "Descrição Oficial": "Gelatina 20g Apti Cereja", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584593, "Descrição Oficial": "Gelatina 20g Apti Limao", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584609, "Descrição Oficial": "Gelatina 20g Apti Maracuja", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584618, "Descrição Oficial": "Gelatina 20g Apti Morango", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584627, "Descrição Oficial": "Gelatina 20g Apti Uva", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 47722, "Descrição Oficial": "Gelatina Incolor 24g Apti S/Sabor", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 647467, "Descrição Oficial": "Grao Bico 400g Pinduca", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 27403, "Descrição Oficial": "Ketchup 3,4kg Calcuta", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 474977, "Descrição Oficial": "Leite Coco 500ml Adel Coco Vidro", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 4972, "Descrição Oficial": "Leite l Vida 1l Lider Integral", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 301699, "Descrição Oficial": "Leite Po 400g Piracanjuba Instantaneo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 38805, "Descrição Oficial": "Mac Sem 500g Floriani Penne", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 30526, "Descrição Oficial": "Mac Trad 500g Floriani Espaguete", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 34494, "Descrição Oficial": "Mac Trad 500g Floriani Parafuso", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 251127, "Descrição Oficial": "Maionese 3kg Dajuda Bag", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 21067, "Descrição Oficial": "Maria Mole 50g Apti Coco", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 575461, "Descrição Oficial": "Mel 280g Sendeski Silvestre", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 521178, "Descrição Oficial": "Milho Verde 1,7kg Predilecta Lata", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 65511, "Descrição Oficial": "Molho Shoyu 900ml Kirin Tradicional", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 475587, "Descrição Oficial": "Molho Tom 1,020kg Sc Tarantella Trad", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 103848, "Descrição Oficial": "Mostarda 3,2kg Calcuta", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 364908, "Descrição Oficial": "Oleo Comp 500ml Maria Trad Pet", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 31431, "Descrição Oficial": "Oleo Soja 900ml Coamo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 130967, "Descrição Oficial": "Palmito Picado Palm Real 1.8kg Rio Cerro", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 252463, "Descrição Oficial": "Pessego Calda 450g Schramm Extra", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 607609, "Descrição Oficial": "Queijo Ral Grosso 40g Floriani", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 594138, "Descrição Oficial": "Ref 18g Tang Nutri Laranja", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584779, "Descrição Oficial": "Ref 18g Tang Nutri Limao", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584788, "Descrição Oficial": "Ref 18g Tang Nutri Maracuja", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584760, "Descrição Oficial": "Ref 18g Tang Nutri Morango", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 584803, "Descrição Oficial": "Ref 18g Tang Nutri Uva", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 649270, "Descrição Oficial": "Sagu 400g Pinduca", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 653796, "Descrição Oficial": "Sal Grosso P/Churrasco 1kg Marfim", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 653787, "Descrição Oficial": "Sal Refinado 1kg Marfim", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 4194, "Descrição Oficial": "Temp 100g Ajinomoto Refil", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 4404, "Descrição Oficial": "Temp 60g Sazon Frango", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 4413, "Descrição Oficial": "Temp 60g Sazon Legumes", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 347558, "Descrição Oficial": "Temp 60g Sazon Limao", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 197795, "Descrição Oficial": "Temp 900g Sazon Carnes", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 493130, "Descrição Oficial": "Toalha Pap C2 Tropicos 100 Toalhas", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 338709, "Descrição Oficial": "Vinagre Alcool 750ml Chemim", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 602491, "Descrição Oficial": "Cafe Sol 40g Nescafe Tradicao Ex Forte Sache", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 58902, "Descrição Oficial": "Arroz 5kg Zaeli Agulhinha Tp1", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 628930, "Descrição Oficial": "Fosforo Parana Gigante Eco C/50", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 69663, "Descrição Oficial": "Mac Inst Yakissoba 500g Nissin", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 640640, "Descrição Oficial": "Papel Alum Life Clean 45cmx7,5m", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 461797, "Descrição Oficial": "Sabao Barra 800g Alpes Super", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 50999, "Descrição Oficial": "Saco Lixo Sert Rolo 30l C/30", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 50982, "Descrição Oficial": "Saco Lixo Sert Rolo 15l C/60", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 326865, "Descrição Oficial": "Saponaceo Crem 450ml Cif Original", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 44349, "Descrição Oficial": "Vinagre Arroz 750ml Castelo", "Nome Personalizado": ""},
    {"Fornecedor": "MATÉRIA PRIMA GERAL", "Código": 587314, "Descrição Oficial": "Espeto Bambu Bompack 25cm C/50", "Nome Personalizado": ""},

    # USO DE PRODUTOS LIMPEZA
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 1133, "Descrição Oficial": "Agua Sanitaria 1l Qboa", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 623014, "Descrição Oficial": "Brilha Alum 500ml Provita Inox Original", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 158088, "Descrição Oficial": "Copo Desc C100 Coposul 180ml Transp", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 158103, "Descrição Oficial": "Copo Desc C100 Coposul 50ml Transp", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 168243, "Descrição Oficial": "Desinf L500p450ml Pinho Brill Silvestre Orig", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 178, "Descrição Oficial": "Deterg Liq 500ml Ype Neutro", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 373331, "Descrição Oficial": "Esponja C4 Brilhus Multiuso Preço Economica", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 422956, "Descrição Oficial": "Flanela Martimpanos 28x28 Ouro", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 265333, "Descrição Oficial": "Lamina C3 Fiat Lux Flexor", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 205313, "Descrição Oficial": "Limp 500ml m Uso 3x1 Facille Tradicional", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 38485, "Descrição Oficial": "Palha Aco 25g Brillo N01", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 1063, "Descrição Oficial": "Palito Dente C200 Gina", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 491910, "Descrição Oficial": "Pano Microfibra Martimpanos 50x70 Azul", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 604541, "Descrição Oficial": "Pano Microfibra Martimpanos 80x90cm Verde", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 464679, "Descrição Oficial": "Querosene 800ml Petrus", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 556721, "Descrição Oficial": "Sabao Po 800g Tixan Primavera", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 414717, "Descrição Oficial": "Saneante Desinf Hospitalar 1l Araucaria 70", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 53174, "Descrição Oficial": "Vassoura Alklin Multiuso C/Cabo", "Nome Personalizado": ""},
    {"Fornecedor": "USO DE PRODUTOS LIMPEZA", "Código": 419925, "Descrição Oficial": "Saco Martimpanos Alvejado 40x62 Branco", "Nome Personalizado": ""},

    # TEMPEROS
    {"Fornecedor": "TEMPEROS", "Código": 1656, "Descrição Oficial": "Colorau 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574675, "Descrição Oficial": "Cons Acafrao 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 588030, "Descrição Oficial": "Cons Alho Granulado 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 490300, "Descrição Oficial": "Cons Amac Carne 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 257769, "Descrição Oficial": "Cons Canela Moida 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 594767, "Descrição Oficial": "Cons Cebolinha Desidr 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 558464, "Descrição Oficial": "Cons Chimichurri 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574693, "Descrição Oficial": "Cons Cominho Em Pó 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574718, "Descrição Oficial": "Cons Curry 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 398419, "Descrição Oficial": "Cons Erva Doce 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 588049, "Descrição Oficial": "Cons Lemon Pepper 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574736, "Descrição Oficial": "Cons Louro Em Pó 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 575249, "Descrição Oficial": "Cons Noz Moscada Moida 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 263478, "Descrição Oficial": "Cons Oregano 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574745, "Descrição Oficial": "Cons Paprica Defumada 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574754, "Descrição Oficial": "Cons Paprica Doce 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 98755, "Descrição Oficial": "Cons Paprica Picante 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 499170, "Descrição Oficial": "Cons Pimenta Reino Moida 500g Androinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574684, "Descrição Oficial": "Cons Salsinha Desidratada 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 490294, "Descrição Oficial": "Cons Tempero Churrasco 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574709, "Descrição Oficial": "Cons Tempero Sirio 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 574727, "Descrição Oficial": "Cons Tomilho 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 229395, "Descrição Oficial": "Cons Vinagrete 500g Andorinha", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 425467, "Descrição Oficial": "Pimenta Calabresa 20g Andorinha Moida", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 61186, "Descrição Oficial": "Pimenta Calabresa 30g Andorinha Flocos", "Nome Personalizado": ""},
    {"Fornecedor": "TEMPEROS", "Código": 594758, "Descrição Oficial": "Cons Louro Folhas 200g Andorinha", "Nome Personalizado": ""}
]

# ─────────────────────────────────────────────
# FUNÇÃO DE ESTILIZAÇÃO DE EXCEL COM CORES
# ─────────────────────────────────────────────
def gerar_excel_estilizado(df, sheet_name="Resumo"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        worksheet = writer.sheets[sheet_name]

        header_fill = PatternFill(start_color='8B5A2B', end_color='8B5A2B', fill_type='solid') 
        alt_row_fill = PatternFill(start_color='FAF5F0', end_color='FAF5F0', fill_type='solid') 
        header_font = Font(color='FFFFFF', bold=True)
        border_style = Border(
            left=Side(style='thin', color='CCCCCC'),
            right=Side(style='thin', color='CCCCCC'),
            top=Side(style='thin', color='CCCCCC'),
            bottom=Side(style='thin', color='CCCCCC')
        )

        for row_idx, row in enumerate(worksheet.iter_rows(min_row=1, max_row=worksheet.max_row, min_col=1, max_col=worksheet.max_column)):
            for cell in row:
                cell.border = border_style
                if row_idx == 0:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    if row_idx % 2 == 0:
                        cell.fill = alt_row_fill
                    if isinstance(cell.value, (int, float)):
                        cell.alignment = Alignment(horizontal='center', vertical='center')
                    else:
                        cell.alignment = Alignment(horizontal='left', vertical='center')

        for col in worksheet.columns:
            col_letter = col[0].column_letter
            col_header = col[0].value
            
            # Se for a coluna "Observação:", fixar largura equivalente a 3 colunas padrão
            if col_header == 'Observação:':
                worksheet.column_dimensions[col_letter].width = 30
            else:
                max_length = 0
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                worksheet.column_dimensions[col_letter].width = max_length + 2

        worksheet.sheet_properties.pageSetUpPr.fitToPage = True
        worksheet.page_setup.fitToWidth = 1
        worksheet.page_setup.fitToHeight = False
        worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
        worksheet.page_margins.left = 0.25
        worksheet.page_margins.right = 0.25
        worksheet.page_margins.top = 0.75
        worksheet.page_margins.bottom = 0.75

    return buffer.getvalue()

# ─────────────────────────────────────────────
# CONEXÃO GOOGLE SHEETS & FUNÇÕES DE DADOS
# ─────────────────────────────────────────────
conn = st.connection("gsheets", type=GSheetsConnection)

WS_PRODUTOS = "MateriaP_Produtos"
WS_PEDIDOS  = "MateriaP_Pecas"

def parse_bool(x):
    if isinstance(x, bool): return x
    if isinstance(x, (int, float)): return bool(x) and not pd.isna(x)
    return str(x).strip().upper() in ['TRUE', 'VERDADEIRO', '1', 'V', 'SIM', 'YES', 'T', 'X']

@st.cache_data(ttl=15)
def carregar_catalogo_materia():
    try:
        df = conn.read(worksheet=WS_PRODUTOS, ttl=0, usecols=list(range(20)))
    except ValueError as e:
        if "Spreadsheet must be specified" in str(e):
            st.error("🚨 **Erro Crítico:** URL da Planilha não especificada nas configurações.")
            st.stop()
        else:
            raise e
    except Exception as e:
        st.error(f"Erro ao conectar na aba {WS_PRODUTOS}: {e}")
        st.stop()

    if df.empty or "Fornecedor" not in df.columns:
        df_init = pd.DataFrame(produtos_iniciais)
        for loja in LOJAS: df_init[loja] = True
        conn.update(worksheet=WS_PRODUTOS, data=df_init)
        df = df_init.copy()

    need_update = False

    if "Descrição" in df.columns and "Descrição Oficial" not in df.columns:
        df = df.rename(columns={"Descrição": "Descrição Oficial"})
        need_update = True

    if "Nome Personalizado" not in df.columns:
        df["Nome Personalizado"] = ""
        need_update = True

    if need_update:
        conn.update(worksheet=WS_PRODUTOS, data=df.drop(columns=["Descrição"], errors="ignore"))

    novas_colunas = {}
    for col in df.columns:
        col_str = str(col).strip()
        for loja in LOJAS:
            if loja.lower() in col_str.lower():
                novas_colunas[col] = loja
    df = df.rename(columns=novas_colunas)

    for loja in LOJAS:
        if loja not in df.columns:
            df[loja] = False
        else:
            df[loja] = df[loja].apply(parse_bool)

    if "Código" in df.columns:
        df["Código"] = pd.to_numeric(df["Código"], errors='coerce').fillna(0).astype(int)

    def obter_nome_final(row):
        apelido = str(row.get("Nome Personalizado", "")).strip()
        if apelido and apelido.lower() != "nan":
            return apelido
        return str(row.get("Descrição Oficial", "")).strip()

    df["Descrição"] = df.apply(obter_nome_final, axis=1)

    return df

@st.cache_data(ttl=15)
def carregar_pedidos():
    try:
        df_pedidos = conn.read(worksheet=WS_PEDIDOS, ttl=0)
    except ValueError as e:
        if "Spreadsheet must be specified" in str(e):
            st.error("🚨 **Erro Crítico:** URL da Planilha não especificada nas configurações.")
            st.stop()
        else:
            raise e
    except Exception as e:
        st.error(f"Erro ao conectar na aba {WS_PEDIDOS}: {e}")
        st.stop()

    df_cat = carregar_catalogo_materia()

    if df_pedidos.empty or "Código" not in df_pedidos.columns:
        df_init = df_cat[["Fornecedor", "Código", "Descrição"]].copy()
        for loja in LOJAS:
            df_init[loja] = "" 
        if not df_init.empty:
            conn.update(worksheet=WS_PEDIDOS, data=df_init)
        return df_init

    df_pedidos = df_pedidos.drop(columns=["Descrição", "Fornecedor"], errors="ignore")
    df_pedidos = pd.merge(df_cat[["Código", "Fornecedor", "Descrição"]], df_pedidos, on="Código", how="left")

    if "Código" in df_pedidos.columns:
        df_pedidos["Código"] = pd.to_numeric(df_pedidos["Código"], errors='coerce').fillna(0).astype(int)

    for loja in LOJAS:
        if loja in df_pedidos.columns:
            df_pedidos[loja] = df_pedidos[loja].fillna("").astype(str).apply(lambda x: "" if str(x).strip() in ["0", "0.0", "0,0"] else x)
        else:
            df_pedidos[loja] = ""

    return df_pedidos

def salvar_pedidos(df_to_save):
    conn.update(worksheet=WS_PEDIDOS, data=df_to_save)
    st.cache_data.clear()

def salvar_catalogo(df_to_save):
    df_clean = df_to_save.drop(columns=["Descrição"], errors="ignore")
    conn.update(worksheet=WS_PRODUTOS, data=df_clean)
    st.cache_data.clear()

# ─────────────────────────────────────────────────────────────
# MÓDULO PRINCIPAL - A FUNÇÃO QUE O MAIN.PY VAI CHAMAR
# ─────────────────────────────────────────────────────────────
def iniciar_tela():

    if 'reset_counter_materia_prima' not in st.session_state:
        st.session_state['reset_counter_materia_prima'] = 0

    usuario_atual = st.session_state.get('usuario_logado', None)
    acesso_total  = (usuario_atual == "Administrador")

    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@400;500;700&display=swap');

    :root {
        --bg-main:        #0d1117;
        --bg-card:        #161b22;
        --bg-sidebar:     #0d1117;
        --brown-dark:     #3a2610; 
        --brown-mid:      #8b5a2b; 
        --brown-accent:   #cd853f; 
        --brown-bright:   #eebb88; 
        --brown-glow:     rgba(205, 133, 63, .25);
        --text-primary:   #e6edf3;
        --text-muted:     #7d8590;
        --text-header:    #f5deb3; 
        --border:         #21262d;
        --border-active:  #cd853f;
        --row-hover:      rgba(205, 133, 63, .08);
        --row-selected:   rgba(205, 133, 63, .18);
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, var(--brown-mid) 0%, var(--brown-accent) 100%) !important;
        color: #fff !important; border: 1px solid var(--brown-accent) !important;
    }
    .stButton > button[kind="primary"]:hover {
        box-shadow: 0 4px 18px var(--brown-glow) !important;
    }

    .stTextInput input:focus, .stSelectbox > div > div:focus-within {
        border-color: var(--brown-accent) !important;
        box-shadow: 0 0 0 3px var(--brown-glow) !important;
    }

    .title-input input { color: var(--brown-bright) !important; }
    .title-input input:focus { border: 1px dashed var(--brown-accent) !important; }

    [data-testid="stDataEditor"] [data-testid="glideDataEditor"] .gdg-header-cell,
    [data-testid="stDataEditor"] .dvn-stack .gdg-header {
        background-color: var(--brown-dark) !important;
        color: var(--text-header) !important;
    }

    [data-testid="stDataEditor"] { border: 1px solid var(--brown-mid) !important; }
    [data-testid="stDataEditor"] .gdg-cell.gdg-selected,
    [data-testid="stDataEditor"] .gdg-cell[data-state="focused"],
    [data-testid="stDataEditor"] .gdg-cell[aria-selected="true"] {
        background-color: var(--row-selected) !important;
        outline: 2px solid var(--brown-accent) !important;
    }
    [data-testid="stDataEditor"] .gdg-row:hover .gdg-cell { background-color: var(--row-hover) !important; }

    div[data-testid="stVerticalBlockBorderWrapper"]:hover { border-color: var(--brown-mid) !important; }
    [data-testid="stMetricValue"] { color: var(--brown-bright) !important; }

    .topbar-loja {
        background: linear-gradient(90deg, var(--brown-dark) 0%, #1c140d 100%);
        border: 1px solid var(--brown-mid); border-radius: 10px; padding: 10px 18px; margin-bottom: 18px;
        display: flex; align-items: center; justify-content: space-between;
    }
    .topbar-left { display: flex; align-items: center; gap: 12px; }
    .topbar-title { font-size: 18px; font-weight: 700; color: var(--text-header); }
    .topbar-sub { font-size: 11px; color: var(--text-muted); margin-top: 2px; }
    .erp-badge { background-color: #2ea043; color: white; padding: 2px 8px; border-radius: 12px; font-size: 10px; font-weight: 600; margin-left: 8px;}

    /* IMPRESSÃO */
    @media print {
        @page { size: A4 portrait; margin: 8mm 8mm; }
        .stApp, .main, body, html { background-color: #ffffff !important; background-image: none !important; color: #000000 !important; padding: 0 !important; margin: 0 !important; }
        header, [data-testid="stSidebar"], [data-testid="stHeader"] { display: none !important; }
        [data-testid="stElementContainer"], [data-testid="stHorizontalBlock"], div[data-testid="stVerticalBlockBorderWrapper"] { display: none !important; }
        [data-testid="stElementContainer"]:has(#print-section) { display: block !important; width: 100% !important; }
        #print-section { display: block !important; width: 100% !important; }
        #print-section h2 { font-size: 13px !important; margin: 0 0 6px 0 !important; padding-bottom: 3px !important; border-bottom: 2px solid #000 !important; color: #000 !important; display: block !important; text-align: center !important; }
        #print-section h3 { font-size: 10px !important; font-weight: 700 !important; border-bottom: none !important; margin-top: 12px !important; margin-bottom: 3px !important; color: #000 !important; }
        .print-container { width: 100% !important; display: block !important; }
        table.print-table { width: 100% !important; border-collapse: collapse !important; color: #000000 !important; font-family: 'IBM Plex Sans', sans-serif !important; line-height: 1.2 !important; display: table !important; table-layout: fixed !important; margin-bottom: 4px !important; }
        table.print-table th, table.print-table td { border: 1px solid #444 !important; padding: 2px 3px !important; color: #000000 !important; background-color: #ffffff !important; overflow: hidden !important; text-overflow: ellipsis !important; }
        table.print-table td { white-space: nowrap !important; }
        table.print-table th { background-color: #d5d5d5 !important; font-weight: bold !important; text-align: center !important; white-space: normal !important; word-break: break-word !important; vertical-align: middle !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        table.print-table tr { break-inside: avoid !important; page-break-inside: avoid !important; }
        table.print-loja { font-size: 10px !important; }
        table.print-loja th:nth-child(1), table.print-loja td:nth-child(1) { width: 15% !important; text-align: left !important; }
        table.print-loja th:nth-child(2), table.print-loja td:nth-child(2) { width: 10% !important; text-align: center !important; }
        table.print-loja th:nth-child(3), table.print-loja td:nth-child(3) { width: 45% !important; text-align: left !important; }
        table.print-loja th:nth-child(4), table.print-loja td:nth-child(4) { width: 15% !important; text-align: center !important; }
        table.print-loja th:nth-child(5), table.print-loja td:nth-child(5) { width: 15% !important; text-align: center !important; font-weight: bold !important; background-color: #eeeeee !important; -webkit-print-color-adjust: exact !important; print-color-adjust: exact !important; }
        table.print-forn { font-size: 8px !important; }
        table.print-forn th:nth-child(1), table.print-forn td:nth-child(1) { width: 8% !important; text-align: center !important; }
        table.print-forn th:nth-child(2), table.print-forn td:nth-child(2) { width: 36% !important; text-align: left !important; }
        table.print-forn th:nth-child(n+3):nth-child(-n+10), table.print-forn td:nth-child(n+3):nth-child(-n+10) { width: 7% !important; text-align: center !important; }
        table.print-sep { font-size: 9px !important; }
        table.print-sep th:nth-child(1), table.print-sep td:nth-child(1) { width: 14% !important; text-align: left !important; }
        table.print-sep th:nth-child(2), table.print-sep td:nth-child(2) { width: 6%  !important; text-align: center !important; }
        table.print-sep th:nth-child(3), table.print-sep td:nth-child(3) { width: 24% !important; text-align: left !important; }
        table.print-sep th:nth-child(n+4):nth-child(-n+11), table.print-sep td:nth-child(n+4):nth-child(-n+11) { width: 7% !important; text-align: center !important; }
    }
    @media screen { #print-section { display: none !important; } }
    </style>
    """, unsafe_allow_html=True)

    if not acesso_total:
        st.markdown("""
        <style>
            section[data-testid="stSidebar"] { display: none !important; }
            [data-testid="collapsedControl"]  { display: none !important; }
            .main .block-container { max-width: 100% !important; padding-left: 2.5rem !important; padding-right: 2.5rem !important; }
        </style>
        """, unsafe_allow_html=True)

    @st.dialog("🚨 Confirmação Necessária")
    def modal_zerar_pedidos():
        st.markdown("Tem certeza que deseja **zerar todos os pedidos** de todas as lojas?")
        st.markdown("⚠️ *Esta ação irá zerar as quantidades diretamente no Google Sheets.*")
        st.write("<br>", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button("❌ Não, cancelar", use_container_width=True):
                st.rerun()
        with c2:
            if st.button("✔️ Sim, zerar tudo", type="primary", use_container_width=True):
                st.session_state['reset_counter_materia_prima'] += 1
                df_main = carregar_pedidos()
                for loja in LOJAS:
                    if loja in df_main.columns:
                        df_main[loja] = "" 
                salvar_pedidos(df_main)
                st.rerun()

    # ─────────────────────────────────────────────
    # SIDEBAR INTERNA
    # ─────────────────────────────────────────────
    with st.sidebar:
        st.markdown(f"### Módulo: Matéria Prima")
        if acesso_total:
            perfil_navegacao = st.radio("📍 Navegação:", [
                "Separação e Fechamento",
                "Visão das Lojas",
                "Visão por Fornecedor (Resumo)",
                "Catálogo de Produtos"
            ])
        else:
            perfil_navegacao = "Visão das Lojas"

        st.divider()

        df_ped = carregar_pedidos()
        if not df_ped.empty and set(LOJAS).issubset(df_ped.columns):
            temp_sum = df_ped[LOJAS].map(extrair_numero_quantidade)
            total_preenchidos = (temp_sum > 0).any(axis=1).sum()
        else:
            total_preenchidos = 0

        st.metric("Itens c/ pedido", total_preenchidos, help="Itens com ao menos 1 quantidade preenchida")
        st.divider()

    # ─────────────────────────────────────────────
    # ROTA 1: SEPARAÇÃO E FECHAMENTO
    # ─────────────────────────────────────────────
    if perfil_navegacao == "Separação e Fechamento":
        st.markdown("""
        <div class="page-header hide-print" style="background: linear-gradient(90deg, var(--brown-dark) 0%, #1c140d 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">📊</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Separação e Fechamento — Matéria Prima</div>
                <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">Consolidado geral de quantidades por Categoria e Código</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        with st.container(border=True):
            df_base = carregar_pedidos()
            if df_base.empty:
                st.warning("A base de pedidos está vazia. Cadastre produtos no Catálogo primeiro.")
                st.stop()

            col_cfg = {
                "Fornecedor":  st.column_config.TextColumn("Categoria", disabled=True),
                "Código":      st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                "Descrição":   st.column_config.TextColumn("Produto", disabled=True),
            }
            for loja, novo_nome in MAPA_LOJAS.items():
                col_cfg[loja] = st.column_config.TextColumn(novo_nome)

            cols_order = ["Fornecedor", "Código", "Descrição"] + LOJAS
            df_exibir = df_base[cols_order]

            df_editado = st.data_editor(
                df_exibir, hide_index=True, use_container_width=True, height=600,
                column_config=col_cfg, key=f"sep_editor_{st.session_state['reset_counter_materia_prima']}"
            )

            html_table = df_editado.to_html(index=False, classes=["print-table", "print-sep"])
            
            st.markdown(f"""<div id="print-section">
            <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
                Resumo de Separação — Matéria Prima
            </h2>
            <div class="print-container">{html_table}</div>
            </div>""", unsafe_allow_html=True)

            st.divider()
            col_salvar, col_csv, col_excel, col_print, col_zerar = st.columns([2.5, 1.2, 1.2, 1.5, 2.5])

            with col_salvar:
                if st.button("💾 Salvar Alterações", type="primary", use_container_width=True):
                    df_to_save = carregar_pedidos()
                    for _, row_edit in df_editado.iterrows():
                        mask = (df_to_save["Fornecedor"] == row_edit["Fornecedor"]) & (df_to_save["Código"] == row_edit["Código"])
                        for loja in LOJAS:
                            df_to_save.loc[mask, loja] = str(row_edit[loja])
                    salvar_pedidos(df_to_save)
                    st.success("✅ Pedidos salvos na nuvem com sucesso!")
                    st.rerun()

            with col_csv:
                df_csv = df_editado.copy().rename(columns=MAPA_LOJAS)
                df_csv["Observação:"] = ""
                csv = df_csv.to_csv(index=False).encode("utf-8")
                st.download_button("⬇️ CSV", data=csv, file_name="separacao_materia_prima.csv", mime="text/csv", use_container_width=True)

            with col_excel:
                df_exp = df_editado.copy().rename(columns=MAPA_LOJAS)
                df_exp["Observação:"] = ""
                excel_data = gerar_excel_estilizado(df_exp, "Separação Matéria Prima")
                st.download_button("⬇️ Excel", data=excel_data, file_name="separacao_materia_prima.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

            with col_print:
                if st.button("🖨️ Imprimir", use_container_width=True):
                    components.html(
                        "<script>"
                        "var s=document.createElement('style');"
                        "s.id='__sep_land__';"
                        "s.innerHTML='@media print{@page{size:A4 landscape!important;margin:6mm 10mm!important;}}';"
                        "window.parent.document.head.appendChild(s);"
                        "window.parent.print();"
                        "setTimeout(function(){var e=window.parent.document.getElementById('__sep_land__');if(e)e.remove();},3000);"
                        "</script>", height=0
                    )

            with col_zerar:
                if st.button("🚨 Zerar Todos os Pedidos", use_container_width=True, key="btn_zerar_sep_mat"):
                    modal_zerar_pedidos()

    # ─────────────────────────────────────────────
    # ROTA 2: VISÃO DAS LOJAS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Visão das Lojas":
        loja_selecionada = st.selectbox("👁️ Visualizar como:", LOJAS) if acesso_total else usuario_atual

        col_info, col_logout = st.columns([8, 2])
        with col_info:
            id_loja = MAPA_LOJAS.get(loja_selecionada, loja_selecionada)
            st.markdown(f"""
            <div class="topbar-loja hide-print">
                <div class="topbar-left">
                    <span style="font-size:22px">📦</span>
                    <div>
                        <div class="topbar-title">{loja_selecionada} — Matéria Prima <span class="erp-badge">🟢 Conectado ao ERP</span></div>
                        <div class="topbar-sub">Preencha a quantidade de cada produto</div>
                    </div>
                </div>
            </div>
            """, unsafe_allow_html=True)
        with col_logout:
            st.markdown("<div style='height:14px'></div>", unsafe_allow_html=True)
            if st.button("🏠 Voltar ao Início", use_container_width=True):
                st.session_state['modulo_ativo'] = 'Home'
                st.rerun()

        df_cat = carregar_catalogo_materia()
        df_cat_loja = df_cat[df_cat[loja_selecionada] == True].copy()

        if df_cat_loja.empty:
            st.warning(f"Nenhum produto habilitado para a {loja_selecionada} no momento.")
            st.stop()

        df_all = carregar_pedidos()
        df_loja_view = pd.merge(
            df_cat_loja[["Fornecedor", "Código", "Descrição"]],
            df_all[["Fornecedor", "Código", loja_selecionada]],
            on=["Fornecedor", "Código"],
            how="left"
        )
        
        df_loja_view[loja_selecionada] = df_loja_view[loja_selecionada].fillna("").astype(str)
        df_loja_view = df_loja_view.rename(columns={loja_selecionada: "Qtde"})

        try:
            conn_pg = st.connection("banco_erp", type="sql")
            mapa_banco_erp = {"Loja 01": "001", "Loja 02": "002", "Loja 03": "003", "Loja 04": "004", "Loja 05": "005", "Loja 06": "006", "Loja 07": "007", "Loja 08": "008"}
            cod_empresa_banco = mapa_banco_erp.get(loja_selecionada, "001")

            query_erp = f"""
                SELECT cade_codempresa, cade_codigo, cadp_descricao, estoque
                FROM "python_estoque" WHERE cade_codempresa::text = '{cod_empresa_banco}' ORDER BY cade_codempresa, cade_codigo
            """
            df_erp = conn_pg.query(query_erp, ttl=300)

            if not df_erp.empty:
                df_erp = df_erp.rename(columns={"cade_codigo": "Código", "estoque": "Estoque"})
                df_loja_view = pd.merge(df_loja_view, df_erp[["Código", "Estoque"]], on="Código", how="left")
            else:
                df_loja_view["Estoque"] = 0

        except Exception as e:
            df_loja_view["Estoque"] = 0

        df_loja_view["Estoque"] = df_loja_view["Estoque"].fillna(0).astype(int)
        df_loja_view = df_loja_view[["Fornecedor", "Código", "Descrição", "Estoque", "Qtde"]]

        with st.container(border=True):
            st.info("💡 **Dica:** O **Estoque** foi preenchido automaticamente com base no sistema ERP. Você pode preencher apenas a **Qtde** do pedido.")

            col_cfg_loja = {
                "Fornecedor": st.column_config.TextColumn("Categoria", width=150, disabled=True),
                "Código":     st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                "Descrição":  st.column_config.TextColumn("Produto", width=250, disabled=True),
                "Estoque":    st.column_config.NumberColumn("📦 Estoque", width=100, format="%d", disabled=True),
                "Qtde":       st.column_config.TextColumn("🛒 Qtde", width=120),
            }

            df_editado = st.data_editor(
                df_loja_view, column_config=col_cfg_loja, hide_index=True, use_container_width=True, height=520,
                key=f"loja_materia_prima_{st.session_state['reset_counter_materia_prima']}"
            )

        df_imprimir = df_editado.copy()
        df_imprimir = df_imprimir.rename(columns={"Estoque": "Est.", "Qtde": "Ped."})
        html_table_loja = df_imprimir.to_html(index=False, classes=["print-table", "print-loja"])

        st.markdown(f"""<div id="print-section">
        <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
            Resumo do Pedido — {loja_selecionada}
        </h2>
        <div class="print-container">{html_table_loja}</div>
        </div>""", unsafe_allow_html=True)

        df_editado_nums = df_editado["Qtde"].map(extrair_numero_quantidade)
        itens_com_pedido = int((df_editado_nums > 0).sum())
        total_itens      = len(df_editado)
        total_unidades   = int(df_editado_nums.sum())
        pct              = round(itens_com_pedido / total_itens * 100) if total_itens else 0

        st.divider()
        m1, m2, m3, col_print, col_btn = st.columns([2.5, 2.2, 1.8, 1.5, 3])
        with m1: st.metric("Itens preenchidos", f"{itens_com_pedido} / {total_itens}")
        with m2: st.metric("Total de unidades", total_unidades)
        with m3: st.metric("Cobertura", f"{pct}%")

        with col_print:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("🖨️ Imprimir", use_container_width=True):
                components.html("<script>window.parent.print();</script>", height=0)

        with col_btn:
            st.write("<br>", unsafe_allow_html=True)
            if st.button("💾 Salvar Pedido da Semana", type="primary", use_container_width=True):
                df_main = carregar_pedidos()
                for _, row in df_editado.iterrows():
                    mask = (df_main["Fornecedor"] == row["Fornecedor"]) & (df_main["Código"] == row["Código"])
                    if mask.any():
                        df_main.loc[mask, loja_selecionada] = str(row["Qtde"])
                    else:
                        nova_linha = {"Fornecedor": row["Fornecedor"], "Código": row["Código"], "Descrição": row["Descrição"]}
                        for l in LOJAS: nova_linha[l] = ""
                        nova_linha[loja_selecionada] = str(row["Qtde"])
                        df_main = pd.concat([df_main, pd.DataFrame([nova_linha])], ignore_index=True)

                salvar_pedidos(df_main)
                st.success(f"✅ Pedido da {loja_selecionada} enviado para a nuvem com sucesso!")

    # ─────────────────────────────────────────────
    # ROTA 3: VISÃO POR FORNECEDOR / RESUMO (Admin)
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Visão por Fornecedor (Resumo)":
        st.markdown("""
        <div class="hide-print" style="background: linear-gradient(90deg, var(--brown-dark) 0%, #1c140d 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">📦</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Visão por Categoria — Matéria Prima</div>
                <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">Resumo consolidado agrupado pelas categorias de produtos</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        df_all = carregar_pedidos()
        df_cat = carregar_catalogo_materia()

        if df_all.empty or df_cat.empty:
            st.warning("Não há dados de pedidos ou catálogo preenchidos.")
            st.stop()

        nomes_forn = df_cat["Fornecedor"].dropna().unique().tolist()
        html_print_content = ""

        for i in range(0, len(nomes_forn), 1):
            cols = st.columns(1, gap="small")
            for j, fornecedor in enumerate(nomes_forn[i:i+1]):

                df_forn = df_all[df_all["Fornecedor"] == fornecedor].copy()
                colunas_presentes = LOJAS
                df_forn = df_forn[["Código", "Descrição"] + colunas_presentes].copy()
                
                lojas_renomeadas = {l: MAPA_LOJAS[l] for l in colunas_presentes}
                df_forn = df_forn.rename(columns=lojas_renomeadas)

                lojas_cols_renomeadas = [MAPA_LOJAS[l] for l in colunas_presentes]

                col_cfg_forn = {
                    "Código":    st.column_config.NumberColumn("Cód.", width=80, format="%d", disabled=True),
                    "Descrição": st.column_config.TextColumn("Produto", disabled=False),
                }
                for c in lojas_cols_renomeadas:
                    col_cfg_forn[c] = st.column_config.TextColumn(c, disabled=False)

                altura = (len(df_forn) * 35) + 42

                with cols[j]:
                    with st.container(border=True):
                        st.markdown('<div class="title-input">', unsafe_allow_html=True)
                        st.text_input(
                            "Categoria", value=f"📦 {fornecedor}", label_visibility="collapsed",
                            key=f"title_forn_{fornecedor}_{st.session_state['reset_counter_materia_prima']}"
                        )
                        st.markdown('</div>', unsafe_allow_html=True)

                        cols_order_forn = ["Código", "Descrição"] + lojas_cols_renomeadas
                        df_forn_edit = st.data_editor(
                            df_forn[cols_order_forn], hide_index=True, use_container_width=True,
                            column_config=col_cfg_forn, height=altura, num_rows="fixed",
                            key=f"forn_materia_prima_{fornecedor}_{st.session_state['reset_counter_materia_prima']}"
                        )

                        df_forn_nums = df_forn[lojas_cols_renomeadas].map(extrair_numero_quantidade)
                        total_geral = int(df_forn_nums.sum().sum())
                        
                        st.markdown(f"""<div style="text-align:right; font-weight:700; margin-top:6px; color:var(--brown-bright); font-size:15px;">Total Geral: {total_geral} unidades</div>""", unsafe_allow_html=True)

                        html_table = df_forn_edit.to_html(index=False, classes=["print-table", "print-forn"])
                        for loja in LOJAS:
                            partes = loja.split(" ")
                            if len(partes) == 2:
                                html_table = html_table.replace(f"<th>{loja}</th>", f"<th>{partes[0]}<br>{partes[1]}</th>")
                        html_print_content += f"<h3 style='color: black; margin-top: 10px; margin-bottom: 4px;'>📦 {fornecedor}</h3>\n{html_table}\n<div style='text-align:right; font-weight:bold; font-size:11px; margin-top:3px; margin-bottom: 8px; color: black;'>Total da Categoria: {total_geral} unidades</div>\n"

        st.write("<br>", unsafe_allow_html=True)

        st.markdown(f"""<div id="print-section">
        <h2 style="color: black; margin-bottom: 10px; text-align: center; border-bottom: 2px solid black; padding-bottom: 5px;">
            Visão por Categoria (Resumo) — Matéria Prima
        </h2>
        <div class="print-container">{html_print_content}</div>
        </div>""", unsafe_allow_html=True)

        st.divider()
        
        col_csv, col_excel, col_space, col_print, col_zerar = st.columns([1.5, 1.5, 2.0, 2.5, 2.5])
        
        df_export = pd.DataFrame()
        for forn in nomes_forn:
            df_f = df_all[df_all["Fornecedor"] == forn].copy()
            if not df_f.empty:
                df_export = pd.concat([df_export, df_f], ignore_index=True)
                
        if not df_export.empty:
            cols_final_export = ["Código", "Descrição", "Fornecedor"] + LOJAS
            df_export = df_export[cols_final_export]
            df_export = df_export.rename(columns=MAPA_LOJAS)

        with col_csv:
            if not df_export.empty:
                df_export_csv = df_export.copy()
                df_export_csv["Observação:"] = ""
                csv_str = df_export_csv.to_csv(index=False).encode('utf-8')
                st.download_button("⬇️ Exportar CSV", data=csv_str, file_name="visao_categoria.csv", mime="text/csv", use_container_width=True)

        with col_excel:
            if not df_export.empty:
                df_export_excel = df_export.copy()
                df_export_excel["Observação:"] = ""
                excel_data = gerar_excel_estilizado(df_export_excel, "Visão Categorias")
                st.download_button("⬇️ Exportar Excel", data=excel_data, file_name="visao_categoria.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)

        with col_print:
            if st.button("🖨️ Imprimir Resumo Geral", use_container_width=True):
                components.html(
                    "<script>"
                    "var s=document.createElement('style');"
                    "s.id='__forn_port__';"
                    "s.innerHTML='@media print{@page{size:A4 portrait!important;margin:8mm 8mm!important;}}';"
                    "window.parent.document.head.appendChild(s);"
                    "window.parent.print();"
                    "setTimeout(function(){var e=window.parent.document.getElementById('__forn_port__');if(e)e.remove();},3000);"
                    "</script>", height=0
                )

        with col_zerar:
            if st.button("🚨 Zerar Todos os Pedidos", use_container_width=True, key="btn_zerar_forn_mat"):
                modal_zerar_pedidos()

    # ─────────────────────────────────────────────
    # ROTA 4: CATÁLOGO DE PRODUTOS
    # ─────────────────────────────────────────────
    elif perfil_navegacao == "Catálogo de Produtos":
        st.markdown("""
        <div class="page-header hide-print" style="background: linear-gradient(90deg, var(--brown-dark) 0%, #1c140d 100%); padding: 14px 20px; border-radius: 10px; margin-bottom: 22px;">
            <span style="font-size: 26px; margin-right: 12px;">🗂️</span>
            <div style="display: inline-block; vertical-align: top;">
                <div style="font-size: 20px; font-weight: 700; color: var(--text-header);">Catálogo de Matéria Prima</div>
                <div style="font-size: 12px; color: var(--text-muted); margin-top: 2px;">Atualize nomes direto do ERP ou crie apelidos personalizados. Os apelidos terão prioridade em todo o sistema.</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

        df_catalogo = carregar_catalogo_materia()
        df_editor_input = df_catalogo.drop(columns=["Descrição"], errors="ignore")

        # --- BLINDAGEM DE TIPOS ---
        df_editor_input["Código"] = pd.to_numeric(df_editor_input["Código"], errors='coerce').fillna(0).astype(int)
        
        cols_texto = ["Fornecedor", "Descrição Oficial", "Nome Personalizado"]
        for col in cols_texto:
            if col in df_editor_input.columns:
                df_editor_input[col] = df_editor_input[col].fillna("").astype(str).str.strip()

        for loja in LOJAS:
            if loja in df_editor_input.columns:
                df_editor_input[loja] = df_editor_input[loja].fillna(False).astype(bool)
        # --------------------------

        ordem_colunas = ["Fornecedor", "Código", "Descrição Oficial", "Nome Personalizado"] + LOJAS
        df_editor_input = df_editor_input[ordem_colunas]

        with st.container(border=True):
            opcoes_forn = ["MATÉRIA PRIMA GERAL", "USO DE PRODUTOS LIMPEZA", "TEMPEROS"]
            col_cfg_cat = {
                "Fornecedor":         st.column_config.SelectboxColumn("Categoria", options=opcoes_forn, width=200, required=True),
                "Código":             st.column_config.NumberColumn("Cód.", width=80, format="%d", required=True),
                "Descrição Oficial":  st.column_config.TextColumn("Nome Oficial (ERP)", width=280, disabled=False),
                "Nome Personalizado": st.column_config.TextColumn("Nome Personalizado (Apelido)", width=230),
            }
            for loja in LOJAS:
                col_cfg_cat[loja] = st.column_config.CheckboxColumn(loja, default=False)

            edited_cat = st.data_editor(
                df_editor_input, use_container_width=True, hide_index=True,
                height=600, num_rows="dynamic", column_config=col_cfg_cat,
                key=f"editor_catalogo_materia_prima_{st.session_state['reset_counter_materia_prima']}"
            )

            st.divider()
            col_atualizar, col_sync, _ = st.columns([2.5, 3, 3])

            with col_atualizar:
                if st.button("💾 Salvar Catálogo", type="primary", use_container_width=True):
                    salvar_catalogo(edited_cat)
                    st.session_state['reset_counter_materia_prima'] += 1
                    st.success("✅ Catálogo atualizado com sucesso!")
                    st.rerun()

            with col_sync:
                if st.button("📥 Puxar Nomes do ERP", use_container_width=True):
                    try:
                        conn_pg = st.connection("banco_erp", type="sql")
                        cods = tuple(edited_cat["Código"].tolist())
                        cods_str = f"({cods[0]})" if len(cods) == 1 else str(cods)
                        
                        query_nomes = f"SELECT cadp_codigo, cadp_descricao FROM cadprod WHERE cadp_codigo IN {cods_str}"
                        df_nomes = conn_pg.query(query_nomes, ttl=0)

                        for _, row in df_nomes.iterrows():
                            mask = edited_cat["Código"] == row["cadp_codigo"]
                            edited_cat.loc[mask, "Descrição Oficial"] = row["cadp_descricao"]

                        salvar_catalogo(edited_cat)
                        st.success("✅ Nomes Oficiais sincronizados com sucesso!")
                        st.rerun()
                    except Exception as e:
                        if "No database configured" in str(e) or "missing" in str(e).lower():
                             st.error("⚠️ Aviso: As credenciais de acesso ao PostgreSQL não foram encontradas no painel Secrets.")
                        else:
                             st.error(f"⚠️ Erro ao buscar nomes no banco ERP: {e}")
