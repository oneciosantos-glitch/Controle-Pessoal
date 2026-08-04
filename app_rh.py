# -*- coding: utf-8 -*-
"""Módulo Sistema de Compras e Entregas - Streamlit"""
import streamlit as st
import pandas as pd
import uuid
from datetime import datetime, date
import io
import base64

# ========== DADOS ==========
CLIENTES = ["Smart Fit", "Self Fit", "Assaí Atacadista"]

LOJAS_POR_CLIENTE = {
    "Smart Fit": [
        "Smart Fit Shopping Manoa", "Smart Fit Shopping Cidade Leste", "Smart Fit Macapá Shopping",
        "Smart Fit Shopping Grande Circular", "Smart Fit Shopping Via Norte", "Smart Fit Cidade Nova",
        "Smart Fit Parque Mosaico", "Smart Fit Cachoeirinha", "Smart Fit Flores", "Smart Fit Ponta Negra",
        "Smart Fit Nova Porto Velho", "Smart Fit Porto Velho Flodoaldo", "Smart Fit Alvorada",
        "Smart Fit Novo Aleixo", "Smart Fit São José do Operário", "Smart Fit Santana Macapá",
        "Smart Fit Toequato Tapajós"
    ],
    "Self Fit": [
        "Self Fit Hiper DB Ponta Negra", "Self Fit Manaus Plaza Shopping", "Self Fit Vieira Alves"
    ],
    "Assaí Atacadista": [
        "Assaí Atacadista Batista Campos", "Assaí Atacadista Almirante Barroso", "Assaí Atacadista Castanhal",
        "Assaí Atacadista Ananindeua", "Assaí Atacadista Augusto Monte Negro", "Assaí Atacadista Boa Vista",
        "Assaí Atacadista Manaus", "Assaí Atacadista Macapá", "Assaí Atacadista Belém"
    ]
}

MATERIAIS_POR_CLIENTE = {
    "Smart Fit": [
        "ÁGUA SANITÁRIA", "ASPIRADOR SEMI-INDUSTRIAL 23L", "BALDE 15L", "BALDE 6L",
        "BALDE ESPREMEDOR COMPLETO", "CABO DE ALUMINIO SEM ROSCA", "DISCO VERMELHO 510",
        "ENCERADEIRA INDUSTRIAL", "ESCOVA DE MÃO", "ESCOVA SANITÁRIA", "ESPONJA DUPLA FACE",
        "EXTENSÃO DE 30 M", "FIBRAS DE LIMPEZA PESADA", "FLANELAS", "KIT LIMPA VIDRO 2 EM 1 BRALIMPIA",
        "LIMPA TUDO (MINI LOK)", "PÁ DE LIXO COMPLETA", "PANO DE CHÃO ALVEJADO", "PLACAS SINALIZADORA",
        "REFIL MOP ÁGUA", "RODO DE 60cm COMPLETO", "SABÃO EM PÓ", "SACO DE LIXO 100 L",
        "SACO DE LIXO 40 L", "SACO DE LIXO 60 L", "VASSOURA DE NYLON COMPLETA", "VASSOURA DE TETO COMPLETA"
    ],
    "Self Fit": [
        "ASPIRADOR DE PÓ E BATERIA SEM FIO", "BAUDE EXPREMEDOR", "REFIL MOP ÁGUA", "REFIL MOP PÓ",
        "PLACAS SINALIZADORA", "ENCERADEIRA INDUSTRIAL", "KIT LIMPA VIDRO 2 EM 1 BRALIMPIA",
        "PÁ COLETORA DE LIXO", "PULVERIZADOR", "CESTA MULTIUSO"
    ],
    "Assaí Atacadista": [
        "Disco 510 mm - Preto", "Disco 510 mm - Marron para Remoção", "Disco 510 mm - vermelho",
        "Disco 510 mm - Verde", "Disco pelo de porco para Polidora", "Disco champanhe para Polidora",
        "Starlock frange 510mm aço", "Starlock frange 510mm com Velcon", "Starlock frange 510mm escova",
        "Enceradeira industrial 510", "Armação Mop Pó 60 cm", "Armação Mop Pó 1,20 cm", "Refil mop cera",
        "Suporte mop cera", "Lt", "Esponja p/LT", "Rodo madeira 1,20mts", "Rodo 60 cm",
        "Cabeleira Mop Agua", "Mop Pó 60 cm", "Mop Pó 1,20 cm", "Saco Amarelo P/Carrinho",
        "Pa coletora azul pop", "Raspador pesado sem cabo", "Raspador pesado com cabo", "Garra mop Agua",
        "Carro coletor de lixo 240lts", "Cabo de aluminio com rosca", "Cabo de aluminio sem rosca",
        "Lâminas p/ raspdor", "Raspador de mão", "Extenção cabo PP 3x2,5", "Borracha Organizadora Carrinho Funcional",
        "Carrinho funcional kit completo", "Balde expremedor", "Kit manunteção para carrinho", "Vassoura nylon",
        "Vassoura Piassava", "Vassourão gari 60 cm", "Alongador 9mts", "Regador"
    ]
}

EPIS_POR_CLIENTE = {
    "Smart Fit": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ],
    "Self Fit": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ],
    "Assaí Atacadista": [
        "Luva látex", "Óculos de proteção", "Luva de Vinil", "Máscara de Proteção",
        "Protetor auricular plug", "Protetor tipo concha", "Luva para jardineiro", "Avental de raspa",
        "Viseira", "Perneira", "Meia Térmica", "Japonha Térmica", "Calça Térmica", "Luvas térmicas",
        "Capuz Térmico", "Avental Térmico", "Bota C.Médio Nº34", "Bota C.Médio Nº35", "Bota C.Médio Nº36",
        "Bota C.Médio Nº37", "Bota C.Médio Nº38", "Bota C.Médio Nº39", "Bota C.Médio Nº40",
        "Bota C.Médio Nº41", "Bota C.Médio Nº42", "Bota C.Médio Nº43", "Bota C.Médio Nº44",
        "Bota C.Médio Nº45", "Bota C.Médio Nº46", "Bota de Couro Nº34", "Bota de Couro Nº35",
        "Bota de Couro Nº36", "Bota de Couro Nº37", "Bota de Couro Nº38", "Bota de Couro Nº39",
        "Bota de Couro Nº40", "Bota de Couro Nº41", "Bota de Couro Nº42", "Bota de Couro Nº43",
        "Bota de Couro Nº44", "Bota de Couro Nº45", "Bota de Couro Nº46", "Sapato Ant-derrapante Nº34",
        "Sapato Ant-derrapante Nº35", "Sapato Ant-derrapante Nº36", "Sapato Ant-derrapante Nº37",
        "Sapato Ant-derrapante Nº38", "Sapato Ant-derrapante Nº39", "Sapato Ant-derrapante Nº40",
        "Sapato Ant-derrapante Nº41", "Sapato Ant-derrapante Nº42", "Sapato Ant-derrapante Nº43",
        "Sapato Ant-derrapante Nº44", "Sapato Ant-derrapante Nº45", "Sapato Ant-derrapante Nº46",
        "Farda C.Feminino (P)", "Farda C.Feminino (M)", "Farda C.Feminino (G)", "Farda C.Feminino (GG)",
        "Farda C.Feminino (XG)", "Farda C.Masculino (P)", "Farda C.Masculino (M)", "Farda C.Masculino (G)",
        "Farda C.Masculino (GG)", "Farda C.Masculino (XG)", "Farda p/ Jardineiro", "Farda p/ Encarregado & Líder",
        "Farda p/ Supervisor", "Camisa Branca", "Cauça", "Chapéu"
    ]
}

STATUS_OPCOES = ["Pendente", "Aprovado", "Em Trânsito", "Entregue", "Cancelado"]
TIPOS_SOLICITACAO = ["Material", "EPI"]
PRIORIDADES = ["Normal", "Urgente", "Baixa"]

TAMANHOS_EPI = ["", "P", "M", "G", "GG", "37", "38", "39", "40", "41", "42", "43", "44"]


def gerar_id():
    return "SOL-" + uuid.uuid4().hex[:8].upper()


def formatar_data_br(d):
    if not d:
        return "-"
    if isinstance(d, str):
        if "T" in d:
            d = d.split("T")[0]
        try:
            y, m, day = d.split("-")
            return f"{day}/{m}/{y}"
        except Exception:
            return d
    if isinstance(d, (datetime, date)):
        return d.strftime("%d/%m/%Y")
    return str(d)


def formatar_moeda(v):
    if not v:
        return "R$ 0,00"
    try:
        return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)


def init_session_state():
    if "compras_solicitacoes" not in st.session_state:
        st.session_state["compras_solicitacoes"] = []
    if "compras_entregas" not in st.session_state:
        st.session_state["compras_entregas"] = []
    if "compras_page" not in st.session_state:
        st.session_state["compras_page"] = "Dashboard"
    if "compras_edit_id" not in st.session_state:
        st.session_state["compras_edit_id"] = None
    if "compras_nova_itens_material" not in st.session_state:
        st.session_state["compras_nova_itens_material"] = []
    if "compras_nova_itens_epi" not in st.session_state:
        st.session_state["compras_nova_itens_epi"] = []
    if "compras_form_reset" not in st.session_state:
        st.session_state["compras_form_reset"] = False


def switch_page(page):
    st.session_state["compras_page"] = page
    st.session_state["compras_edit_id"] = None
    st.rerun()


def get_badge_color(status):
    return {
        "Pendente": "🔴",
        "Aprovado": "🟢",
        "Em Trânsito": "🔵",
        "Entregue": "🟢",
        "Cancelado": "⚫"
    }.get(status, "⚪")


# ========== GERAÇÃO DE DOCUMENTOS ==========
def gerar_doc_epi(sol):
    data_geracao = datetime.now().strftime("%d/%m/%Y")
    cliente = sol.get("cliente", "_________________")
    loja = sol.get("loja", "_________________")
    map_itens = {i.get("epi", ""): i for i in sol.get("itens", [])}

    epi_list = [
        "Luva látex", "Luva para jardineiro", "Japona Térmica",
        "Óculos de proteção", "Avental de raspa", "Calça Térmica",
        "Máscara de Proteção", "Viseira", "Luvas térmicas",
        "Protetor auricular plug ( ) ou concha ( )", "Perneira", "Cap", "", "", ""
    ]
    epi_rows = []
    for i in range(0, len(epi_list), 3):
        e1 = epi_list[i] if i < len(epi_list) else ""
        e2 = epi_list[i+1] if i+1 < len(epi_list) else ""
        e3 = epi_list[i+2] if i+2 < len(epi_list) else ""
        q1 = map_itens.get(e1, {}).get("qtd", "&nbsp;") if e1 else "&nbsp;"
        q2 = map_itens.get(e2, {}).get("qtd", "&nbsp;") if e2 else "&nbsp;"
        q3 = map_itens.get(e3, {}).get("qtd", "&nbsp;") if e3 else "&nbsp;"
        epi_rows.append(
            f'<tr><td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e1 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q1}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e2 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q2}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{e3 or "&nbsp;"}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px;text-align:center">{q3}</td></tr>'
        )

    itens_arr = [(i.get("epi", i.get("nome", "Item")) + (f' x{i.get("qtd", "")}' if i.get("qtd") else "")) for i in sol.get("itens", [])]
    botas_rows = []
    total_rows = max(17, len(itens_arr))
    for i in range(total_rows):
        item_obj = sol.get("itens", [])[i] if i < len(sol.get("itens", [])) else None
        nome = item_obj.get("colaborador", sol.get("nomeFuncionario", "&nbsp;")) if item_obj else "&nbsp;"
        loja_val = sol.get("loja", "&nbsp;") if item_obj else "&nbsp;"
        enc = sol.get("encarregado", "&nbsp;") if item_obj else "&nbsp;"
        sup = sol.get("supervisor", "&nbsp;") if item_obj else "&nbsp;"
        item = itens_arr[i] if i < len(itens_arr) else "&nbsp;"
        botas_rows.append(
            f'<tr><td style="border:1px solid #000;padding:3px 5px;font-size:10px">{nome}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{loja_val}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{enc}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{sup}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">{item}</td>'
            f'<td style="border:1px solid #000;padding:3px 5px;font-size:10px">&nbsp;</td></tr>'
        )

    states = 'ALAGOAS ( &nbsp;) &nbsp;&nbsp;BAHIA ( &nbsp;) &nbsp;&nbsp;CEARÁ ( &nbsp;) &nbsp;&nbsp;MARANHÃO ( &nbsp;) &nbsp;&nbsp;PARAIBA ( &nbsp;) &nbsp;&nbsp;PARÁ ( &nbsp;) &nbsp;&nbsp;PERNAMBUCO ( &nbsp;) &nbsp;&nbsp;PIAUÍ ( &nbsp;) &nbsp;&nbsp;RIO GRANDE DO NORTE ( &nbsp;) &nbsp;&nbsp;SERGIPE ( &nbsp;) &nbsp;&nbsp;AMAPÁ ( &nbsp;) &nbsp;&nbsp;RORAIMA ( &nbsp;) &nbsp;&nbsp;AMAZONAS ( &nbsp;) &nbsp;&nbsp;RONDÔNIA ( &nbsp;)'

    html = f"""<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset='utf-8'><title>Solicitacao de EPI - {cliente} - {loja}</title>
<style>@page Section1 {{ size: 841.95pt 595.35pt; margin: 36pt 36pt 36pt 36pt; mso-page-orientation: landscape; }}
div.Section1 {{ page: Section1; }}</style></head>
<body style="font-family:Arial,sans-serif;font-size:10pt;margin:0;padding:0">
<div class="Section1" style="mso-page-orientation: landscape;">
<table style="width:100%;border-collapse:collapse">
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:9px;text-align:center">{states}</td></tr>
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:11px;font-weight:bold;background:#DEEBF6;text-align:center">BOTAS</td></tr>
<tr>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Nome</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Loja</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Encarregado</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Supervisor</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Itens da Solicitação de EPI</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Situação ( * )</th>
</tr>
{''.join(botas_rows)}
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:11px;font-weight:bold;background:#DEEBF6;text-align:center">EPIS</td></tr>
<tr>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Equipamento de Proteção</th>
<th style="border:1px solid #000;padding:3px 5px;font-size:10px;font-weight:bold;background:#DEEBF6">Quantidade</th>
</tr>
{''.join(epi_rows)}
<tr><td colspan="6" style="border:1px solid #000;padding:3px 5px;font-size:9px">Observação: A coluna Situação ( * ), é para o setor SST preencher.</td></tr>
</table>
<p style="font-size:8pt;color:#666;text-align:center;margin-top:6px">Data de geração: {data_geracao}</p>
</div></body></html>"""
    return html


def gerar_xls_material(sol):
    data_geracao = datetime.now().strftime("%d/%m/%Y")
    cliente = sol.get("cliente", "")
    itens = sol.get("itens", [])

    def build_rows_smart_self(items):
        if not items:
            return '<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">UN.</td></tr>'
        return "".join([
            f'<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">{i.get("material","")}</td>'
            f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{i.get("qtd","&nbsp;")}</td>'
            f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">UN.</td></tr>'
            for i in items
        ])

    def build_rows_assai(items):
        if not items:
            return '<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">1</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center">peças</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right">&nbsp;</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right">&nbsp;</td></tr>'
        rows = []
        for idx, i in enumerate(items, 1):
            qtd = i.get("qtd", "&nbsp;")
            unit = f'R$ {i["valorUnit"]:.2f}'.replace(".", ",") if i.get("valorUnit") else "&nbsp;"
            total = f'R$ {(i["valorUnit"]*i["qtd"]):.2f}'.replace(".", ",") if i.get("valorUnit") and i.get("qtd") else "&nbsp;"
            rows.append(
                f'<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{idx}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">{i.get("material","")}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;white-space:nowrap">&nbsp;</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">{qtd}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:center;white-space:nowrap">peças</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right;white-space:nowrap">{unit}</td>'
                f'<td style="border:1px solid #000;padding:4px 6px;font-size:11px;text-align:right;white-space:nowrap">{total}</td></tr>'
            )
        return "".join(rows)

    if cliente == "Smart Fit":
        rows = build_rows_smart_self(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body><table>
<tr><th colspan="3" style="border:1px solid #000;padding:6px;font-size:13px;font-weight:bold;background:#DEEBF6;text-align:center">SMART FIT</th></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">SMART FIT</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">QTD</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">REF.</th></tr>
{rows}
<tr><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Pedido_Material_SmartFit_{sol.get("id","")}.xls'
    elif cliente == "Self Fit":
        rows = build_rows_smart_self(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body><table>
<tr><th colspan="3" style="border:1px solid #000;padding:6px;font-size:13px;font-weight:bold;background:#DEEBF6;text-align:center">SELF FIT</th></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">SELF FIT</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">QTD</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">REF.</th></tr>
{rows}
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">DATA DO PEDIDO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">MÊS DE REFERÊNCIA</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">SEPARADO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">ENVIADO</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Pedido_Material_SelfFit_{sol.get("id","")}.xls'
    elif cliente == "Assaí Atacadista":
        item_rows = build_rows_assai(itens)
        html = f"""<html xmlns:x="urn:schemas-microsoft-com:office:excel"><head><meta charset="utf-8"><style>table{{border-collapse:collapse;width:100%}}th,td{{border:1px solid #000}}</style></head><body>
<table style="width:100%;border-collapse:collapse">
<tr><td colspan="7" style="border:1px solid #000;padding:6px;font-size:11px;text-align:center;font-weight:bold">R: Sgto Jeter Augusto Pereira N° 02 e 04 - São Paulo - CEP: 02188-070 - E-mail: vendas@thamesjlara.com.br - Site www.thamesjlara.com.br</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Orçamento</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">At.: Sr.(a): Mendonça</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Data: {data_geracao}</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Vendedor: Hélio</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Razão Social</td><td colspan="3" style="border:1px solid #000;padding:4px 6px;font-size:11px">FG Services Eireli - ME</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">CNPJ/CPF</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">23.585.374/0001-11</td></tr>
<tr><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Endereço</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">Av. Barão de Vera Cruz, 586 BR 101 Norte</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Bairro</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">Cruz de Rebouças</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Cidade</td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px">Igarassu</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">UF</td><td style="border:1px solid #000;padding:4px 6px;font-size:11px">PE</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">CEP</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">53635-015</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><td style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">Telefone</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">(0xx81) 3545-3990</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold">E-mail</td><td colspan="2" style="border:1px solid #000;padding:4px 6px;font-size:11px">&nbsp;</td></tr>
<tr><td colspan="7" style="height:10px"></td></tr>
<tr><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Item</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Descrição</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Marca</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Qtde</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Unid</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Valor Unit.</th><th style="border:1px solid #000;padding:4px 6px;font-size:11px;font-weight:bold;background:#DEEBF6;white-space:nowrap">Total</th></tr>
{item_rows}
<tr><td colspan="7" style="border:1px solid #000;padding:4px 6px;font-size:10px">Solicitação: {sol.get("id","")} | Loja: {sol.get("loja","")} | Data: {data_geracao}</td></tr>
</table></body></html>"""
        filename = f'Orcamento_Assai_{sol.get("id","")}.xls'
    else:
        return None, None
    return html, filename


# ========== PÁGINAS ==========
def page_dashboard():
    st.markdown("### 📊 Dashboard")
    sols = st.session_state["compras_solicitacoes"]
    ents = st.session_state["compras_entregas"]

    total = len(sols)
    pendentes = sum(1 for s in sols if s.get("status") == "Pendente")
    transito = sum(1 for s in sols if s.get("status") == "Em Trânsito")
    entregues = sum(1 for s in sols if s.get("status") == "Entregue")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Solicitações", total)
    c2.metric("Pendentes", pendentes)
    c3.metric("Em Trânsito", transito)
    c4.metric("Entregues", entregues)

    st.markdown("---")
    st.markdown("#### 📋 Últimas Solicitações")
    if sols:
        recentes = sorted(sols, key=lambda x: x.get("dataCriacao", ""), reverse=True)[:10]
        df = pd.DataFrame([
            {
                "ID": s["id"],
                "Data": formatar_data_br(s.get("data")),
                "Loja": s.get("loja", ""),
                "Cliente": s.get("cliente", ""),
                "Tipo": s.get("tipo", ""),
                "Solicitante": s.get("solicitante", ""),
                "Itens": len(s.get("itens", [])),
                "Valor": formatar_moeda(s.get("valorTotal", 0)),
                "Status": f"{get_badge_color(s.get('status'))} {s.get('status', '')}"
            }
            for s in recentes
        ])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma solicitação cadastrada.")

    if st.button("➕ Nova Solicitação", type="primary"):
        switch_page("Nova Solicitação")


def page_solicitacoes():
    st.markdown("### 📋 Todas as Solicitações")

    sols = st.session_state["compras_solicitacoes"]

    # Filtros
    with st.expander("🔍 Filtros", expanded=True):
        c1, c2, c3, c4, c5, c6 = st.columns(6)
        with c1:
            busca = st.text_input("Buscar", placeholder="ID, loja, solicitante...", key="sol_busca")
        with c2:
            fstatus = st.selectbox("Status", ["Todos"] + STATUS_OPCOES, key="sol_fstatus")
        with c3:
            floja = st.selectbox("Loja", ["Todas"] + sorted(list(set(l for c in LOJAS_POR_CLIENTE.values() for l in c))), key="sol_floja")
        with c4:
            fcliente = st.selectbox("Cliente", ["Todos"] + CLIENTES, key="sol_fcliente")
        with c5:
            fdi = st.date_input("Data Início", value=None, key="sol_fdi")
        with c6:
            fdf = st.date_input("Data Fim", value=None, key="sol_fdf")

    filtradas = []
    for s in sols:
        txt = f"{s.get('id','')} {s.get('loja','')} {s.get('solicitante','')}".lower()
        if busca and busca.lower() not in txt:
            continue
        if fstatus != "Todos" and s.get("status") != fstatus:
            continue
        if floja != "Todas" and s.get("loja") != floja:
            continue
        if fcliente != "Todos" and s.get("cliente") != fcliente:
            continue
        if fdi:
            try:
                if s.get("data") and str(s["data"]) < str(fdi):
                    continue
            except Exception:
                pass
        if fdf:
            try:
                if s.get("data") and str(s["data"]) > str(fdf):
                    continue
            except Exception:
                pass
        filtradas.append(s)

    if filtradas:
        for s in filtradas:
            with st.container(border=True):
                col1, col2, col3 = st.columns([4, 2, 2])
                with col1:
                    st.write(f"**{s['id']}** | {formatar_data_br(s.get('data'))} | {s.get('loja','')} | {s.get('cliente','')}")
                    st.caption(f"Tipo: `{s.get('tipo','')}` | Solicitante: {s.get('solicitante','')} | Itens: {len(s.get('itens',[]))} | Valor: {formatar_moeda(s.get('valorTotal',0))}")
                with col2:
                    st.write(f"{get_badge_color(s.get('status'))} **{s.get('status','')}**")
                with col3:
                    c_a, c_b, c_c = st.columns(3)
                    with c_a:
                        if st.button("👁️", key=f"ver_{s['id']}"):
                            st.session_state["compras_ver_id"] = s["id"]
                            st.session_state["compras_page"] = "Detalhes"
                            st.rerun()
                    with c_b:
                        if st.button("✏️", key=f"edit_{s['id']}"):
                            st.session_state["compras_edit_id"] = s["id"]
                            st.session_state["compras_page"] = "Nova Solicitação"
                            st.rerun()
                    with c_c:
                        if st.button("🗑️", key=f"del_{s['id']}"):
                            st.session_state["compras_solicitacoes"] = [x for x in sols if x["id"] != s["id"]]
                            st.session_state["compras_entregas"] = [e for e in st.session_state["compras_entregas"] if e.get("idSolicitacao") != s["id"]]
                            st.success("Excluído!")
                            st.rerun()

        # Exportar CSV
        if filtradas:
            csv_data = []
            for s in filtradas:
                csv_data.append([
                    s["id"], s.get("data",""), s.get("loja",""), s.get("cliente",""),
                    s.get("tipo",""), s.get("solicitante",""), len(s.get("itens",[])),
                    s.get("valorTotal",0), s.get("status","")
                ])
            df_csv = pd.DataFrame(csv_data, columns=["ID","Data","Loja","Cliente","Tipo","Solicitante","Qtd Itens","Valor Total","Status"])
            csv_buffer = io.StringIO()
            df_csv.to_csv(csv_buffer, index=False, sep=";", encoding="utf-8")
            st.download_button("📥 Exportar CSV", data=csv_buffer.getvalue().encode("utf-8"), file_name="solicitacoes.csv", mime="text/csv")
    else:
        st.info("Nenhuma solicitação encontrada.")


def page_nova_solicitacao():
    st.markdown("### ➕ Nova Solicitação de Compra")

    edit_id = st.session_state.get("compras_edit_id")
    edit_sol = None
    if edit_id:
        for s in st.session_state["compras_solicitacoes"]:
            if s["id"] == edit_id:
                edit_sol = s
                break

    with st.form("form_nova_solicitacao"):
        c1, c2, c3 = st.columns(3)
        with c1:
            cliente = st.selectbox("Cliente *", CLIENTES, index=CLIENTES.index(edit_sol["cliente"]) if edit_sol and edit_sol.get("cliente") in CLIENTES else 0)
        with c2:
            lojas = LOJAS_POR_CLIENTE.get(cliente, [])
            idx_loja = lojas.index(edit_sol["loja"]) if edit_sol and edit_sol.get("loja") in lojas else 0
            loja = st.selectbox("Loja *", lojas, index=idx_loja)
        with c3:
            tipo = st.selectbox("Tipo *", TIPOS_SOLICITACAO, index=TIPOS_SOLICITACAO.index(edit_sol["tipo"]) if edit_sol and edit_sol.get("tipo") in TIPOS_SOLICITACAO else 0)

        c4, c5, c6 = st.columns(3)
        with c4:
            solicitante = st.text_input("Solicitante *", value=edit_sol.get("solicitante","") if edit_sol else "")
        with c5:
            data_sol = st.date_input("Data", value=datetime.strptime(edit_sol["data"], "%Y-%m-%d").date() if edit_sol and edit_sol.get("data") else date.today())
        with c6:
            prioridade = st.selectbox("Prioridade", PRIORIDADES, index=PRIORIDADES.index(edit_sol["prioridade"]) if edit_sol and edit_sol.get("prioridade") in PRIORIDADES else 0)

        previsao = st.date_input("Previsão de Entrega", value=datetime.strptime(edit_sol["previsao"], "%Y-%m-%d").date() if edit_sol and edit_sol.get("previsao") else date.today())
        observacoes = st.text_area("Observações", value=edit_sol.get("observacoes","") if edit_sol else "")

        # Campos específicos EPI
        if tipo == "EPI":
            st.markdown("---")
            st.markdown("#### 👷 Informações EPI")
            c7, c8, c9 = st.columns(3)
            with c7:
                nome_func = st.text_input("Nome do Funcionário", value=edit_sol.get("nomeFuncionario","") if edit_sol else "")
            with c8:
                encarregado = st.text_input("Encarregado", value=edit_sol.get("encarregado","") if edit_sol else "")
            with c9:
                supervisor = st.text_input("Supervisor", value=edit_sol.get("supervisor","") if edit_sol else "")
            data_bota = st.date_input("Data Última Bota", value=datetime.strptime(edit_sol["dataUltimaBota"], "%Y-%m-%d").date() if edit_sol and edit_sol.get("dataUltimaBota") else None)

        st.markdown("---")
        submitted = st.form_submit_button("💾 Salvar Solicitação", type="primary")

    # Itens (fora do form para permitir adicionar dinamicamente)
    st.markdown("#### 📦 Itens")

    if tipo == "Material":
        st.markdown("**Materiais**")
        materiais = MATERIAIS_POR_CLIENTE.get(cliente, [])

        # Inicializar itens do edit
        if edit_sol and not st.session_state.get("compras_edit_loaded"):
            st.session_state["compras_nova_itens_material"] = [
                {"material": i.get("material",""), "qtd": i.get("qtd",1), "valorUnit": i.get("valorUnit",0)}
                for i in edit_sol.get("itens", [])
            ]
            st.session_state["compras_edit_loaded"] = True

        itens_mat = st.session_state.get("compras_nova_itens_material", [])

        for idx, item in enumerate(itens_mat):
            cols = st.columns([3, 1, 1, 1])
            with cols[0]:
                item["material"] = st.selectbox(f"Material {idx+1}", materiais, index=materiais.index(item["material"]) if item.get("material") in materiais else 0, key=f"mat_sel_{idx}")
            with cols[1]:
                item["qtd"] = st.number_input(f"Qtd {idx+1}", min_value=1, value=int(item.get("qtd", 1)), key=f"mat_qtd_{idx}")
            with cols[2]:
                v = item.get("valorUnit", 0)
                item["valorUnit"] = st.number_input(f"Valor {idx+1}", min_value=0.0, value=float(v), step=0.01, format="%.2f", key=f"mat_val_{idx}")
            with cols[3]:
                st.write("")
                st.write("")
                if st.button("🗑️", key=f"mat_rem_{idx}"):
                    itens_mat.pop(idx)
                    st.session_state["compras_nova_itens_material"] = itens_mat
                    st.rerun()

        if st.button("➕ Adicionar Material"):
            itens_mat.append({"material": materiais[0] if materiais else "", "qtd": 1, "valorUnit": 0})
            st.session_state["compras_nova_itens_material"] = itens_mat
            st.rerun()

    elif tipo == "EPI":
        st.markdown("**EPIs**")
        epis = EPIS_POR_CLIENTE.get(cliente, EPIS_POR_CLIENTE.get("Smart Fit", []))

        if edit_sol and not st.session_state.get("compras_edit_loaded"):
            st.session_state["compras_nova_itens_epi"] = [
                {"epi": i.get("epi",""), "colaborador": i.get("colaborador",""), "qtd": i.get("qtd",1), "tamanho": i.get("tamanho","")}
                for i in edit_sol.get("itens", [])
            ]
            st.session_state["compras_edit_loaded"] = True

        itens_epi = st.session_state.get("compras_nova_itens_epi", [])

        for idx, item in enumerate(itens_epi):
            cols = st.columns([2, 1, 1, 1, 1])
            with cols[0]:
                item["epi"] = st.selectbox(f"EPI {idx+1}", epis, index=epis.index(item["epi"]) if item.get("epi") in epis else 0, key=f"epi_sel_{idx}")
            with cols[1]:
                item["colaborador"] = st.text_input(f"Colab {idx+1}", value=item.get("colaborador",""), key=f"epi_col_{idx}")
            with cols[2]:
                item["qtd"] = st.number_input(f"Qtd {idx+1}", min_value=1, value=int(item.get("qtd", 1)), key=f"epi_qtd_{idx}")
            with cols[3]:
                tams = TAMANHOS_EPI
                item["tamanho"] = st.selectbox(f"Tam {idx+1}", tams, index=tams.index(item["tamanho"]) if item.get("tamanho") in tams else 0, key=f"epi_tam_{idx}")
            with cols[4]:
                st.write("")
                st.write("")
                if st.button("🗑️", key=f"epi_rem_{idx}"):
                    itens_epi.pop(idx)
                    st.session_state["compras_nova_itens_epi"] = itens_epi
                    st.rerun()

        if st.button("➕ Adicionar EPI"):
            itens_epi.append({"epi": epis[0] if epis else "", "colaborador": "", "qtd": 1, "tamanho": ""})
            st.session_state["compras_nova_itens_epi"] = itens_epi
            st.rerun()

    # Ao salvar
    if submitted:
        itens = []
        valor_total = 0
        if tipo == "Material":
            itens = st.session_state.get("compras_nova_itens_material", [])
            valor_total = sum(i.get("valorUnit", 0) * i.get("qtd", 0) for i in itens)
            if not itens:
                st.error("Adicione pelo menos um material.")
                return
        elif tipo == "EPI":
            itens = st.session_state.get("compras_nova_itens_epi", [])
            if not itens:
                st.error("Adicione pelo menos um EPI.")
                return

        nova_sol = {
            "id": edit_sol["id"] if edit_sol else gerar_id(),
            "data": str(data_sol),
            "cliente": cliente,
            "loja": loja,
            "tipo": tipo,
            "solicitante": solicitante,
            "nomeFuncionario": nome_func if tipo == "EPI" else "",
            "encarregado": encarregado if tipo == "EPI" else "",
            "supervisor": supervisor if tipo == "EPI" else "",
            "dataUltimaBota": str(data_bota) if tipo == "EPI" and data_bota else "",
            "prioridade": prioridade,
            "previsao": str(previsao),
            "observacoes": observacoes,
            "itens": itens,
            "valorTotal": valor_total,
            "anexos": [],
            "status": edit_sol.get("status", "Pendente") if edit_sol else "Pendente",
            "dataCriacao": edit_sol.get("dataCriacao", datetime.now().isoformat()) if edit_sol else datetime.now().isoformat()
        }

        # Gera documentos
        if tipo == "EPI":
            html_doc = gerar_doc_epi(nova_sol)
            nova_sol["anexos"].append({
                "nome": f'Solicitacao_de_EPI_{nova_sol["id"]}_{cliente}_{loja}.doc',
                "conteudo": html_doc,
                "tipo": "doc"
            })
        elif tipo == "Material":
            html_xls, filename_xls = gerar_xls_material(nova_sol)
            if html_xls:
                nova_sol["anexos"].append({
                    "nome": filename_xls,
                    "conteudo": html_xls,
                    "tipo": "xls"
                })

        # Atualiza lista
        sols = st.session_state["compras_solicitacoes"]
        if edit_sol:
            sols = [s for s in sols if s["id"] != edit_id]
        sols.append(nova_sol)
        st.session_state["compras_solicitacoes"] = sols

        # Atualiza entregas
        if not edit_sol:
            st.session_state["compras_entregas"].append({
                "idSolicitacao": nova_sol["id"],
                "loja": loja,
                "tipo": tipo,
                "dataEnvio": "",
                "dataPrevista": str(previsao),
                "dataEntrega": "",
                "transportadora": "",
                "rastreio": "",
                "status": "Pendente",
                "observacoes": ""
            })

        # Limpa estado
        st.session_state["compras_nova_itens_material"] = []
        st.session_state["compras_nova_itens_epi"] = []
        st.session_state["compras_edit_id"] = None
        st.session_state["compras_edit_loaded"] = False
        st.success("Solicitação salva com sucesso!")
        st.session_state["compras_page"] = "Solicitações"
        st.rerun()


def page_detalhes():
    ver_id = st.session_state.get("compras_ver_id")
    s = None
    for sol in st.session_state["compras_solicitacoes"]:
        if sol["id"] == ver_id:
            s = sol
            break
    if not s:
        st.error("Solicitação não encontrada.")
        return

    st.markdown(f"### 👁️ Detalhes da Solicitação {s['id']}")

    col1, col2 = st.columns(2)
    with col1:
        st.write(f"**Data:** {formatar_data_br(s.get('data'))}")
        st.write(f"**Cliente:** {s.get('cliente','')}")
        st.write(f"**Loja:** {s.get('loja','')}")
        st.write(f"**Tipo:** {s.get('tipo','')}")
        st.write(f"**Solicitante:** {s.get('solicitante','')}")
    with col2:
        st.write(f"**Prioridade:** {s.get('prioridade','')}")
        st.write(f"**Status:** {get_badge_color(s.get('status'))} {s.get('status','')}")
        st.write(f"**Previsão:** {formatar_data_br(s.get('previsao'))}")
        st.write(f"**Valor Total:** {formatar_moeda(s.get('valorTotal',0))}")

    if s.get("observacoes"):
        st.write(f"**Observações:** {s['observacoes']}")

    if s.get("tipo") == "EPI":
        st.markdown("---")
        st.markdown("#### 👷 Informações EPI")
        st.write(f"**Funcionário:** {s.get('nomeFuncionario','-')}")
        st.write(f"**Encarregado:** {s.get('encarregado','-')}")
        st.write(f"**Supervisor:** {s.get('supervisor','-')}")
        st.write(f"**Data Última Bota:** {formatar_data_br(s.get('dataUltimaBota'))}")

    st.markdown("---")
    st.markdown("#### 📦 Itens")
    if s.get("itens"):
        if s["tipo"] == "Material":
            df = pd.DataFrame([{"Material": i.get("material",""), "Qtd": i.get("qtd",1), "Valor Unit.": formatar_moeda(i.get("valorUnit",0)), "Total": formatar_moeda(i.get("valorUnit",0)*i.get("qtd",0))} for i in s["itens"]])
        else:
            df = pd.DataFrame([{"EPI": i.get("epi",""), "Colaborador": i.get("colaborador",""), "Qtd": i.get("qtd",1), "Tamanho": i.get("tamanho","-")} for i in s["itens"]])
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum item.")

    # Anexos
    if s.get("anexos"):
        st.markdown("---")
        st.markdown("#### 📎 Anexos Gerados")
        for a in s["anexos"]:
            st.download_button(
                label=f"📥 {a['nome']}",
                data=a["conteudo"].encode("utf-8"),
                file_name=a["nome"],
                mime="application/octet-stream",
                key=f"anexo_{a['nome']}"
            )

    if st.button("🔙 Voltar"):
        st.session_state["compras_page"] = "Solicitações"
        st.rerun()


def page_entregas():
    st.markdown("### 🚚 Controle de Entregas")

    ents = st.session_state["compras_entregas"]

    c1, c2 = st.columns(2)
    with c1:
        busca = st.text_input("Buscar", placeholder="ID, loja...", key="ent_busca")
    with c2:
        fstatus = st.selectbox("Status", ["Todos"] + STATUS_OPCOES, key="ent_fstatus")

    filtradas = []
    for e in ents:
        txt = f"{e.get('idSolicitacao','')} {e.get('loja','')}".lower()
        if busca and busca.lower() not in txt:
            continue
        if fstatus != "Todos" and e.get("status") != fstatus:
            continue
        filtradas.append(e)

    if filtradas:
        for e in filtradas:
            with st.container(border=True):
                col1, col2 = st.columns([3, 2])
                with col1:
                    st.write(f"**{e.get('idSolicitacao','')}** | {e.get('loja','')} | Tipo: {e.get('tipo','')}")
                    st.caption(f"Transportadora: {e.get('transportadora','-')} | Rastreio: {e.get('rastreio','-')}")
                with col2:
                    st.write(f"{get_badge_color(e.get('status'))} **{e.get('status','')}**")
                    st.caption(f"Prevista: {formatar_data_br(e.get('dataPrevista'))} | Entrega: {formatar_data_br(e.get('dataEntrega'))}")
                with st.expander("✏️ Editar"):
                    with st.form(f"form_entrega_{e['idSolicitacao']}"):
                        ne_status = st.selectbox("Status", STATUS_OPCOES, index=STATUS_OPCOES.index(e.get("status","Pendente")) if e.get("status") in STATUS_OPCOES else 0)
                        ne_transportadora = st.text_input("Transportadora", value=e.get("transportadora",""))
                        ne_rastreio = st.text_input("Rastreio", value=e.get("rastreio",""))
                        ne_data_envio = st.date_input("Data Envio", value=datetime.strptime(e["dataEnvio"], "%Y-%m-%d").date() if e.get("dataEnvio") else None)
                        ne_data_entrega = st.date_input("Data Entrega", value=datetime.strptime(e["dataEntrega"], "%Y-%m-%d").date() if e.get("dataEntrega") else None)
                        ne_obs = st.text_area("Observações", value=e.get("observacoes",""))
                        if st.form_submit_button("💾 Salvar"):
                            e["status"] = ne_status
                            e["transportadora"] = ne_transportadora
                            e["rastreio"] = ne_rastreio
                            e["dataEnvio"] = str(ne_data_envio) if ne_data_envio else ""
                            e["dataEntrega"] = str(ne_data_entrega) if ne_data_entrega else ""
                            e["observacoes"] = ne_obs
                            # Sincroniza status da solicitação
                            for s in st.session_state["compras_solicitacoes"]:
                                if s["id"] == e["idSolicitacao"]:
                                    s["status"] = ne_status
                                    break
                            st.success("Entrega atualizada!")
                            st.rerun()
    else:
        st.info("Nenhuma entrega encontrada.")


def page_materiais():
    st.markdown("### 📋 Catálogo de Materiais")
    fcliente = st.selectbox("Cliente", ["Todos"] + CLIENTES, key="mat_fcliente")

    dados = []
    for cliente, lista in MATERIAIS_POR_CLIENTE.items():
        if fcliente != "Todos" and cliente != fcliente:
            continue
        for idx, nome in enumerate(lista, 1):
            dados.append({"ID": idx, "Nome": nome, "Cliente": cliente, "Categoria": "Material"})

    if dados:
        df = pd.DataFrame(dados)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Nenhum material encontrado.")


def page_lojas():
    st.markdown("### 🏬 Lojas e Clientes")

    tab1, tab2 = st.tabs(["Lojas", "Clientes"])
    with tab1:
        dados = []
        for cliente, lista in LOJAS_POR_CLIENTE.items():
            for loja in lista:
                dados.append({"Loja": loja, "Cliente": cliente})
        df = pd.DataFrame(dados)
        st.dataframe(df, use_container_width=True, hide_index=True)
    with tab2:
        df = pd.DataFrame([{"Cliente": c, "Qtd Lojas": len(LOJAS_POR_CLIENTE.get(c, [])), "Qtd Materiais": len(MATERIAIS_POR_CLIENTE.get(c, [])), "Qtd EPIs": len(EPIS_POR_CLIENTE.get(c, []))} for c in CLIENTES])
        st.dataframe(df, use_container_width=True, hide_index=True)


def page_relatorios():
    st.markdown("### 📁 Relatórios e Downloads")

    st.markdown("#### 📝 Templates de Documentos")
    st.info("Os documentos são gerados automaticamente ao salvar uma solicitação.")

    # Listar solicitações com anexos
    sols_com_anexo = [s for s in st.session_state["compras_solicitacoes"] if s.get("anexos")]
    if sols_com_anexo:
        for s in sols_com_anexo:
            with st.container(border=True):
                st.write(f"**{s['id']}** - {s.get('loja','')} ({s.get('cliente','')})")
                for a in s["anexos"]:
                    st.download_button(
                        label=f"📥 {a['nome']}",
                        data=a["conteudo"].encode("utf-8"),
                        file_name=a["nome"],
                        mime="application/octet-stream",
                        key=f"rel_anexo_{s['id']}_{a['nome']}"
                    )
    else:
        st.info("Nenhum documento gerado ainda. Crie uma solicitação para gerar documentos.")

    # Download dados completos
    st.markdown("---")
    st.markdown("#### 💾 Exportar Dados")
    if st.session_state["compras_solicitacoes"]:
        sols = st.session_state["compras_solicitacoes"]
        df = pd.DataFrame([
            {
                "ID": s["id"], "Data": s.get("data",""), "Loja": s.get("loja",""),
                "Cliente": s.get("cliente",""), "Tipo": s.get("tipo",""),
                "Solicitante": s.get("solicitante",""), "Status": s.get("status",""),
                "Valor Total": s.get("valorTotal",0)
            }
            for s in sols
        ])
        csv = df.to_csv(index=False, sep=";").encode("utf-8")
        st.download_button("📥 Exportar Solicitações (CSV)", data=csv, file_name="solicitacoes_completo.csv", mime="text/csv")


# ========== MAIN ==========
def render_compras():
    init_session_state()

    st.subheader("🛒 SISTEMA DE COMPRAS E ENTREGAS")
    st.info("Módulo completo de controle de compras, solicitações e entregas.")

    # Menu lateral dentro da aba
    menu = st.sidebar.radio(
        "Navegação Compras",
        ["Dashboard", "Solicitações", "Nova Solicitação", "Controle de Entregas", "Catálogo de Materiais", "Lojas e Clientes", "Relatórios e Downloads"],
        index=["Dashboard", "Solicitações", "Nova Solicitação", "Controle de Entregas", "Catálogo de Materiais", "Lojas e Clientes", "Relatórios e Downloads"].index(st.session_state["compras_page"])
        if st.session_state["compras_page"] in ["Dashboard", "Solicitações", "Nova Solicitação", "Controle de Entregas", "Catálogo de Materiais", "Lojas e Clientes", "Relatórios e Downloads"]
        else 0
    )
    if menu != st.session_state["compras_page"]:
        st.session_state["compras_page"] = menu
        st.session_state["compras_edit_id"] = None
        st.session_state["compras_edit_loaded"] = False
        st.rerun()

    page = st.session_state["compras_page"]
    if page == "Dashboard":
        page_dashboard()
    elif page == "Solicitações":
        page_solicitacoes()
    elif page == "Nova Solicitação":
        page_nova_solicitacao()
    elif page == "Detalhes":
        page_detalhes()
    elif page == "Controle de Entregas":
        page_entregas()
    elif page == "Catálogo de Materiais":
        page_materiais()
    elif page == "Lojas e Clientes":
        page_lojas()
    elif page == "Relatórios e Downloads":
        page_relatorios()
