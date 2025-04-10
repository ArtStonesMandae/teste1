# app_mandae.py original aqui (exemplo simplificado para reconstrução)



import streamlit as st
import pandas as pd
import zipfile
import xml.etree.ElementTree as ET
import openpyxl
from io import BytesIO
from datetime import datetime, timedelta

# 🔁 NOVA SEÇÃO
st.title("🔄 Atualizar CHAVE NF com XMLs de Nota Fiscal")
st.write("Atualize a coluna 'CHAVE NF' da planilha Mandae automaticamente com base nos arquivos XML de NF-e.")

planilha_file = st.file_uploader("1. Selecione a planilha Mandae (.xlsx):", type=["xlsx"], key="xlsx_upload")
zip_file = st.file_uploader("2. Selecione o arquivo .ZIP com os XMLs de NF-e:", type=["zip"], key="zip_upload")

if planilha_file and zip_file:
    try:
        # Abrir planilha
        wb = openpyxl.load_workbook(planilha_file)
        ws = wb.active

        # Obter índice da coluna do CPF e da CHAVE NF
        header = [cell.value for cell in ws[2]]
        idx_cpf = header.index("CPF / CNPJ CLIENTE*") + 1
        idx_chave = header.index("CHAVE NF") + 1

        # Criar dicionário de CPF -> linha
        planilha_cpfs = {}
        for row in ws.iter_rows(min_row=3, min_col=idx_cpf, max_col=idx_cpf):
            cpf_val = str(row[0].value).zfill(11)
            planilha_cpfs[cpf_val] = row[0].row

        # Abrir ZIP de XMLs
        cpf_para_chave = {}
        with zipfile.ZipFile(zip_file) as z:
            for name in z.namelist():
                if name.endswith(".xml"):
                    with z.open(name) as f:
                        try:
                            tree = ET.parse(f)
                            root = tree.getroot()
                            ns = { 'ns': root.tag.split('}')[0].strip('{') }

                            cpf = root.findtext('.//ns:CPF', namespaces=ns)
                            chave = root.findtext('.//ns:chNFe', namespaces=ns)

                            if cpf and chave:
                                cpf = cpf.zfill(11)
                                cpf_para_chave[cpf] = chave
                        except:
                            continue

        # Atualizar planilha
        atualizados = 0
        for cpf, row_idx in planilha_cpfs.items():
            if cpf in cpf_para_chave:
                ws.cell(row=row_idx, column=idx_chave, value=cpf_para_chave[cpf])
                atualizados += 1

        # Salvar arquivo de saída
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        hoje = datetime.today()
        dia_util = hoje + timedelta(days=1)
        if hoje.weekday() == 4:
            dia_util += timedelta(days=2)
        nome_final = f"{len(planilha_cpfs)}Pedidos - {dia_util.strftime('%d.%m')} - L2.xlsx"

        st.success(f"Chaves atualizadas com sucesso: {atualizados} de {len(planilha_cpfs)} pedidos.")
        st.download_button("📅 Baixar Planilha Atualizada", data=output, file_name=nome_final, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"Erro no processamento: {e}")
