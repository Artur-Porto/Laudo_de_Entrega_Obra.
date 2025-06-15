
import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import fitz  # PyMuPDF
import tempfile
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📄 Análise de Conformidades - Documento Word")

# --- Proteção com senha usando secrets ---
st.sidebar.title("🔒 Acesso Restrito")
senha_correta = st.secrets["senha"]
senha_digitada = st.sidebar.text_input("Digite a senha:", type="password")

if senha_digitada != senha_correta:
    st.warning("Acesso negado. Insira a senha correta.")
    st.stop()

# Upload do arquivo .docx
uploaded_file = st.file_uploader("Faça upload do arquivo Word (.docx)", type="docx")

if uploaded_file:
    # Lê o documento
    doc = Document(uploaded_file)

    # Contagem de palavras-chave e extração de descrições em vermelho
    count_conforme = 0
    count_nao_conforme = 0
    descricoes_docx = []

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                text = cell.text
                if "Não conforme" in text:
                    count_nao_conforme += 1
                if "Conforme" in text:
                    count_conforme += 1

                if "Descrição" in text:
                    for paragraph in cell.paragraphs:
                        if "Descrição" in paragraph.text:
                            texto_vermelho = ""
                            for run in paragraph.runs:
                                cor = run.font.color
                                if cor and cor.rgb == RGBColor(255, 0, 0):
                                    texto_vermelho += run.text.strip() + " "
                            if texto_vermelho:
                                descricoes_docx.append(texto_vermelho.strip())

    # Mostrar resultados
    st.subheader("🔢 Contagem")
    st.write(f"✅ Conforme: **{count_conforme}**")
    st.write(f"❌ Não conforme: **{count_nao_conforme}**")

    # Mostrar descrições encontradas
    st.subheader("📌 Descrições encontradas")
    for d in descricoes_docx:
        st.markdown(f"- {d}")

    # Criar gráfico
    fig, ax = plt.subplots(figsize=(5, 3), subplot_kw=dict(aspect="equal"))
    labels = ["Conforme", "Não conforme"]
    data = [count_conforme, count_nao_conforme]
    colors = ['#4CAF50', '#F44336']

    def func(pct, allvals):
        absolute = int(np.round(pct/100.*np.sum(allvals)))
        return f"{pct:.1f}%\n({absolute:d})"

    wedges, texts, autotexts = ax.pie(
        data,
        autopct=lambda pct: func(pct, data),
        textprops=dict(color="w"),
        colors=colors
    )

    ax.legend(wedges, labels,
              title="Situação",
              loc="center left",
              bbox_to_anchor=(1, 0, 0.5, 1))
    plt.setp(autotexts, size=8, weight="bold")
    ax.set_title("Análise de Conformidades")

    # Salvar em buffer
    img_buffer = BytesIO()
    plt.savefig(img_buffer, format='png')
    plt.close()
    img_buffer.seek(0)

    st.subheader("📊 Gráfico de Pizza")
    st.image(img_buffer)

    # Criar novo documento Word
    st.subheader("📄 Gerar novo documento Word")

    if st.button("Gerar e baixar novo documento"):
        novo_doc = Document()
        novo_doc.add_heading("Resumo Não Conformidades", level=1)

        # Criar tabela
        tabela = novo_doc.add_table(rows=len(descricoes_docx)+1, cols=5)
        tabela.style = 'Table Grid'

        cabecalhos = ["Descrição", "Normativo", "Projeto", "Boas práticas", "Situação"]

        for col, texto in enumerate(cabecalhos):
            cell = tabela.cell(0, col)
            cell.text = texto
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.bold = True

        for i, descricao in enumerate(descricoes_docx, start=1):
            cell_desc = tabela.cell(i, 0)
            cell_desc.text = ""
            p = cell_desc.paragraphs[0]
            run = p.add_run(descricao)
            run.font.size = Pt(10)

            for col in range(1, 5):
                cell = tabela.cell(i, col)
                cell.text = ""
                for paragraph in cell.paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)

        # Adicionar gráfico
        novo_doc.add_page_break()
        novo_doc.add_heading("Gráfico de Análise Conformidades", level=1)
        paragrafo = novo_doc.add_paragraph()
        run = paragrafo.add_run()
        run.add_picture(img_buffer, width=Inches(5))
        paragrafo.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Salvar em buffer e permitir download
        output_buffer = BytesIO()
        novo_doc.save(output_buffer)
        output_buffer.seek(0)

        st.success("Documento gerado com sucesso!")
        st.download_button("📥 Baixar Documento Word", data=output_buffer, file_name="analise_conformidades.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
