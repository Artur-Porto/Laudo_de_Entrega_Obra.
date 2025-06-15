import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import numpy as np
import io

# Título do app
st.title("📄 Analisador de Conformidades em Documento Word")

# Campo de senha
st.subheader("🔒 Acesso Restrito")
senha_correta = st.secrets["senha"]
senha_digitada = st.text_input("Digite a senha para continuar:", type="password")

if senha_digitada != senha_correta:
    st.warning("Acesso negado. Insira a senha correta.")
    st.stop()

# Upload do arquivo
uploaded_file = st.file_uploader("Envie o arquivo Word (.docx)", type="docx")

if uploaded_file:
    st.success("Arquivo carregado com sucesso!")
    
    # Carregar documento original
    doc = Document(uploaded_file)

    # Contagem e extração de descrições
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

    st.write(f"✔️ Total 'Conforme': {count_conforme}")
    st.write(f"❌ Total 'Não Conforme': {count_nao_conforme}")

    # Gerar gráfico de pizza
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
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
    ax.legend(wedges, labels, title="Situação", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.setp(autotexts, size=8, weight="bold")
    ax.set_title("Análise de Conformidades")

    # Salvar gráfico
    grafico_path = "grafico_pizza.png"
    plt.savefig(grafico_path)
    plt.close()

    st.subheader("📊 Gráfico de Conformidades")
    st.pyplot(fig)

    st.subheader("📝 Descrições Encontradas")
    st.write(descricoes_docx)

    # Inserir no documento
    doc.add_page_break()

    # Adicionar tabela
    cabecalhos = ["Descrição", "Normativo", "Projeto", "Boas práticas", "Situação"]
    tabela = doc.add_table(rows=len(descricoes_docx) + 1, cols=5)
    tabela.style = 'Table Grid'

    for col, texto in enumerate(cabecalhos):
        cell = tabela.cell(0, col)
        cell.text = texto
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(10)
            run.font.bold = True

    for i, descricao in enumerate(descricoes_docx, start=1):
        cell_desc = tabela.cell(i, 0)
        run = cell_desc.paragraphs[0].add_run(descricao)
        run.font.size = Pt(10)
        for col in range(1, 5):
            cell = tabela.cell(i, col)
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(10)

    # Adicionar gráfico ao final
    paragrafo_imagem = doc.add_paragraph()
    run = paragrafo_imagem.add_run()
    run.add_picture(grafico_path, width=Inches(5))
    paragrafo_imagem.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Salvar novo arquivo para download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Documento finalizado com tabela e gráfico ao final.")
    st.download_button(
        label="📥 Baixar novo Word",
        data=buffer,
        file_name = uploaded_file.name.replace(".docx", " - Análise.docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
