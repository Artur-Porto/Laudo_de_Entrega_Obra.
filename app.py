import streamlit as st
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import numpy as np
import io
import pandas as pd 
import re

# Função robusta para verificar se uma descrição está "vazia de verdade"
def is_vazio(texto):
    texto_limpo = re.sub(r'[\s\u200b\u200c\u200d\uFEFF]+', '', texto)
    return texto_limpo == ''

st.title("📄 Analisador de Conformidades em Documento Word")

st.info(
    "🔒 **Aviso de privacidade**:\n\n"
    "Este aplicativo não armazena permanentemente os arquivos enviados. "
    "Todos os documentos são processados apenas durante a sessão atual e são descartados ao final."
)

st.subheader("🔒 Acesso Restrito")
senha_correta = st.secrets["senha"]
senha_digitada = st.text_input("Digite a senha para continuar:", type="password")
if senha_digitada != senha_correta:
    st.warning("Acesso negado. Insira a senha correta.")
    st.stop()

uploaded_file = st.file_uploader("📤 Envie o arquivo Word (.docx)", type="docx")

if uploaded_file:
    st.success("📁 Arquivo carregado com sucesso!")
    doc = Document(uploaded_file)

    count_conforme = 0
    count_nao_conforme = 0
    descricoes_docx = []

    for idx_table, table in enumerate(doc.tables, start=1):
        for row in table.rows:
            for cell in row.cells:
                text = cell.text
                count_nao_conforme += text.count("Não conforme")
                count_conforme += text.count("Conforme")

                if "Descrição" in text:
                    for paragraph in cell.paragraphs:
                        if "Descrição" in paragraph.text:
                            texto_vermelho = ""
                            for run in paragraph.runs:
                                cor = run.font.color
                                if cor and cor.rgb in [RGBColor(255, 0, 0), RGBColor(238, 0, 0)]:
                                    texto_vermelho += run.text.strip() + " "
                            if texto_vermelho:
                                descricoes_docx.append((texto_vermelho.strip(), idx_table))
     # Verificação final de descrições vazias
    descricoes_vazias = [d for d in descricoes_docx if is_vazio(d[0])]
    num_descricoes_vazias = len(descricoes_vazias)
    num_descricoes_validas = len(descricoes_docx) - num_descricoes_vazias

# Se o número de descrições vazias bater com o excesso na contagem, remove
    if (count_nao_conforme - num_descricoes_validas) == num_descricoes_vazias:
        descricoes_docx = [d for d in descricoes_docx if not is_vazio(d[0])]
        st.info(f"⚠️ {num_descricoes_vazias} descrições vazias removidas automaticamente.")


    st.write(f"✔️ Total 'Conforme': {count_conforme}")
    st.write(f"❌ Total 'Não Conforme': {count_nao_conforme}")

    # Mostrar descrições em formato de tabela
    st.subheader("📝 Descrições Encontradas")

    if descricoes_docx:
        df_descricoes = pd.DataFrame(descricoes_docx, columns=["Descrição", "Figura"])
        st.table(df_descricoes)
    else:
        st.info("Nenhuma descrição em vermelho foi encontrada.")

    # Gráfico
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
    labels = ["Conforme", "Não conforme"]
    data = [count_conforme, count_nao_conforme]
    colors = ['#4CAF50', '#F44336']

    def func(pct, allvals):
        absolute = int(np.round(pct / 100. * np.sum(allvals)))
        return f"{pct:.1f}%\n({absolute:d})"

    wedges, _, autotexts = ax.pie(
        data, autopct=lambda pct: func(pct, data),
        textprops=dict(color="w"),
        colors=colors
    )
    ax.legend(wedges, labels, title="Situação", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.setp(autotexts, size=8, weight="bold")
    ax.set_title("Análise de Conformidades")
    grafico_path = "grafico_pizza.png"
    plt.savefig(grafico_path)
    st.subheader("📊 Gráfico de Conformidades")
    st.pyplot(fig)
    plt.close()

    # Inserir tabela e gráfico no documento original
    doc.add_page_break()

    tabela = doc.add_table(rows=len(descricoes_docx) + 1, cols=3)
    tabela.style = 'Table Grid'
    cabecalhos = ["Descrição", "Figura", "Situação"]

    # Cabeçalhos
    for col, texto in enumerate(cabecalhos):
        cell = tabela.cell(0, col)
        cell.text = texto
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(10)
            run.font.bold = True

    # Linhas da tabela
    for i, (descricao, num_tabela) in enumerate(descricoes_docx, start=1):
        # Descrição
        cell_desc = tabela.cell(i, 0)
        run = cell_desc.paragraphs[0].add_run(descricao)
        run.font.size = Pt(10)
        # Figura
        cell_fig = tabela.cell(i, 1)
        cell_fig.text = str(num_tabela)
        for run in cell_fig.paragraphs[0].runs:
            run.font.size = Pt(10)
        # Situação (vazio)
        cell_sit = tabela.cell(i, 2)
        cell_sit.text = ""
        for run in cell_sit.paragraphs[0].runs:
            run.font.size = Pt(10)

    # Inserir gráfico
    paragrafo_img = doc.add_paragraph()
    run = paragrafo_img.add_run()
    run.add_picture(grafico_path, width=Inches(5))
    paragrafo_img.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Salvar
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Documento atualizado com gráfico e tabela ao final.")
    st.download_button(
        label="📥 Baixar novo Word",
        data=buffer,
        file_name=uploaded_file.name.replace(".docx", " - Análise.docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
