import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.pyplot as plt
import numpy as np
import io

# Título do app
st.title("📄 Analisador de Conformidades em Documento Word")

st.info(
    "🔒 **Aviso de privacidade**:\n\n"
    "Este aplicativo não armazena permanentemente os arquivos enviados. "
    "Todos os documentos são processados apenas temporariamente durante a sessão atual e são descartados ao final. "
    "Nenhuma informação é salva em banco de dados ou compartilhada com terceiros."
)

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

    doc = Document(uploaded_file)

    # Etapa 1: Contagem e extração de descrições
    count_conforme = 0
    count_nao_conforme = 0
    descricoes_docx = []

    for idx_table, table in enumerate(doc.tables, start=1):
        for row in table.rows:
            for cell in row.cells:
                text = cell.text
                count_conforme += text.count("Conforme")
                count_nao_conforme += text.count("Não conforme")

                if "Descrição" in text:
                    for paragraph in cell.paragraphs:
                        if "Descrição" in paragraph.text:
                            texto_vermelho = ""
                            for run in paragraph.runs:
                                cor = run.font.color
                                if cor and (cor.rgb == RGBColor(255, 0, 0) or cor.rgb == RGBColor(238, 0, 0)):
                                    texto_vermelho += run.text.strip() + " "
                            if texto_vermelho:
                                descricoes_docx.append((texto_vermelho.strip(), idx_table))

    st.write(f"✔️ Total 'Conforme': {count_conforme}")
    st.write(f"❌ Total 'Não Conforme': {count_nao_conforme}")
    st.subheader("📝 Descrições Encontradas")
    st.write([d[0] for d in descricoes_docx])

    # Etapa 2: Gráfico de pizza
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
    labels = ["Conforme", "Não conforme"]
    data = [count_conforme, count_nao_conforme]
    colors = ['#4CAF50', '#F44336']

    def func(pct, allvals):
        absolute = int(np.round(pct / 100. * np.sum(allvals)))
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

    grafico_path = "grafico_pizza.png"
    plt.savefig(grafico_path)
    plt.close()

    st.subheader("📊 Gráfico de Conformidades")
    st.pyplot(fig)

    # Etapa 3: Modificar documento
    doc = Document(uploaded_file)
    doc.add_page_break()

    # Criar nova tabela (3 colunas)
    tabela = doc.add_table(rows=len(descricoes_docx) + 1, cols=3)
    tabela.style = 'Table Grid'
    cabecalhos = ["Descrição", "Figura", "Situação"]

    for col, texto in enumerate(cabecalhos):
        cell = tabela.cell(0, col)
        cell.text = texto
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)
                run.font.bold = True

    for i, (descricao, num_tabela) in enumerate(descricoes_docx, start=1):
        # Descrição
        cell_desc = tabela.cell(i, 0)
        p = cell_desc.paragraphs[0]
        run = p.add_run(descricao)
        run.font.size = Pt(10)

        # Figura (número da tabela)
        cell_fig = tabela.cell(i, 1)
        cell_fig.text = str(num_tabela)
        for paragraph in cell_fig.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

        # Situação (vazio)
        cell_sit = tabela.cell(i, 2)
        cell_sit.text = ""
        for paragraph in cell_sit.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    # Inserir imagem do gráfico
    paragrafo_imagem = doc.add_paragraph()
    run = paragrafo_imagem.add_run()
    run.add_picture(grafico_path, width=Inches(5))
    paragrafo_imagem.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Etapa 4: Salvar novo documento
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Documento finalizado com tabela e gráfico ao final.")
    st.download_button(
        label="📥 Baixar novo Word",
        data=buffer,
        file_name=uploaded_file.name.replace(".docx", " - Análise.docx"),
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
