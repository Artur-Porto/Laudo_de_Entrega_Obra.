import streamlit as st
from docx import Document
from docx.shared import RGBColor, Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import matplotlib.pyplot as plt
import numpy as np
import io

# Função para legenda "Tabela X –"
def add_caption_field_before(table, idx, cor_hex="FFFFFF"):
    p = OxmlElement('w:p')

    # Estilo de parágrafo "Caption"
    pPr = OxmlElement('w:pPr')
    pStyle = OxmlElement('w:pStyle')
    pStyle.set(qn('w:val'), 'Caption')
    pPr.append(pStyle)
    p.append(pPr)

    # Campo SEQ begin
    r1 = OxmlElement('w:r')
    fld1 = OxmlElement('w:fldChar')
    fld1.set(qn('w:fldCharType'), 'begin')
    r1.append(fld1)
    p.append(r1)

    # Campo SEQ instrução
    r2 = OxmlElement('w:r')
    instr = OxmlElement('w:instrText')
    instr.set(qn('xml:space'), 'preserve')
    instr.text = 'SEQ Table \\* ARABIC'
    r2.append(instr)
    p.append(r2)

    # Campo SEQ separate
    r3 = OxmlElement('w:r')
    fld2 = OxmlElement('w:fldChar')
    fld2.set(qn('w:fldCharType'), 'separate')
    r3.append(fld2)
    p.append(r3)

    # Texto "Tabela X – " com cor personalizada
    r4 = OxmlElement('w:r')

    rPr = OxmlElement('w:rPr')
    color = OxmlElement('w:color')
    color.set(qn('w:val'), cor_hex)  # Cor personalizada
    rPr.append(color)
    r4.append(rPr)

    t = OxmlElement('w:t')
    t.text = f'Figura {idx} – '
    r4.append(t)
    p.append(r4)

    # Campo SEQ end
    r5 = OxmlElement('w:r')
    fld3 = OxmlElement('w:fldChar')
    fld3.set(qn('w:fldCharType'), 'end')
    r5.append(fld3)
    p.append(r5)

    # Inserir antes da tabela
    tbl = table._element
    tbl.addprevious(p)

# Função para criar campo de índice ("Lista de Tabelas")
def add_field_code(paragraph, field_code):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = field_code
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

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

# Upload
uploaded_file = st.file_uploader("Envie o arquivo Word (.docx)", type="docx")

if uploaded_file:
    st.success("Arquivo carregado com sucesso!")
    file_name = uploaded_file.name.replace(".docx", " - Análise.docx")

    doc = Document(uploaded_file)

    # Adicionar legendas antes de cada tabela
    for i, tbl in enumerate(doc.tables, start=1):
        add_caption_field_before(tbl, i)

    # Análise de conformidade e extração
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
                                if cor and cor.rgb == RGBColor(255, 0, 0):
                                    texto_vermelho += run.text.strip() + " "
                            if texto_vermelho:
                                descricoes_docx.append((texto_vermelho.strip(), idx_table))

    st.write(f"✔️ Total 'Conforme': {count_conforme}")
    st.write(f"❌ Total 'Não Conforme': {count_nao_conforme}")

    st.subheader("📝 Descrições Encontradas")
    if descricoes_docx:
        for i, (desc, fig) in enumerate(descricoes_docx, 1):
            st.markdown(f"**{i}.** {desc} *(Figura {fig})*")
    else:
        st.info("Nenhuma descrição encontrada.")

    # Gráfico
    fig, ax = plt.subplots(figsize=(6, 3), subplot_kw=dict(aspect="equal"))
    labels = ["Conforme", "Não conforme"]
    data = [count_conforme, count_nao_conforme]
    colors = ['#4CAF50', '#F44336']

    def func(pct, allvals):
        absolute = int(np.round(pct/100.*np.sum(allvals)))
        return f"{pct:.1f}%\n({absolute:d})"

    wedges, texts, autotexts = ax.pie(
        data, autopct=lambda pct: func(pct, data),
        textprops=dict(color="w"), colors=colors
    )
    ax.legend(wedges, labels, title="Situação", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
    plt.setp(autotexts, size=8, weight="bold")
    ax.set_title("Análise de Conformidades")
    st.subheader("📊 Gráfico de Conformidades")
    st.pyplot(fig)

    # Salvar gráfico
    plt.savefig("grafico_pizza.png")
    plt.close()

    # Inserção de nova página
    doc.add_page_break()

    # Tabela de resultados
    tabela = doc.add_table(rows=len(descricoes_docx) + 1, cols=6)
    tabela.style = 'Table Grid'
    cabecalhos = ["Descrição", "Normativo", "Projeto", "Boas práticas", "Situação", "Figura"]

    for col, texto in enumerate(cabecalhos):
        cell = tabela.cell(0, col)
        cell.text = texto
        for run in cell.paragraphs[0].runs:
            run.font.size = Pt(10)
            run.font.bold = True

    for i, (descricao, num_tabela) in enumerate(descricoes_docx, start=1):
        cell_desc = tabela.cell(i, 0)
        run = cell_desc.paragraphs[0].add_run(descricao)
        run.font.size = Pt(10)
        for col in range(1, 5):
            cell = tabela.cell(i, col)
            for run in cell.paragraphs[0].runs:
                run.font.size = Pt(10)
        tabela.cell(i, 5).text = str(num_tabela)

    # Gráfico no final
    paragrafo_imagem = doc.add_paragraph()
    run = paragrafo_imagem.add_run()
    run.add_picture('grafico_pizza.png', width=Inches(5))
    paragrafo_imagem.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Lista de Tabelas
    #doc.add_page_break()
    doc.add_paragraph("Lista de Figuras").style = 'Heading 1'
    p_lista = doc.add_paragraph()
    add_field_code(p_lista, 'TOC \\h \\z \\c "Table"')

    # Gerar download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Documento finalizado com gráfico, tabela e lista de tabelas.")
    st.download_button(
        label="📥 Baixar novo Word",
        data=buffer,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
