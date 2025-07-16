import streamlit as st
from docx import Document
from docx.shared import Pt

def gerar_parecer_garantia(dados):
    doc = Document()

    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)

    doc.add_paragraph().add_run("PARECER SIMPLIFICADO").bold = True
    doc.add_paragraph("Para fins de monitoramento de garantias")

    para1 = doc.add_paragraph()
    para1.add_run(
        f"Em atenção à solicitação feita pela {dados['solicitante']} na qualidade de gestora do(s) SDL - FUNDO DE INVESTIMENTO EM DIREITOS CREDITÓRIOS (“Fundo”), apresentamos o presente parecer jurídico simplificado a respeito da capacidade de execução de garantias ligadas a ativos financeiros representativos de dívidas ou obrigações titularizados pelo(s) Fundo(s)."
    )

    doc.add_paragraph().add_run("A) INFORMAÇÕES PRELIMINARES").bold = True
    doc.add_paragraph("1. Para a elaboração deste parecer, foram acessadas as seguintes informações e/ou documentos:")

    table1 = doc.add_table(rows=3, cols=2)
    table1.style = 'Table Grid'
    table1.cell(0, 0).text = "TÍTULO"
    table1.cell(0, 1).text = f"CCB nº {dados['numero_celula']}"
    table1.cell(1, 0).text = "GARANTIA"
    table1.cell(1, 1).text = "Alienação fiduciária de Veículo"
    table1.cell(2, 0).text = "DOCUMENTOS RECEBIDOS"
    table1.cell(2, 1).text = dados['docs']

    doc.add_paragraph().add_run("B) DADOS BÁSICOS DA OPERAÇÃO").bold = True

    table2 = doc.add_table(rows=7, cols=2)
    table2.style = 'Table Grid'
    dados_basicos = [
        ("CESSIONÁRIO", dados["cessionario"]),
        ("CEDENTE", dados["cedente"]),
        ("DEVEDOR(A)", dados["devedor"]),
        ("DATA DE EMISSÃO/REFERÊNCIA", dados["data de emissão"]),
        ("VALOR DA OPERAÇÃO", dados["valor operação"]),
        ("DATA DA PARCELA 1", dados["data primeira parcela"]),
        ("DATA DA PARCELA FINAL", dados["data ultima parcela"]),
    ]

    for i, (campo, valor) in enumerate(dados_basicos):
        table2.cell(i, 0).text = campo
        table2.cell(i, 1).text = valor

    nome_arquivo = f"Parecer_CCB_{dados['numero_celula']}.docx"
    doc.save(nome_arquivo)

    return nome_arquivo


# --- STREAMLIT APP ---
st.title("Gerador de Parecer de Garantia")

with st.form("form_parecer"):
    solicitante = st.text_input("Solicitante", key="solicitante")
    numero_celula = st.text_input("Número da CCB", key="numero_celula")
    docs = st.text_area("Documentos recebidos", key="docs")
    cessionario = st.text_input("Cessionário", key="cessionario")
    cedente = st.text_input("Cedente", key="cedente")
    devedor = st.text_input("Devedor", key="devedor")
    data_emissao = st.date_input("Data de emissão", key="data_emissao").strftime("%d/%m/%Y")
    valor_operacao = st.text_input("Valor da operação", key="valor_operacao")
    data_parcela_1 = st.date_input("Data da primeira parcela", key="data_parcela_1").strftime("%d/%m/%Y")
    data_parcela_final = st.date_input("Data da última parcela", key="data_parcela_final").strftime("%d/%m/%Y")

    gerar = st.form_submit_button("Gerar Parecer", key="botao_submit")
    
    if gerar:
        dados = {
            "solicitante": solicitante,
            "numero_celula": numero_celula,
            "docs": docs,
            "cessionario": cessionario,
            "cedente": cedente,
            "devedor": devedor,
            "data de emissão": data_emissao,
            "valor operação": valor_operacao,
            "data primeira parcela": data_parcela_1,
            "data ultima parcela": data_parcela_final,
        }

        arquivo = gerar_parecer_garantia(dados)
        with open(arquivo, "rb") as file:
            st.success("Parecer gerado com sucesso!")
            st.download_button("Baixar Documento", file, arquivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")


