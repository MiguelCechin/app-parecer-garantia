import streamlit as st
from docx import Document
from docx.shared import Pt

def avaliar_titulo(respostas):
    r11 = respostas.get("1.1", "").upper()

    if r11 == "NÃO":
        return "  2.  Não foi apresentado, até o momento, instrumento formal de constituição da dívida, impossibilitando a verificação de sua regularidade e eventual cessibilidade."
    elif r11 == "NÃO SE APLICA":
        return "  2.  A análise do título/instrumento de formalização do crédito não se aplica à presente garantia."
    elif r11 == "SIM":
        r12 = respostas.get("1.2", "").upper()
        if r12 == "NÃO":
            return "  2.  Inicialmente, verificou-se a existência de documento representativo de dívida que carece de assinatura e não preenche os requisitos formais mínimos aplicáveis."
        elif r12 == "SIM":
            r13 = respostas.get("1.3", "").upper()
            if r13 == "SIM":
                r14 = respostas.get("1.4", "").upper()
                if r14 == "NÃO":
                    return "  2.  Inicialmente, verificou-se a existência de documento representativo de dívida que carece da assinatura de testemunhas, e portanto não preenche os requisitos formais mínimos aplicáveis."
                elif r14 == "SIM":
                    return "  2.  Inicialmente, verificou-se a existência de documento representativo de dívida com o preenchimento de requisitos formais mínimos aplicáveis."
            elif r13 == "NÃO":
                return "  2.  Inicialmente, verificou-se a existência de documento representativo de dívida que carece da assinatura de testemunhas, e portanto não preenche os requisitos formais mínimos aplicáveis."
    return "Resposta inválida ou incompleta."


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
        f"Em atenção à solicitação feita pela {dados['solicitante']} na qualidade de gestora do(s) {dados['gestora']}, apresentamos o presente parecer jurídico simplificado a respeito da capacidade de execução de garantias ligadas a ativos financeiros representativos de dívidas ou obrigações titularizados pelo(s) Fundo(s)."
    )

    doc.add_paragraph().add_run("A) INFORMAÇÕES PRELIMINARES").bold = True
    doc.add_paragraph("1. Para a elaboração deste parecer, foram acessadas as seguintes informações e/ou documentos:")

    table1 = doc.add_table(rows=3, cols=2)
    table1.style = 'Table Grid'
    table1.cell(0,0).text = "TÍTULO"
    table1.cell(0,1).text = f"CCB nº {dados['numero_celula']}"
    table1.cell(1,0).text = "GARANTIA"
    table1.cell(1,1).text = "Alienação fiduciária de Veículo"
    table1.cell(2,0).text = "DOCUMENTOS RECEBIDOS"
    table1.cell(2,1).text = dados["docs"]

    doc.add_paragraph().add_run("B) DADOS BÁSICOS DA OPERAÇÃO").bold = True

    table2 = doc.add_table(7,2)
    table2.style = 'Table Grid'
    table2.cell(0,0).text = "CESSIONÁRIO"
    table2.cell(0,1).text = dados["cessionario"]
    table2.cell(1,0).text = "CEDENTE"
    table2.cell(1,1).text = dados["cedente"]
    table2.cell(2,0).text = "DEVEDOR(A)"
    table2.cell(2,1).text = dados["devedor"]
    table2.cell(3,0).text = "DATA DE EMISSÃO/REFERÊNCIA"
    table2.cell(3,1).text = dados["data de emissão"]
    table2.cell(4,0).text = "VALOR DA OPERAÇÃO"
    table2.cell(4,1).text = f"R$ {dados['valor operação']}"
    table2.cell(5,0).text = "DATA DA PARCELA 1"
    table2.cell(5,1).text = dados["data primeira parcela"]
    table2.cell(6,0).text = "DATA DA PARCELA FINAL"
    table2.cell(6,1).text = dados["data ultima parcela"]

    doc.add_paragraph().add_run("C) GARANTIA").bold = True
    table3 = doc.add_table(5,2)
    table3.style = 'Table Grid'
    table3.cell(0,0).text = "FIDUCIANTE"
    table3.cell(0,1).text = dados["fiduciante"]
    table3.cell(1,0).text = "FIDUCIÁRIO(A) ORIGINAL"
    table3.cell(1,1).text = dados["fiduciário"]
    table3.cell(2,0).text = "BEM OBJETO DE GARANTIA"
    table3.cell(2,1).text = dados["informações do objeto de garantia"]
    table3.cell(3,0).text = "VALOR DE AVALIAÇÃO HISTÓRICO"
    table3.cell(3,1).text = f"R$ {dados['valor do objeto inicial']}"
    table3.cell(4,0).text = "VALOR DE AVALIAÇÃO ATUAL"
    table3.cell(4,1).text = f"R$ {dados['valor do objeto atual']}"

    doc.add_paragraph().add_run("D) PRINCIPAIS CONSTATAÇÕES E APONTAMENTOS").bold = True
    doc.add_paragraph(avaliar_titulo(dados))

    nome_arquivo = f"Parecer_CCB_{dados['numero_celula']}.docx"
    doc.save(nome_arquivo)
    return nome_arquivo


# --- APP STREAMLIT ---
st.title("Gerador de Parecer de Garantia")

with st.form("formulario"):
    st.subheader("Informações Gerais")
    solicitante = st.text_input("Solicitante")
    gestora = st.text_input("Gestora")
    numero_celula = st.text_input("Número da CCB")
    docs = st.text_area("Documentos Recebidos")

    st.subheader("Dados da Operação")
    cessionario = st.text_input("Cessionário")
    cedente = st.text_input("Cedente")
    devedor = st.text_input("Devedor")
    data_emissao = st.text_input("Data de emissão")
    valor_operacao = st.text_input("Valor da operação")
    data_parcela_1 = st.text_input("Data da parcela 1")
    data_parcela_final = st.text_input("Data da parcela final")

    st.subheader("Dados da Garantia")
    fiduciante = st.text_input("Fiduciante")
    fiduciario = st.text_input("Fiduciário(a)")
    info_objeto = st.text_input("Informações do objeto de garantia")
    valor_objeto_inicial = st.text_input("Valor do objeto inicial")
    valor_objeto_atual = st.text_input("Valor do objeto atual")

    st.subheader("Decisão sobre o TÍTULO")
    r11 = st.selectbox("1.1", ["SIM", "NÃO", "NÃO SE APLICA"])
    r12 = st.selectbox("1.2", ["SIM", "NÃO"])
    r13 = st.selectbox("1.3", ["SIM", "NÃO"])
    r14 = st.selectbox("1.4", ["SIM", "NÃO"])

    gerar = st.form_submit_button("Gerar Parecer")

    if gerar:
        dados = {
            "solicitante": solicitante,
            "gestora": gestora,
            "numero_celula": numero_celula,
            "docs": docs,
            "cessionario": cessionario,
            "cedente": cedente,
            "devedor": devedor,
            "data de emissão": data_emissao,
            "valor operação": valor_operacao,
            "data primeira parcela": data_parcela_1,
            "data ultima parcela": data_parcela_final,
            "fiduciante": fiduciante,
            "fiduciário": fiduciario,
            "informações do objeto de garantia": info_objeto,
            "valor do objeto inicial": valor_objeto_inicial,
            "valor do objeto atual": valor_objeto_atual,
            "1.1": r11,
            "1.2": r12,
            "1.3": r13,
            "1.4": r14,
        }

        arquivo = gerar_parecer_garantia(dados)
        with open(arquivo, "rb") as file:
            st.success("Parecer gerado com sucesso!")
            st.download_button("Baixar Documento", file, arquivo, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

