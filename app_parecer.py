import streamlit as st
import docx
from docx import Document
from docx.shared import Pt
from io import BytesIO

def avaliar_titulo(respostas):
    r11 = respostas.get("1.1", "").upper()
    if r11 == "NÃO":
        return "  2.    Não foi apresentado, até o momento, instrumento formal de constituição da dívida, impossibilitando a verificação de sua regularidade e eventual cessibilidade."
    elif r11 == "NÃO SE APLICA":
        return "  2.    A análise do título/instrumento de formalização do crédito não se aplica à presente garantia."
    elif r11 == "SIM":
        r12 = respostas.get("1.2", "").upper()
        if r12 == "NÃO":
            return "  2.    Inicialmente, verificou-se a existência de documento representativo de dívida que carece de assinatura, e portanto não preenche os requisitos formais mínimos aplicáveis."
        elif r12 == "SIM":
            r13 = respostas.get("1.3", "").upper()
            if r13 == "SIM":
                return " 2.    Inicialmente, verificou-se a existência de documento representativo de dívida com o preenchimento do requisitos formais mínimos aplicáveis."
            elif r13 == "NÃO":
                r14 = respostas.get("1.4", "").upper()
                if r14 == "NÃO":
                    return "  2.    Inicialmente, verificou-se a existência de documento representativo de dívida que carece da assinatura de testemunhas, e portanto não preenche os requisitos formais mínimos aplicáveis."
                elif r14 == "SIM":
                    return "  2.    Inicialmente, verificou-se a existência de documento representativo de dívida com o preenchimento de requisitos formais mínimos aplicáveis."
    return "Resposta inválida ou incompleta."
def avaliar_vedacao(respostas):
    r21 = respostas.get("2.1", "").upper()
    if r21 == "NÃO":
        return " Sem cláusula impeditiva da cessão/endosso."
    elif r21 == "NÃO SE APLICA":
        return ""
    elif r21 == "SIM":
        r22 = respostas.get("2.2", "").upper()
        if r22 == "SIM":
            return " Verificou-se que o devedor expressamente autorizou a cessão do crédito, não havendo óbice à sua realização, ainda que presente cláusula contratual originalmente impeditiva."
        elif r22 == "NÃO":
            return " Consta cláusula expressa impeditiva da cessão/endosso do crédito, e não se verificou a apresentação de autorização por parte deste."
        elif r22 == "NÃO SE APLICA":
            return ""
    return "Resposta inválida ou incompleta."
def avaliar_cessao(respostas):
    r31 = respostas.get("3.1", "").upper()
    if r31 == "NÃO":
        return avaliar_comunicacao_cessao(respostas)
    elif r31 == "NÃO SE APLICA":
        return "  3.    A análise do título/instrumento da cessão do crédito não se aplica ao presente caso, conforme as especificidades da operação."
    elif r31 == "SIM":
        r32 = respostas.get("3.2", "").upper()
        if r32 == "NÃO":
            return "  3.    Verificou-se a existência de documento representativo de cessão de crédito que carece de assinatura, e portanto não preenche os requisitos formais mínimos aplicáveis."
        elif r32 == "SIM":
            r33 = respostas.get("3.3", "").upper()
            if r33 == "SIM":
                return " 3.    Verificou-se a existência de documento representativo de cessão de crédito com o preenchimento do requisitos formais mínimos aplicáveis."
            elif r33 == "NÃO":
                r34 = respostas.get("3.4", "").upper()
                if r34 == "NÃO":
                    return "  3.    Verificou-se a existência de documento representativo de cessão de crédito que carece da assinatura de testemunhas, e portanto não preenche os requisitos formais mínimos aplicáveis."
                elif r34 == "SIM":
                    return "  3.    Verificou-se a existência de documento representativo de cessão de crédito com o preenchimento do requisitos formais mínimos aplicáveis."
    return "Resposta inválida ou incompleta." 

def avaliar_comunicacao_cessao(respostas):
    r41 = respostas.get("4.1", "").upper()
    if r41 == "NÃO":
        return "  3.    Não foi submetido à analise o instrumento que resultou na cessão do crédito ou qualquer tipo de notificação para fins de comunicação da cessão ao Fundo de Investimento. "
    elif r41 == "SIM":
        r42 = respostas.get("4.2", "").upper()
        if r42 == "NÃO":
            return "  3.    Não foi submetido à analise o intrumento que resultou na cessão de crédito ao Fundo de Investimento. Entretanto, a existência da cessão pode ser aferida por meio da notificação para fins de comunicação da cessão, ressalvando-se que a referida comunicação não permite pleno entendimento dos termos da cessão e carece de assinatura, não preenchendo os requisitos minimos formais aplicáveis."
        elif r42 == "SIM":
            r43 = respostas.get("4.3", "").upper()
            if r43 == "SIM":
                return " 3.    Não foi submetido à analise o intrumento que resultou na cessão de crédito ao Fundo de Investimento. Entretanto, a existência da cessão pode ser aferida por meio da notificação para fins de comunicação da cessão, ressalvando-se que a referida comunicação não permite pleno entendimento dos termos da cessão."
            elif r43 == "NÃO":
                r44 = respostas.get("4.4", "").upper()
                if r44 == "NÃO":
                    return "  3.    Não foi submetido à analise o intrumento que resultou na cessão de crédito ao Fundo de Investimento. Entretanto, a existência da cessão pode ser aferida por meio da notificação para fins de comunicação da cessão, ressalvando-se que a referida comunicação não permite pleno entendimento dos termos da cessão e carece de assinatura das testemunhas, não preenchendo os requisitos minimos formais aplicáveis."
                elif r44 == "SIM":
                    return "  3.    Não foi submetido à analise o intrumento que resultou na cessão de crédito ao Fundo de Investimento. Entretanto, a existência da cessão pode ser aferida por meio da notificação para fins de comunicação da cessão, ressalvando-se que a referida comunicação não permite pleno entendimento dos termos da cessão."
    return "Resposta inválida ou incompleta."
    
def gerar_parecer_garantia(dados):
    doc = Document()
    # Configurar fonte padrão
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(9)

    # Título
    titulo = doc.add_paragraph()
    titulo.alignment = 1
    titulo.add_run("PARECER SIMPLIFICADO").bold = True

    # Subtítulo
    subtitulo = doc.add_paragraph()
    subtitulo.alignment = 1
    subtitulo.add_run("Para fins de monitoramento de garantias")

    # Parágrafo inicial
    para1 = doc.add_paragraph()
    para1.alignment = 3
    para1.add_run(
        f"Em atenção à solicitação feita pela {dados['solicitante']} na qualidade de gestora do(s) {dados['gestora']}, apresentamos o presente parecer jurídico simplificado a respeito da capacidade de execução de garantias ligadas a ativos financeiros representativos de dívidas ou obrigações titularizados pelo(s) Fundo(s)."
    )
    # Seção A
    doc.add_paragraph().add_run("A) INFORMAÇÕES PRELIMINARES").bold = True
    doc.add_paragraph().add_run("1. Para a elaboração deste parecer, foram acessadas as seguintes informações e/ou documentos:")

    table1 = doc.add_table(rows=3, cols=2)
    table1.style = 'Table Grid'
    table1.cell(0,0).text = "TÍTULO"
    table1.cell(0,1).text = f"CCB nº {dados['numero_celula']}"
    table1.cell(1,0).text = "GARANTIA"
    table1.cell(1,1).text = "Alienação fiduciária de veículo"
    table1.cell(2,0).text = "DOCUMENTOS RECEBIDOS"
    table1.cell(2,1).text = "\n".join(dados['docs'])
    doc.add_paragraph()

    # Seção B
    doc.add_paragraph().add_run("B) DADOS BÁSICOS DA OPERAÇÃO").bold = True
    table2 = doc.add_table(rows=7, cols=2)
    table2.style = 'Table Grid'
    table2.cell(0,0).text = "CESSIONÁRIO"
    table2.cell(0,1).text = dados['cessionario']
    table2.cell(1,0).text = "CEDENTE"
    table2.cell(1,1).text = dados['cedente']
    table2.cell(2,0).text = "DEVEDOR(A)"
    table2.cell(2,1).text = dados['devedor']
    table2.cell(3,0).text = "DATA DE EMISSÃO/REFERÊNCIA"
    table2.cell(3,1).text = dados['data_emissao']
    table2.cell(4,0).text = "VALOR DA OPERAÇÃO"
    table2.cell(4,1).text = f"R$ {dados['valor_operacao']}"
    table2.cell(5,0).text = "DATA DA PARCELA 1"
    table2.cell(5,1).text = dados['data_primeira_parcela']
    table2.cell(6,0).text = "DATA DA PARCELA FINAL"
    table2.cell(6,1).text = dados['data_ultima_parcela']
    doc.add_paragraph()

    # Seção C
    doc.add_paragraph().add_run("C) GARANTIA").bold = True
    table3 = doc.add_table(rows=5, cols=2)
    table3.style = 'Table Grid'
    table3.cell(0,0).text = "FIDUCIANTE"
    table3.cell(0,1).text = dados['fiduciante']
    table3.cell(1,0).text = "FIDUCIÁRIO(A) ORIGINAL"
    table3.cell(1,1).text = dados['fiduciario']
    table3.cell(2,0).text = "BEM OBJETO DE GARANTIA"
    table3.cell(2,1).text = dados['obj_garantia']
    table3.cell(3,0).text = "VALOR DE AVALIAÇÃO HISTÓRICO"
    table3.cell(3,1).text = f"R$ {dados['valor_obj_inicial']}"
    table3.cell(4,0).text = "VALOR DE AVALIAÇÃO ATUAL"
    table3.cell(4,1).text = f"R$ {dados['valor_obj_atual']}"
    doc.add_paragraph()

    # Seção D
    doc.add_paragraph().add_run("D) PRINCIPAIS CONSTATAÇÕES E APONTAMENTOS").bold = True
    resultado1 = avaliar_titulo(dados['respostas_titulo']) + " " + avaliar_vedacao(dados['respostas_vedacao'])
    doc.add_paragraph().add_run(resultado1)
    resultado2 = avaliar_cessao(dados['respostas_cessao'])
    doc.add_paragraph().add_run(resultado2)
    return doc


def main():
    st.title("Gerador de Parecer Jurídico Simplificado")

    st.header("Dados da Operação")
    solicitante = st.text_input("Solicitante")
    gestora = st.text_input("Gestora")
    numero_celula = st.text_input("Número da Célula (CCB)")
    docs_input = st.text_area("Documentos Recebidos (um por linha)")
    cessionario = st.text_input("Cessionário")
    cedente = st.text_input("Cedente")
    devedor = st.text_input("Devedor(a)")
    data_emissao = st.date_input("Data de Emissão/Referência")
    valor_operacao = st.text_input("Valor da Operação")
    data_primeira = st.date_input("Data da Parcela 1")
    data_ultima = st.date_input("Data da Parcela Final")

    st.header("Dados da Garantia")
    fiduciante = st.text_input("Fiduciante")
    fiduciario = st.text_input("Fiduciário(a) Original")
    obj_garantia = st.text_input("Informações do Objeto de Garantia")
    valor_inicial = st.text_input("Valor de Avaliação Histórico")
    valor_atual = st.text_input("Valor de Avaliação Atual")

    st.header("Título")
    r11 = st.radio("1.1 - Foi submetido à análise título/instrumento de formalização do crédito?", ["SIM", "NÃO", "NÃO SE APLICA"])
    r12 = st.radio("1.2 - O documento mencionado neste item está assinado?", ["SIM", "NÃO"])
    r13 = st.radio("1.3 - A assinatura é eletrônica?", ["SIM", "NÃO"])
    r14 = st.radio("1.4 - Em não sendo eletrônica, o documento está assinado por 2 testemunhas?", ["SIM", "NÃO"])

    st.header("Vedação")
    r21 = st.radio("2.1 - O título contém vedação à cessão do crédito sem a prévia autorização do(a) devedor(a)?", ["SIM", "NÃO", "NÃO SE APLICA"])
    r22 = st.radio("2.2 - Foi obtida autorização do(a) devedor(a) para a realização da cessão?", ["SIM", "NÃO", "NÃO SE APLICA"]) 

    st.header("Cessão")
    r31 = st.radio("3.1 - Foi submetido à análise instrumento de formalização da cessão do crédito?", ["SIM", "NÃO", "NÃO SE APLICA"])
    r32 = st.radio("3.2 - O documento mencionado neste item está assinado?", ["SIM", "NÃO"])
    r33 = st.radio("3.3 - A assinatura é eletrônica?", ["SIM", "NÃO"])
    r34 = st.radio("3.4 - Em não sendo eletrônica, o documento está assinado por 2 testemunhas?", ["SIM", "NÃO"])

    st.header("Comunicação da Cessão")
    r41 = st.radio("4.1 - Foi submetido à análise instrumento de formalização da cessão do crédito?", ["SIM", "NÃO"])
    r42 = st.radio("4.2 - O documento mencionado neste item está assinado?", ["SIM", "NÃO"])
    r43 = st.radio("4.3 - A assinatura é eletrônica?", ["SIM", "NÃO"])
    r44 = st.radio("4.4 - Em não sendo eletrônica, o documento está assinado por 2 testemunhas?", ["SIM", "NÃO"])
    
    if st.button("Gerar e Baixar Parecer"):
        dados = {
            'solicitante': solicitante,
            'gestora': gestora,
            'numero_celula': numero_celula,
            'docs': docs_input.splitlines(),
            'cessionario': cessionario,
            'cedente': cedente,
            'devedor': devedor,
            'data_emissao': data_emissao.strftime("%d/%m/%Y"),
            'valor_operacao': valor_operacao,
            'data_primeira_parcela': data_primeira.strftime("%d/%m/%Y"),
            'data_ultima_parcela': data_ultima.strftime("%d/%m/%Y"),
            'fiduciante': fiduciante,
            'fiduciario': fiduciario,
            'obj_garantia': obj_garantia,
            'valor_obj_inicial': valor_inicial,
            'valor_obj_atual': valor_atual,
            'respostas_titulo': {'1.1': r11, '1.2': r12, '1.3': r13, '1.4': r14},
            'respostas_vedacao': {'2.1': r21, '2.2': r22},
            'respostas_cessao': {'3.1': r31, '3.2': r32, '3.3': r33, '3.4': r34, '4.1' : r41, '4.2' : r42, '4.3' : r43, '4.4' : r44}
        }
        doc = gerar_parecer_garantia(dados)
        output = BytesIO()
        doc.save(output)
        output.seek(0)
        st.download_button(
            label="Baixar Parecer (DOCX)",
            data=output,
            file_name=f"Parecer_CCB_{numero_celula}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

if __name__ == "__main__":
    main()


