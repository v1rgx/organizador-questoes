import io
import os
import re
import docx
from docx.document import Document
from docx.oxml.text.paragraph import CT_P
from docx.oxml.table import CT_Tbl
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from PIL import Image
import streamlit as st
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph as RLParagraph, Spacer, Image as RLImage, KeepTogether, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfgen import canvas

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Organizador de Questões", layout="wide")

st.title("📚 Organizador Automático de Cadernos de Questões")
st.markdown("Anexe seus arquivos `.docx`, configure as regras do caderno e gere o PDF final com layout e imagens preservadas.")

# --- SIDEBAR / PAINEL DE CONFIGURAÇÕES ---
st.sidebar.header("⚙️ Configurações do Caderno")

escola_nome = st.sidebar.text_input("Nome da Escola", "COLÉGIO DIOCESANO PADRE ANCHIETA")
subtitulo = st.sidebar.text_input("Título da Avaliação", "CADERNO DE QUESTÕES - 2026")
ordem_descritores_str = st.sidebar.text_input("Ordem dos Descritores (separados por vírgula)", "D1, D3, D4, D6")
questoes_por_caderno = st.sidebar.number_input("Questões por Caderno", min_value=1, max_value=50, value=10)

# --- UPLOAD DE ARQUIVOS ---
uploaded_files = st.file_uploader(
    "1. Anexe os arquivos de questões (.docx)", 
    type=["docx"], 
    accept_multiple_files=True
)

def iter_block_items(parent):
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("Tipo de pai não suportado")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extrair_questoes_docx(uploaded_file):
    doc = docx.Document(uploaded_file)
    questoes = []
    
    # Mapeamento das imagens no arquivo
    imagens_rel = {}
    for rel in doc.part.rels.values():
        if "image" in rel.target_ref:
            imagens_rel[rel.rId] = rel.target_part.blob

    questao_atual = {"texto": [], "imagens": [], "descritor": "D1"}
    
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            texto = item.text.strip()
            
            # Captura de imagens
            for run in item.runs:
                if 'drawing' in run._r.xml:
                    for rId in imagens_rel:
                        if rId in run._r.xml:
                            img_data = imagens_rel[rId]
                            questao_atual["imagens"].append(img_data)

            if texto:
                # Detectar descritores ou início de nova questão
                match_desc = re.search(r'\(D(\d+)\)', texto, re.IGNORECASE)
                if match_desc:
                    questao_atual["descritor"] = f"D{match_desc.group(1)}"
                
                if re.match(r'^\d+[\.\)]', texto):  # Ex: "1." ou "1)"
                    if questao_atual["texto"]:
                        questoes.append(questao_atual)
                        questao_atual = {"texto": [], "imagens": [], "descritor": questao_atual.get("descritor", "D1")}
                
                questao_atual["texto"].append(texto)

    if questao_atual["texto"]:
        questoes.append(questao_atual)
        
    return questoes

def gerar_pdf(questoes_organizadas, escola, titulo, tamanho_caderno=10):
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=letter, rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36)
    
    styles = getSampleStyleSheet()
    style_normal = styles['Normal']
    style_normal.fontSize = 10
    style_normal.leading = 13
    
    style_cabecalho = ParagraphStyle(
        'Cabecalho',
        parent=styles['Heading1'],
        alignment=1, # Centralizado
        fontSize=12,
        spaceAfter=12
    )

    story = []
    com_imagem_count = 0
    sem_imagem_count = 0
    
    for idx, q in enumerate(questoes_organizadas):
        num_no_caderno = (idx % tamanho_caderno) + 1
        
        # Cabeçalho no início de cada caderno
        if num_no_caderno == 1:
            if idx > 0:
                story.append(PageBreak())
            story.append(RLParagraph(f"<b>{escola}</b><br/>{titulo}", style_cabecalho))
            story.append(Spacer(1, 10))

        bloco_questao = []
        
        # Ajustar o número da questão para o padrão 1..10
        primeira_linha = re.sub(r'^\d+[\.\)]', f"{num_no_caderno}.", q["texto"][0]) if q["texto"] else ""
        bloco_questao.append(RLParagraph(f"<b>{primeira_linha}</b>", style_normal))
        
        for linha in q["texto"][1:]:
            bloco_questao.append(RLParagraph(linha, style_normal))
            bloco_questao.append(Spacer(1, 4))
            
        # Adicionar imagens centralizadas
        if q["imagens"]:
            com_imagem_count += 1
            for img_bytes in q["imagens"]:
                img = Image.open(io.BytesIO(img_bytes))
                
                # Converter transparência em fundo branco
                if img.mode in ('RGBA', 'LA'):
                    background = Image.new('RGB', img.size, (255, 255, 255))
                    background.paste(img, mask=img.split()[-1])
                    img = background
                
                img_temp = io.BytesIO()
                img.save(img_temp, format='JPEG')
                img_temp.seek(0)
                
                rl_img = RLImage(img_temp, width=300, height=180)
                rl_img.hAlign = 'CENTER'
                bloco_questao.append(Spacer(1, 6))
                bloco_questao.append(rl_img)
                bloco_questao.append(Spacer(1, 6))
        else:
            sem_imagem_count += 1

        story.append(KeepTogether(bloco_questao))
        story.append(Spacer(1, 12))

    doc.build(story)
    pdf_buffer.seek(0)
    
    relatorio = {
        "total_questoes": len(questoes_organizadas),
        "com_imagem": com_imagem_count,
        "sem_imagem": sem_imagem_count
    }
    
    return pdf_buffer, relatorio

# --- PROCESSAMENTO AO CLICAR NO BOTÃO ---
if uploaded_files:
    if st.button("🚀 Processar e Gerar PDF", type="primary"):
        with st.spinner("Lendo arquivos Word, extraindo imagens e organizando as questões..."):
            todas_questoes = []
            for file in uploaded_files:
                questoes = extrair_questoes_docx(file)
                todas_questoes.extend(questoes)

            # Reordenar com base na sequência informada
            ordem_desejada = [d.strip() for d in ordem_descritores_str.split(",")]
            questoes_finais = []
            
            # Intercala uma de cada descritor conforme a sequência informada
            idx = 0
            while len(questoes_finais) < len(todas_questoes):
                desc_alvo = ordem_desejada[idx % len(ordem_desejada)]
                encontrada = False
                for q in todas_questoes:
                    if q not in questoes_finais and q.get("descritor") == desc_alvo:
                        questoes_finais.append(q)
                        encontrada = True
                        break
                if not encontrada:
                    # Se não houver mais daquele descritor específico, pega a próxima disponível
                    for q in todas_questoes:
                        if q not in questoes_finais:
                            questoes_finais.append(q)
                            break
                idx += 1

            pdf_out, relatorio = gerar_pdf(
                questoes_finais, 
                escola_nome, 
                subtitulo, 
                questoes_por_caderno
            )

        st.success("✅ PDF gerado com sucesso!")

        # --- RELATÓRIO DE CONFERÊNCIA ---
        st.subheader("📋 Relatório de Conferência")
        col1, col2, col3 = st.columns(3)
        col1.metric("Total de Questões", relatorio["total_questoes"])
        col2.metric("Questões com Imagem", relatorio["com_imagem"])
        col3.metric("Questões sem Imagem", relatorio["sem_imagem"])

        # --- BOTÃO DE DOWNLOAD ---
        st.download_button(
            label="📥 Baixar Caderno de Questões (PDF)",
            data=pdf_out,
            file_name="Caderno_de_Questoes_Organizado.pdf",
            mime="application/pdf"
        )
