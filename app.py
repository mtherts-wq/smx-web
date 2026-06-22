from flask import Flask, render_template, request, send_file, abort
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

app = Flask(__name__)

# ========================
# TEMPO
# ========================
def calcular_tempo(inicio, fim):
    try:
        t1 = datetime.strptime(inicio, "%H:%M")
        t2 = datetime.strptime(fim, "%H:%M")
        diff = t2 - t1
        h = diff.seconds // 3600
        m = (diff.seconds % 3600) // 60
        return f"{h:02}:{m:02}"
    except:
        return "00:00"

# ========================
# VALIDAÇÃO
# ========================
def validar_campos(form):
    campos = [
        "protocolo", "titulo", "atendente", "loja",
        "local", "tecnico", "data", "inicio",
        "fim", "gerente", "descricao"
    ]

    for campo in campos:
        if not form.get(campo):
            return f"Campo obrigatório não preenchido: {campo}"

    return None

# ========================
# NOME
# ========================
def gerar_nome(dados):
    data = dados.get("{{DATA}}", "")
    loja = dados.get("{{LOJA}}", "")
    local = dados.get("{{LOCAL}}", "")

    try:
        dt = datetime.strptime(data, "%d/%m/%Y")
        data_fmt = dt.strftime("%d%m")
    except:
        data_fmt = "0000"

    base = local if loja and loja.strip().upper() == "ZARA" else loja
    base = (base or "arquivo").replace(" ", "_")

    return f"{base}_{data_fmt}"

# ========================
# SUBSTITUIR CAMPOS ✅ FIX
# ========================
def substituir_campos(doc, dados):

    # Parágrafos
    for p in doc.paragraphs:
        for k, v in dados.items():
            if k in p.text:
                for run in p.runs:
                    run.text = run.text.replace(k, str(v or ""))

    # Tabelas
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in dados.items():
                        if k in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(k, str(v or ""))

# ========================
# DOCX
# ========================
def gerar_doc(dados, fotos):

    if not os.path.exists("MODELO_RAT.docx"):
        raise Exception("MODELO_RAT.docx não encontrado")

    doc = Document("MODELO_RAT.docx")

    # ✅ substituição corrigida
    substituir_campos(doc, dados)

    # ✅ fotos antes do gerente
    for i, p in enumerate(doc.paragraphs):
        if "Validado com a Gerente" in p.text:

            table = doc.add_table(rows=0, cols=2)

            row = None
            for idx, f in enumerate(fotos):
                if idx % 2 == 0:
                    row = table.add_row().cells

                if os.path.exists(f):
                    cell = row[idx % 2]
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(f, width=Cm(5), height=Cm(8))

            doc.element.body.insert(i, table._element)

            # ✅ proteção
            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i+1].insert_paragraph_before("")
            else:
                doc.add_paragraph("")

            break

    os.makedirs("temp", exist_ok=True)
    caminho = "temp/saida.docx"
    doc.save(caminho)

    return caminho

# ========================
# PDF
# ========================
def gerar_pdf(dados, fotos):
    os.makedirs("temp", exist_ok=True)
    caminho = "temp/saida.pdf"

    doc = SimpleDocTemplate(caminho)
    styles = getSampleStyleSheet()

    elementos = []

    elementos.append(Paragraph("<b>SMX TI - Chamado Técnico</b>", styles["Title"]))
    elementos.append(Spacer(1, 15))

    tabela = Table([
        ["Protocolo", dados["{{PROTOCOLO}}"]],
        ["Loja", dados["{{LOJA}}"], "Local", dados["{{LOCAL}}"]],
        ["Atendente", dados["{{ATENDENTE}}"], "Técnico", dados["{{TECNICO}}"]],
        ["Data", dados["{{DATA}}"], "Horário", dados["{{HORARIO}}"]],
        ["Tempo", dados["{{TEMPO}}"]]
    ])

    tabela.setStyle([
        ("GRID", (0,0), (-1,-1), 1, colors.black)
    ])

    elementos.append(tabela)
    elementos.append(Spacer(1, 15))

    elementos.append(Paragraph(dados["{{DESCRICAO}}"], styles["Normal"]))
    elementos.append(Spacer(1, 20))

    # fotos
    imgs = []
    for f in fotos:
        if os.path.exists(f):
            imgs.append(Image(f, width=150, height=200))

    linhas = []
    linha = []
    for img in imgs:
        linha.append(img)
        if len(linha) == 2:
            linhas.append(linha)
            linha = []
    if linha:
        linhas.append(linha)

    if linhas:
        elementos.append(Table(linhas))

    elementos.append(Spacer(1, 20))
    elementos.append(Paragraph(f"Validado com a Gerente {dados['{{GERENTE}}']}", styles["Normal"]))

    doc.build(elementos)

    return caminho

# ========================
# AUXILIARES
# ========================
def montar_dados(form, data_fmt):
    inicio = form.get("inicio")
    fim = form.get("fim")

    return {
        "{{PROTOCOLO}}": form.get("protocolo"),
        "{{TITULO}}": form.get("titulo"),
        "{{ATENDENTE}}": form.get("atendente"),
        "{{LOJA}}": form.get("loja"),
        "{{LOCAL}}": form.get("local"),
        "{{TECNICO}}": form.get("tecnico"),
        "{{DATA}}": data_fmt,
        "{{HORARIO}}": f"{inicio} as {fim}",
        "{{TEMPO}}": calcular_tempo(inicio, fim),
        "{{GERENTE}}": form.get("gerente"),
        "{{DESCRICAO}}": form.get("descricao"),
    }

def salvar_fotos(request):
    os.makedirs("temp", exist_ok=True)

    fotos = []
    for f in request.files.getlist("fotos"):
        if f.filename:
            path = os.path.join("temp", f.filename)
            f.save(path)
            fotos.append(path)

    return fotos

# ========================
# ROTAS
# ========================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/gerar", methods=["POST"])
def gerar():

    erro = validar_campos(request.form)
    if erro:
        return erro, 400

    data_raw = request.form.get("data")
    data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")

    dados = montar_dados(request.form, data_fmt)
    fotos = salvar_fotos(request)

    doc = gerar_doc(dados, fotos)
    salvar_historico(dados)

    nome = gerar_nome(dados)

    return send_file(doc, as_attachment=True, download_name=f"{nome}.docx")

@app.route("/pdf", methods=["POST"])
def pdf():

    erro = validar_campos(request.form)
    if erro:
        return erro, 400

    data_raw = request.form.get("data")
    data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")

    dados = montar_dados(request.form, data_fmt)
    fotos = salvar_fotos(request)

    pdf_file = gerar_pdf(dados, fotos)
    nome = gerar_nome(dados)

    return send_file(pdf_file, as_attachment=True, download_name=f"{nome}.pdf")

# ========================
# START
# ========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)