from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

# ✅ PDF
from reportlab.lib.pagesizes import A4
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
# NOME ARQUIVO
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

    if loja and loja.upper() == "ZARA":
        nome = f"{local}_{data_fmt}"
    else:
        nome = f"{loja}_{data_fmt}"

    return (nome or "arquivo").replace(" ", "_")

# ========================
# HISTORICO
# ========================
def salvar_historico(dados):
    arquivo = "historico.xlsx"

    registro = {
        "Data": dados.get("{{DATA}}", ""),
        "Título": dados.get("{{TITULO}}", ""),
        "Loja": dados.get("{{LOJA}}", ""),
        "Atendente": dados.get("{{ATENDENTE}}", ""),
        "Local": dados.get("{{LOCAL}}", ""),
        "Tempo": dados.get("{{TEMPO}}", ""),
        "Gerente": dados.get("{{GERENTE}}", "")
    }

    df_new = pd.DataFrame([registro])

    if os.path.exists(arquivo):
        df_old = pd.read_excel(arquivo)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    df.to_excel(arquivo, index=False)

# ========================
# DOCX
# ========================
def gerar_doc(dados, fotos):

    doc = Document("MODELO_RAT.docx")

    for p in doc.paragraphs:
        texto = "".join(r.text for r in p.runs)
        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v or "")
            p.clear()
            p.add_run(texto)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    texto = "".join(run.text for run in p.runs)
                    if "{{" in texto:
                        for k, v in dados.items():
                            texto = texto.replace(k, v or "")
                        p.clear()
                        p.add_run(texto)

    # ✅ FOTOS antes do gerente
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

            doc.paragraphs[i+1].insert_paragraph_before("")
            break

    os.makedirs("temp", exist_ok=True)
    caminho = "temp/saida.docx"
    doc.save(caminho)

    return caminho

# ========================
# PDF PROFISSIONAL
# ========================
def gerar_pdf(dados, fotos):

    os.makedirs("temp", exist_ok=True)
    caminho = "temp/saida.pdf"

    doc = SimpleDocTemplate(caminho)
    styles = getSampleStyleSheet()

    elementos = []

    # ✅ título
    elementos.append(Paragraph("<b>SMX TI - Chamado Técnico</b>", styles["Title"]))
    elementos.append(Spacer(1, 20))

    # ✅ tabela
    tabela_dados = [
        ["Protocolo", dados.get("{{PROTOCOLO}}", "")],
        ["Loja", dados.get("{{LOJA}}", ""),
         "Solicitação", "Service Desk",
         "Local Atendimento", dados.get("{{LOCAL}}", "")],
        ["Atendente", dados.get("{{ATENDENTE}}", ""),
         "Nome Técnico", dados.get("{{TECNICO}}", ""),
         "Empresa", "SMX"],
        ["Data", dados.get("{{DATA}}", ""),
         "Horário", dados.get("{{HORARIO}}", ""),
         "Tempo", dados.get("{{TEMPO}}", "")]
    ]

    tabela = Table(tabela_dados)
    tabela.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 1, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("ALIGN", (0,0), (-1,-1), "CENTER")
    ]))

    elementos.append(tabela)
    elementos.append(Spacer(1, 20))

    # ✅ descrição
    elementos.append(Paragraph(dados.get("{{DESCRICAO}}", ""), styles["Normal"]))
    elementos.append(Spacer(1, 20))

    # ✅ fotos lado a lado
    imgs = []
    for f in fotos:
        if os.path.exists(f):
            imgs.append(Image(f, width=150, height=200))

    linhas = []
    linha = []
    for i, img in enumerate(imgs):
        linha.append(img)
        if len(linha) == 2:
            linhas.append(linha)
            linha = []

    if linha:
        linhas.append(linha)

    if linhas:
        elementos.append(Table(linhas))

    elementos.append(Spacer(1, 30))

    # ✅ gerente
    elementos.append(Paragraph(
        f"Validado com a Gerente {dados.get('{{GERENTE}}','')}",
        styles["Normal"]
    ))

    doc.build(elementos)

    return caminho

# ========================
# ROTAS
# ========================

@app.route("/")
def home():
    return render_template("index.html")

# ✅ DOCX
@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form

    data_raw = form.get("data")
    data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y") if data_raw else ""

    dados = montar_dados(form, data_fmt)

    fotos = salvar_fotos(request)

    doc = gerar_doc(dados, fotos)
    salvar_historico(dados)

    nome = gerar_nome(dados)

    return send_file(doc, as_attachment=True, download_name=f"{nome}.docx")

# ✅ PDF
@app.route("/pdf", methods=["POST"])
def pdf():
    form = request.form

    data_raw = form.get("data")
    data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y") if data_raw else ""

    dados = montar_dados(form, data_fmt)

    fotos = salvar_fotos(request)

    pdf_file = gerar_pdf(dados, fotos)

    nome = gerar_nome(dados)

    return send_file(pdf_file, as_attachment=True, download_name=f"{nome}.pdf")

# ========================
# AUXILIARES
# ========================
def montar_dados(form, data_fmt):
    inicio = form.get("inicio")
    fim = form.get("fim")
    tempo = calcular_tempo(inicio, fim)

    return {
        "{{PROTOCOLO}}": form.get("protocolo"),
        "{{TITULO}}": form.get("titulo"),
        "{{ATENDENTE}}": form.get("atendente"),
        "{{LOJA}}": form.get("loja"),
        "{{LOCAL}}": form.get("local"),
        "{{TECNICO}}": form.get("tecnico"),
        "{{DATA}}": data_fmt,
        "{{HORARIO}}": f"{inicio} as {fim}",
        "{{TEMPO}}": tempo,
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
# START
# ========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)