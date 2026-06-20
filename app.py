from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

# ✅ PDF
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

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

    if not os.path.exists("MODELO_RAT.docx"):
        raise Exception("MODELO_RAT.docx não encontrado")

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

    if fotos:
        doc.add_paragraph("\nFotos:\n")
        for f in fotos:
            if os.path.exists(f):
                doc.add_picture(f, width=Cm(12))

    os.makedirs("temp", exist_ok=True)
    caminho = os.path.join("temp", "saida.docx")
    doc.save(caminho)

    return caminho

# ========================
# PDF (SIMPLIFICADO)
# ========================
def gerar_pdf_simples(dados):
    caminho = "temp/saida.pdf"
    os.makedirs("temp", exist_ok=True)

    c = canvas.Canvas(caminho, pagesize=A4)

    y = 800

    for k, v in dados.items():
        if "{{" in k:
            texto = f"{k}: {v}"
            c.drawString(50, y, texto)
            y -= 20

    c.save()
    return caminho

# ========================
# ROTAS
# ========================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        form = request.form

        data_raw = form.get("data")
        if data_raw:
            try:
                data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")
            except:
                data_fmt = ""
        else:
            data_fmt = ""

        inicio = form.get("inicio")
        fim = form.get("fim")

        tempo = calcular_tempo(inicio, fim)

        dados = {
            "{{PROTOCOLO}}": form.get("protocolo"),
            "{{TITULO}}": form.get("titulo"),
            "{{ATENDENTE}}": form.get("atendente"),
            "{{LOJA}}": form.get("loja"),
            "{{LOCAL}}": form.get("local"),
            "{{TECNICO}}": form.get("tecnico"),
            "{{DATA}}": data_fmt,
            "{{HORARIO}}": f"{inicio} as {fim}" if inicio and fim else "",
            "{{TEMPO}}": tempo,
            "{{GERENTE}}": form.get("gerente"),
            "{{DESCRICAO}}": form.get("descricao"),
            "LOCAL": form.get("local") or "arquivo"
        }

        os.makedirs("temp", exist_ok=True)

        fotos = []
        for f in request.files.getlist("fotos"):
            if f.filename:
                path = os.path.join("temp", f.filename)
                f.save(path)
                fotos.append(path)

        doc = gerar_doc(dados, fotos)
        salvar_historico(dados)

        return send_file(doc, as_attachment=True)

    except Exception as e:
        return f"Erro: {str(e)}"

# ✅ PDF FUNCIONAL
@app.route("/pdf", methods=["POST"])
def pdf():
    form = request.form

    dados = {k: v for k, v in form.items()}
    caminho = gerar_pdf_simples(dados)

    return send_file(caminho, as_attachment=True)

# ✅ EXCEL
@app.route("/excel")
def excel():
    mes = request.args.get("mes")

    if not os.path.exists("historico.xlsx"):
        return "Sem registros ainda"

    df = pd.read_excel("historico.xlsx")

    if mes:
        ano, mes_num = mes.split("-")

        def filtro(d):
            try:
                data = datetime.strptime(d, "%d/%m/%Y")
                return data.year == int(ano) and data.month == int(mes_num)
            except:
                return False

        df = df[df["Data"].apply(filtro)]

    caminho = "relatorio.xlsx"
    df.to_excel(caminho, index=False)

    return send_file(caminho, as_attachment=True)

# ========================
# START
# ========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)