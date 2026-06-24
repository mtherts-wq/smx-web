from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd
import requests
import base64

app = Flask(__name__)

PDFSHIFT_API_KEY = os.environ.get("PDFSHIFT_API_KEY")

if not PDFSHIFT_API_KEY:
    print("⚠️ PDFSHIFT_API_KEY não configurada no Render")


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
        "protocolo","titulo","atendente","loja",
        "local","tecnico","data","inicio",
        "fim","gerente","descricao"
    ]
    for c in campos:
        if not form.get(c):
            return f"Campo obrigatório: {c}"
    return None


# ========================
# NOME (REGRA ZARA)
# ========================
def gerar_nome(dados):

    data = dados.get("{{DATA}}","")
    loja = (dados.get("{{LOJA}}") or "").strip()
    local = (dados.get("{{LOCAL}}") or "").strip()

    try:
        dt = datetime.strptime(data,"%d/%m/%Y")
        data_fmt = dt.strftime("%d%m")
    except:
        data_fmt = "0000"

    if loja.upper() == "ZARA":
        base = local
    else:
        base = loja

    base = (base or "arquivo").replace(" ","_")
    return f"{base}_{data_fmt}"


# ========================
# HISTÓRICO
# ========================
def salvar_historico(dados):

    arquivo = "historico.xlsx"

    registro = {
        "Data": dados.get("{{DATA}}"),
        "Título": dados.get("{{TITULO}}"),
        "Loja": dados.get("{{LOJA}}"),
        "Atendente": dados.get("{{ATENDENTE}}"),
        "Local": dados.get("{{LOCAL}}"),
        "Tempo": dados.get("{{TEMPO}}"),
        "Gerente": dados.get("{{GERENTE}}")
    }

    df_new = pd.DataFrame([registro])

    if os.path.exists(arquivo):
        df_old = pd.read_excel(arquivo)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    df.to_excel(arquivo, index=False)


# ========================
# SUBSTITUIÇÃO
# ========================
def substituir_campos(doc, dados):

    for p in doc.paragraphs:
        for k, v in dados.items():
            if k in p.text:
                for run in p.runs:
                    run.text = run.text.replace(k, str(v or ""))

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for k, v in dados.items():
                        if k in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(k, str(v or ""))


# ========================
# DOCX
# ========================
def gerar_doc(dados, fotos):

    doc = Document("MODELO_RAT.docx")
    substituir_campos(doc, dados)

    for i, p in enumerate(doc.paragraphs):
        if "Validado com a Gerente" in p.text:
            table = doc.add_table(rows=0, cols=2)

            for idx, f in enumerate(fotos):
                if idx % 2 == 0:
                    row = table.add_row().cells
                if os.path.exists(f):
                    cell = row[idx % 2]
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(f, width=Cm(5), height=Cm(8))

            doc.element.body.insert(i, table._element)
            break

    os.makedirs("temp", exist_ok=True)
    path = "temp/saida.docx"
    doc.save(path)

    return path


# ========================
# PDF (PDFSHIFT)
# ========================
def converter_pdf(doc_path):

    if not PDFSHIFT_API_KEY:
        raise Exception("PDFSHIFT_API_KEY não configurada")

    with open(doc_path, "rb") as f:
        encoded = base64.b64encode(f.read()).decode("utf-8")

    response = requests.post(
        "https://api.pdfshift.io/v3/convert/docx",
        headers={
            "Authorization": f"Bearer {PDFSHIFT_API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "source": encoded,
            "landscape": False,
            "use_print": False
        }
    )

    if response.status_code != 200:
        raise Exception(f"Erro PDFShift: {response.text}")

    pdf_path = doc_path.replace(".docx", ".pdf")

    with open(pdf_path, "wb") as f:
        f.write(response.content)

    return pdf_path


# ========================
# DADOS
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


# ========================
# FOTOS
# ========================
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

    data_fmt = datetime.strptime(
        request.form.get("data"), "%Y-%m-%d"
    ).strftime("%d/%m/%Y")

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

    data_fmt = datetime.strptime(
        request.form.get("data"), "%Y-%m-%d"
    ).strftime("%d/%m/%Y")

    dados = montar_dados(request.form, data_fmt)
    fotos = salvar_fotos(request)

    doc = gerar_doc(dados, fotos)
    salvar_historico(dados)

    pdf_path = converter_pdf(doc)
    nome = gerar_nome(dados)

    return send_file(pdf_path, as_attachment=True, download_name=f"{nome}.pdf")


@app.route("/excel")
def excel():

    if not os.path.exists("historico.xlsx"):
        return "Sem registros ainda", 400

    df = pd.read_excel("historico.xlsx")

    os.makedirs("temp", exist_ok=True)
    path = "temp/relatorio.xlsx"

    df.to_excel(path, index=False)
    return send_file(path, as_attachment=True, download_name="relatorio.xlsx")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)