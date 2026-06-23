from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd
import requests
import time
import subprocess

app = Flask(__name__)

API_KEY = os.environ.get("API_KEY")

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
    for campo in campos:
        if not form.get(campo):
            return f"Campo obrigatório: {campo}"
    return None

# ========================
# NOME
# ========================
def gerar_nome(dados):
    data = dados.get("{{DATA}}","")
    loja = dados.get("{{LOJA}}","")
    local = dados.get("{{LOCAL}}","")

    try:
        dt = datetime.strptime(data,"%d/%m/%Y")
        data_fmt = dt.strftime("%d%m")
    except:
        data_fmt = "0000"

    if loja and loja.upper() == "ZARA":
        base = local
    else:
        base = loja

    base = (base or "arquivo").replace(" ","_")

    return f"{base}_{data_fmt}"

# ========================
# HISTORICO
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
# SUBSTITUIR CAMPOS
# ========================
def substituir_campos(doc,dados):

    for p in doc.paragraphs:
        for k,v in dados.items():
            if k in p.text:
                for run in p.runs:
                    run.text = run.text.replace(k,str(v or ""))

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for k,v in dados.items():
                        if k in p.text:
                            for run in p.runs:
                                run.text = run.text.replace(k,str(v or ""))

# ========================
# DOCX
# ========================
def gerar_doc(dados,fotos):

    doc = Document("MODELO_RAT.docx")

    substituir_campos(doc,dados)

    for i,p in enumerate(doc.paragraphs):
        if "Validado com a Gerente" in p.text:

            table = doc.add_table(rows=0,cols=2)

            row = None
            for idx,f in enumerate(fotos):
                if idx % 2 == 0:
                    row = table.add_row().cells

                if os.path.exists(f):
                    cell = row[idx % 2]
                    run = cell.paragraphs[0].add_run()
                    run.add_picture(f,width=Cm(5),height=Cm(8))

            doc.element.body.insert(i,table._element)

            if i + 1 < len(doc.paragraphs):
                doc.paragraphs[i+1].insert_paragraph_before("")
            else:
                doc.add_paragraph("")

            break

    os.makedirs("temp",exist_ok=True)
    path = "temp/saida.docx"
    doc.save(path)

    return path

# ========================
# API CONVERTAPI
# ========================
def converter_pdf(doc_path):
    with open(doc_path, "rb") as f:
        response = requests.post(
            "https://v2.convertapi.com/convert/docx/to/pdf",
            params={"Secret": API_KEY},
            files={"File": f}
        )

    result = response.json()

    pdf_url = result["Files"][0]["Url"]
    pdf_path = doc_path.replace(".docx", ".pdf")

    pdf_bytes = requests.get(pdf_url).content

    with open(pdf_path, "wb") as f:
        f.write(pdf_bytes)

    return pdf_path

# ========================
# AUX
# ========================
def montar_dados(form,data_fmt):

    inicio = form.get("inicio")
    fim = form.get("fim")

    return {
        "{{PROTOCOLO}}":form.get("protocolo"),
        "{{TITULO}}":form.get("titulo"),
        "{{ATENDENTE}}":form.get("atendente"),
        "{{LOJA}}":form.get("loja"),
        "{{LOCAL}}":form.get("local"),
        "{{TECNICO}}":form.get("tecnico"),
        "{{DATA}}":data_fmt,
        "{{HORARIO}}":f"{inicio} as {fim}",
        "{{TEMPO}}":calcular_tempo(inicio,fim),
        "{{GERENTE}}":form.get("gerente"),
        "{{DESCRICAO}}":form.get("descricao"),
    }

def salvar_fotos(request):
    os.makedirs("temp",exist_ok=True)

    fotos=[]
    for f in request.files.getlist("fotos"):
        if f.filename:
            path=os.path.join("temp",f.filename)
            f.save(path)
            fotos.append(path)

    return fotos

# ========================
# ROTAS
# ========================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/gerar",methods=["POST"])
def gerar():

    erro = validar_campos(request.form)
    if erro:
        return erro,400

    data_raw = request.form.get("data")
    data_fmt = datetime.strptime(data_raw,"%Y-%m-%d").strftime("%d/%m/%Y")

    dados = montar_dados(request.form,data_fmt)
    fotos = salvar_fotos(request)

    doc = gerar_doc(dados,fotos)

    # ✅ SALVAR HISTORICO
    salvar_historico(dados)

    nome = gerar_nome(dados)

    return send_file(doc,as_attachment=True,download_name=f"{nome}.docx")

@app.route("/pdf",methods=["POST"])
def pdf():

    erro = validar_campos(request.form)
    if erro:
        return erro,400

    data_raw = request.form.get("data")
    data_fmt = datetime.strptime(data_raw,"%Y-%m-%d").strftime("%d/%m/%Y")

    dados = montar_dados(request.form,data_fmt)
    fotos = salvar_fotos(request)

    doc = gerar_doc(dados,fotos)
    nome = gerar_nome(dados)

    try:
        pdf = converter_pdf(doc)
        return send_file(pdf,as_attachment=True,download_name=f"{nome}.pdf")
    except:
        return send_file(doc,as_attachment=True,download_name=f"{nome}.docx")

@app.route("/excel")
def excel():

    mes = request.args.get("mes")

    if not os.path.exists("historico.xlsx"):
        return "Sem registros ainda", 400

    df = pd.read_excel("historico.xlsx")

    if mes:
        try:
            ano, mes_num = mes.split("-")

            def filtro(data_str):
                try:
                    d = datetime.strptime(data_str,"%d/%m/%Y")
                    return d.year == int(ano) and d.month == int(mes_num)
                except:
                    return False

            df = df[df["Data"].apply(filtro)]

        except:
            return "Erro ao processar mês", 400

    caminho = "temp/relatorio.xlsx"
    os.makedirs("temp", exist_ok=True)

    df.to_excel(caminho,index=False)

    return send_file(caminho,as_attachment=True,download_name="relatorio.xlsx")

# ========================
# START
# ========================
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 10000))
    app.run(host="0.0.0.0", port=port, debug=False)