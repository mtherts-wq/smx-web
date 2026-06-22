from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd
import requests
import time

app = Flask(__name__)

# 🔥 COLOQUE SUA API KEY AQUI
API_KEY = "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNmY0YWU3MzAyMjk0YTMzMTg5ZmNiMTUzYzg0Mjg0MjJhMjViNzgwYWZkMjFkN2FmMTliY2E1YjlmNjg5NjNiMzVkZWNjNjYzZWFiM2JhYzAiLCJpYXQiOjE3ODIxNjYwNDcuMjM2MzEyLCJuYmYiOjE3ODIxNjYwNDcuMjM2MzEzLCJleHAiOjQ5Mzc4Mzk2NDcuMjMxMzU3LCJzdWIiOiI3NjA2OTMyOSIsInNjb3BlcyI6WyJ0YXNrLnJlYWQiLCJ0YXNrLndyaXRlIl19.gXK-freAXtMNbA74Wd7NBBGdoaLXK_Mb36tv2HfAcB9He1tRk75JDU5cdl0NIzkcriAp1CINPbmNENTZGSPqh-cV4fw0DXz0SAimUVXiBp2a5edZ8XRp9Lv2fxuzTYR8Gnq9WT-How-ZNANU2jaVJDAFeRWsrw0eqC1g5ZXUOyAWQdi4vo9c1tlk2xGTrLnJbQ-zEdCEpelPYb9JaBvWr_n-jWi9hQ6_OEKLab9ktcvFfy3abi7Xlc58lSUjwSn7W2bqs8wPpMnCWqnL--fCmkI9QNwNLx8b6a24xKzyh3CBKGb7-1sZdDwOoVUMwMN13rczMwmgpcKRfp_u-HkvTsrVzzhj5B-a8QTUV0QHIeb7mils3hOhuK4vyHUw2QUGcsulNX0Vu_6xO5A2Wy-ZeZXNzcs8wdmORgshgqXVsQQaOog2KpDdIkeVG839G1b2Qx6sM4Wynou6oYZqETE14JIwjlNU04fERBzUHhAL59EEI6wq-P3xBKj6GoCAHfmYnjpAIAxQSohuWNREM4jeP_ZFaYOXvi-IdaxOmZ5oAgJwghe9c-1PCNKArMhKda8z7Zf42O7KQ4VBWjo2GlZjN-PIAzmK2rmaMvW8xNTDbHhTRHhwyYk9jKgCsHonGOtJ_PKaavA1lh18lcLXFjLg7b_P2YDkM41645VvmLS3hSs"

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

    base = local if loja.upper()=="ZARA" else loja
    base = (base or "arquivo").replace(" ","_")

    return f"{base}_{data_fmt}"

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

            if i+1 < len(doc.paragraphs):
                doc.paragraphs[i+1].insert_paragraph_before("")
            else:
                doc.add_paragraph("")

            break

    os.makedirs("temp",exist_ok=True)
    path = "temp/saida.docx"
    doc.save(path)

    return path

# ========================
# CLOUDCONVERT
# ========================
def converter_pdf(doc_path):

    response = requests.post(
        "https://api.cloudconvert.com/v2/jobs",
        headers={
            "Authorization": f"Bearer {API_KEY}",
            "Content-Type": "application/json"
        },
        json={
            "tasks":{
                "upload":{
                    "operation":"import/upload"
                },
                "convert":{
                    "operation":"convert",
                    "input":"upload",
                    "output_format":"pdf"
                },
                "export":{
                    "operation":"export/url",
                    "input":"convert"
                }
            }
        }
    ).json()

    upload_task = next(t for t in response["data"]["tasks"] if t["name"]=="upload")

    upload_url = upload_task["result"]["form"]["url"]
    upload_params = upload_task["result"]["form"]["parameters"]

    with open(doc_path,"rb") as f:
        requests.post(upload_url,data=upload_params,files={"file":f})

    job_id = response["data"]["id"]

    # espera conversão
    while True:
        job_status = requests.get(
            f"https://api.cloudconvert.com/v2/jobs/{job_id}",
            headers={"Authorization":f"Bearer {API_KEY}"}
        ).json()

        if job_status["data"]["status"]=="finished":
            break

        time.sleep(2)

    export = next(t for t in job_status["data"]["tasks"] if t["name"]=="export")

    file_url = export["result"]["files"][0]["url"]

    pdf_path = doc_path.replace(".docx",".pdf")

    pdf_bytes = requests.get(file_url).content

    with open(pdf_path,"wb") as f:
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
    except Exception as e:
        # ✅ fallback seguro
        return send_file(doc,as_attachment=True,download_name=f"{nome}.docx")

# ========================
# START
# ========================
if __name__=="__main__":
    app.run(host="0.0.0.0",port=10000)