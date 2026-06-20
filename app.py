from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

app = Flask(__name__)

# ========================
# CALCULAR TEMPO
# ========================
def calcular_tempo(inicio, fim):
    t1 = datetime.strptime(inicio, "%H:%M")
    t2 = datetime.strptime(fim, "%H:%M")
    diff = t2 - t1
    horas = diff.seconds // 3600
    minutos = (diff.seconds % 3600) // 60
    return f"{horas:02}:{minutos:02}"

# ========================
# SALVAR HISTÓRICO
# ========================
def salvar_historico(dados):
    arquivo = "historico.xlsx"

    registro = {
        "Data": dados["{{DATA}}"],
        "Título": dados["{{TITULO}}"],
        "Loja": dados["{{LOJA}}"],
        "Atendente": dados["{{ATENDENTE}}"],
        "Local": dados["{{LOCAL}}"],
        "Tempo": dados["{{TEMPO}}"],
        "Gerente": dados["{{GERENTE}}"]
    }

    df_new = pd.DataFrame([registro])

    if os.path.exists(arquivo):
        df_old = pd.read_excel(arquivo)
        df = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df = df_new

    df.to_excel(arquivo, index=False)

# ========================
# GERAR DOCX + FOTOS
# ========================
def gerar_doc(dados, fotos):
    doc = Document("MODELO_RAT.docx")

    # TEXTO
    for p in doc.paragraphs:
        texto = "".join(run.text for run in p.runs)
        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v)
            p.clear()
            p.add_run(texto)

    # TABELAS
    for t in doc.tables:
        for row in t.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texto = "".join(run.text for run in p.runs)
                    if "{{" in texto:
                        for k, v in dados.items():
                            texto = texto.replace(k, v)
                        p.clear()
                        p.add_run(texto)

    # ✅ FOTOS
    if fotos:
        doc.add_paragraph("\nFotos do Atendimento:\n")
        for imagem in fotos:
            doc.add_picture(imagem, width=Cm(12))

    os.makedirs("temp", exist_ok=True)
    caminho = os.path.join("temp", "saida.docx")
    doc.save(caminho)

    return caminho

# ========================
# ROTAS
# ========================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form

    data_raw = form.get("data")
    data_formatada = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")

    tempo = calcular_tempo(form.get("inicio"), form.get("fim"))

    dados = {
        "{{PROTOCOLO}}": form.get("protocolo"),
        "{{TITULO}}": form.get("titulo"),
        "{{ATENDENTE}}": form.get("atendente"),
        "{{LOJA}}": form.get("loja"),
        "{{LOCAL}}": form.get("local"),
        "{{TECNICO}}": form.get("tecnico"),
        "{{DATA}}": data_formatada,
        "{{HORARIO}}": f"{form.get('inicio')} as {form.get('fim')}",
        "{{TEMPO}}": tempo,
        "{{GERENTE}}": form.get("gerente"),
        "{{DESCRICAO}}": form.get("descricao"),
        "LOCAL": form.get("local")
    }

    # FOTOS
    files = request.files.getlist("fotos")
    fotos = []

    for f in files:
        if f.filename:
            caminho = os.path.join("temp", f.filename)
            f.save(caminho)
            fotos.append(caminho)

    caminho_doc = gerar_doc(dados, fotos)

    salvar_historico(dados)

    return send_file(caminho_doc, as_attachment=True)

@app.route("/excel")
def excel():
    return send_file("historico.xlsx", as_attachment=True)

# ========================
# START
# ========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)