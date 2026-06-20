from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

app = Flask(__name__)

# ========================
# CALCULO TEMPO
# ========================
def calcular_tempo(inicio, fim):
    try:
        t1 = datetime.strptime(inicio, "%H:%M")
        t2 = datetime.strptime(fim, "%H:%M")
        diff = t2 - t1

        horas = diff.seconds // 3600
        minutos = (diff.seconds % 3600) // 60

        return f"{horas:02}:{minutos:02}"
    except:
        return "00:00"

# ========================
# SALVAR HISTORICO
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
# GERAR DOCX
# ========================
def gerar_doc(dados, fotos):
    doc = Document("MODELO_RAT.docx")

    # TEXTO NORMAL
    for p in doc.paragraphs:
        texto = "".join(run.text for run in p.runs)

        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v if v else "")

            p.clear()
            p.add_run(texto)

    # TABELAS
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for p in cell.paragraphs:

                    texto = "".join(run.text for run in p.runs)

                    if "{{" in texto:
                        for k, v in dados.items():
                            texto = texto.replace(k, v if v else "")

                        p.clear()
                        p.add_run(texto)

    # FOTOS
    if fotos:
        doc.add_paragraph("\nFotos do Atendimento:\n")

        for img in fotos:
            if os.path.exists(img):
                doc.add_picture(img, width=Cm(12))

    # SALVAR
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

# ========================
# GERAR DOCX
# ========================
@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form

    # DATA
    data_raw = form.get("data")

    if data_raw:
        try:
            data_fmt = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")
        except:
            data_fmt = ""
    else:
        data_fmt = ""

    # TEMPO
    inicio = form.get("inicio")
    fim = form.get("fim")

    tempo = calcular_tempo(inicio, fim)

    # DADOS
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

    # FOTOS
    fotos = []
    files = request.files.getlist("fotos")

    for f in files:
        if f and f.filename:
            caminho = os.path.join("temp", f.filename)
            f.save(caminho)
            fotos.append(caminho)

    # GERAR DOCUMENTO
    caminho_doc = gerar_doc(dados, fotos)

    # SALVAR HISTORICO
    salvar_historico(dados)

    return send_file(caminho_doc, as_attachment=True)

# ========================
# EXCEL COM FILTRO
# ========================
@app.route("/excel")
def excel():
    mes = request.args.get("mes")

    if not os.path.exists("historico.xlsx"):
        return "Sem dados"

    df = pd.read_excel("historico.xlsx")

    # FILTRO POR MES
    if mes:
        try:
            ano, mes_num = mes.split("-")

            def filtro(data_str):
                try:
                    d = datetime.strptime(data_str, "%d/%m/%Y")
                    return d.year == int(ano) and d.month == int(mes_num)
                except:
                    return False

            df = df[df["Data"].apply(filtro)]

        except:
            pass

    caminho = "relatorio.xlsx"
    df.to_excel(caminho, index=False)

    return send_file(caminho, as_attachment=True)

# ========================
# START
# ========================
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)