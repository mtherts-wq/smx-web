from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Cm
from datetime import datetime
import os
import pandas as pd

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
# NOME DO ARQUIVO
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

    if not os.path.exists("MODELO_RAT.docx"):
        raise Exception("MODELO_RAT.docx não encontrado")

    doc = Document("MODELO_RAT.docx")

    # ✅ substituir campos
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

    # ========================
    # ✅ FOTOS AJUSTADAS (CORRIGIDO)
    # ========================
    if fotos:
        for p in doc.paragraphs:
            if "Validado com a Gerente" in p.text:

                # ✅ título
                p.insert_paragraph_before("Fotos:")

                # ✅ tabela lado a lado
                table = doc.add_table(rows=0, cols=2)

                row = None
                for i, f in enumerate(fotos):
                    if i % 2 == 0:
                        row = table.add_row().cells

                    if os.path.exists(f):
                        cell = row[i % 2]
                        paragraph = cell.paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(f, width=Cm(5), height=Cm(8))

                # ✅ espaçamento antes do gerente
                p.insert_paragraph_before("")
                p.insert_paragraph_before("")

                break

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

# ✅ DOCX
@app.route("/gerar", methods=["POST"])
def gerar():
    try:
        form = request.form

        data_raw = form.get("data")
        if data_raw:
            try:
                data_fmt = datetime.strptime(
                    data_raw, "%Y-%m-%d"
                ).strftime("%d/%m/%Y")
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

        nome = gerar_nome(dados)

        return send_file(
            doc,
            as_attachment=True,
            download_name=f"{nome}.docx"
        )

    except Exception as e:
        return f"Erro: {str(e)}"

# ✅ “PDF” (BAIXA DOCX — ESTÁVEL)
@app.route("/pdf", methods=["POST"])
def pdf():
    try:
        form = request.form

        data_raw = form.get("data")
        if data_raw:
            try:
                data_fmt = datetime.strptime(
                    data_raw, "%Y-%m-%d"
                ).strftime("%d/%m/%Y")
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
        }

        os.makedirs("temp", exist_ok=True)

        fotos = []
        for f in request.files.getlist("fotos"):
            if f.filename:
                path = os.path.join("temp", f.filename)
                f.save(path)
                fotos.append(path)

        doc = gerar_doc(dados, fotos)

        nome = gerar_nome(dados)

        return send_file(
            doc,
            as_attachment=True,
            download_name=f"{nome}.docx"
        )

    except Exception as e:
        return f"Erro PDF: {str(e)}"

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