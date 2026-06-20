from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

# ========================
# CALCULAR TEMPO
# ========================

def calcular_tempo(inicio, fim):
    t1 = datetime.strptime(inicio, "%H:%M")
    t2 = datetime.strptime(fim, "%H:%M")

    diff = t2 - t1
    h = diff.seconds // 3600
    m = (diff.seconds % 3600) // 60

    return f"{h:02}:{m:02}"

# ========================
# GERAR DOC
# ========================

def gerar_doc(dados):
    doc = Document("MODELO_RAT.docx")

    # ✅ PARÁGRAFOS
    for p in doc.paragraphs:
        texto = "".join(run.text for run in p.runs)

        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v)

            p.clear()
            p.add_run(texto)

    # ✅ TABELAS
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    texto = "".join(run.text for run in p.runs)

                    if "{{" in texto:
                        for k, v in dados.items():
                            texto = texto.replace(k, v)

                        p.clear()
                        p.add_run(texto)

    nome = f"{dados['LOCAL']}_{datetime.now().strftime('%d%m')}.docx"
    caminho = os.path.join("temp", nome)

    os.makedirs("temp", exist_ok=True)
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

    # ✅ DATA FORMATADA
    data_raw = form.get("data")
    data_formatada = datetime.strptime(data_raw, "%Y-%m-%d").strftime("%d/%m/%Y")

    # ✅ TEMPO
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

        # ✅ nome do arquivo
        "LOCAL": form.get("local")
    }

    caminho = gerar_doc(dados)

    return send_file(caminho, as_attachment=True)

# ========================
# START
# ========================

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
