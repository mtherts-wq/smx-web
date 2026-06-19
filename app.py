from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

def gerar_doc(dados):
    doc = Document("MODELO_RAT.docx")

    for p in doc.paragraphs:
        for k, v in dados.items():
            if k in p.text:
                p.text = p.text.replace(k, v)

    nome = f"{dados['{{LOCAL}}']}_{datetime.now().strftime('%d%m')}.docx"
    caminho = os.path.join("temp", nome)

    os.makedirs("temp", exist_ok=True)

    doc.save(caminho)
    return caminho


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form

    dados = {
        "{{PROTOCOLO}}": form.get("protocolo"),
        "{{TITULO}}": form.get("titulo"),
        "{{LOCAL}}": form.get("local"),
        "{{DESCRICAO}}": form.get("descricao")
    }

    caminho = gerar_doc(dados)

    return send_file(caminho, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)