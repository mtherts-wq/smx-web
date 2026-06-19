from flask import Flask, render_template, request, send_file
from docx import Document
from datetime import datetime
import os

app = Flask(__name__)

def gerar_doc(dados):
    doc = Document("MODELO_RAT.docx")

    for p in doc.paragraphs:
        texto = "".join(run.text for run in p.runs)
        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v)
            p.clear()
            p.add_run(texto)

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


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/gerar", methods=["POST"])
def gerar():
    form = request.form

    dados = {
        "{{PROTOCOLO}}": form.get("protocolo"),
        "{{TITULO}}": form.get("titulo"),
        "{{ATENDENTE}}": form.get("atendente"),
        "{{LOJA}}": form.get("loja"),
        "{{LOCAL}}": form.get("local"),
        "{{TECNICO}}": form.get("tecnico"),
        "{{DATA}}": form.get("data"),
        "{{HORARIO}}": f"{form.get('inicio')} as {form.get('fim')}",
        "{{TEMPO}}": "00:00",
        "{{GERENTE}}": form.get("gerente"),
        "{{DESCRICAO}}": form.get("descricao"),
        "LOCAL": form.get("local")
    }

    caminho = gerar_doc(dados)

    return send_file(caminho, as_attachment=True)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
