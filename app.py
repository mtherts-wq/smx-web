def gerar_doc(dados):
    doc = Document("MODELO_RAT.docx")

    # ✅ PARÁGRAFOS
    for p in doc.paragraphs:
        texto = "".join(run.text for run in p.runs)

        if "{{" in texto:
            for k, v in dados.items():
                texto = texto.replace(k, v)

            p.clear()
            run = p.add_run(texto)

    # ✅ TABELAS (SEM PERDER FORMATAÇÃO)
    for tabela in doc.tables:
        for row in tabela.rows:
            for cell in row.cells:
                for p in cell.paragraphs:

                    texto = "".join(run.text for run in p.runs)

                    if "{{" in texto:
                        for k, v in dados.items():
                            texto = texto.replace(k, v)

                        p.clear()
                        run = p.add_run(texto)

    # ✅ NOME LIMPO
    nome = f"{dados['LOCAL']}_{datetime.now().strftime('%d%m')}.docx"

    caminho = os.path.join("temp", nome)
    os.makedirs("temp", exist_ok=True)

    doc.save(caminho)

    return caminho