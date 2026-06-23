# Usar imagem Python com suporte a LibreOffice
FROM python:3.11-slim

# Instalar LibreOffice
RUN apt-get update && \
    apt-get install -y libreoffice && \
    apt-get clean

# Criar diretório da app
WORKDIR /app

# Copiar requirements
COPY requirements.txt .

# Instalar dependências
RUN pip install --no-cache-dir -r requirements.txt

# Copiar todo o código
COPY . .

# Expor a porta do Flask
EXPOSE 10000

# Entrypoint
CMD ["python", "app.py"]