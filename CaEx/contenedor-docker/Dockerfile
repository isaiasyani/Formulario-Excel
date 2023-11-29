FROM python:3.11

WORKDIR /app

COPY CaEx.py .
COPY config.ini .
COPY Crear.png .
COPY formu.py .
COPY Formulario.py .
COPY images.png .
COPY intento2.py .
COPY primera.png .
COPY requirements.txt .
COPY segunda.png .

RUN pip install --no-cache-dir -r requirements.txt

ENV DISPLAY host.docker.internal:0

CMD ["python", "CaEx.py"]