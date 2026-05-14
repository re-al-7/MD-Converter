FROM python:3.11-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    tesseract-ocr \
    tesseract-ocr-spa \
    libgl1 \
    libglib2.0-0 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN mkdir -p /app/correos /app/md_output

ENV WATCH_DIR=/app/correos
ENV OUTPUT_DIR=/app/md_output
ENV PORT=3200

EXPOSE ${PORT}

VOLUME ["/app/correos", "/app/md_output"]

CMD ["python", "converter_ui.py"]
