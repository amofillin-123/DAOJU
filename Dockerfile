FROM python:3.9-slim

WORKDIR /app

COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

COPY . .

RUN apt-get update && apt-get install -y fonts-wqy-zenhei && apt-get clean && rm -rf /var/lib/apt/lists/*

ENV PORT=8080
ENV FONT_PATH=/usr/share/fonts/truetype/wqy/wqy-zenhei.ttc

EXPOSE 8080

CMD ["gunicorn", "--bind", "0.0.0.0:8080", "app:app"]
