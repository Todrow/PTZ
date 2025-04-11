FROM python:3.11-alpine

WORKDIR /app

# копируем requirements.txt и устанавливаем зависимости
COPY requirements.txt .

# Устанавливаем системные зависимости через apk 
RUN apk add --no-cache --virtual .build-deps \
    gcc \
    musl-dev \
    postgresql-dev \
    && pip install --no-cache-dir -r requirements.txt \
    && apk del .build-deps

COPY . .

CMD ["gunicorn", "--bind", "0.0.0.0:8000", "excel.wsgi:application"]