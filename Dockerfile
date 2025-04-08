# Используем официальный Python образ как базовый
FROM python:3.11-bookworm

# Устанавливаем переменные среды
ENV PYTHONDONTWRITEBYTECODE 1
ENV PYTHONUNBUFFERED 1

# Создаем и переходим в рабочую директорию
WORKDIR /app

# Устанавливаем зависимости системы
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libpq-dev \
    && rm -rf /var/lib/apt/lists/*

# Устанавливаем зависимости Python
COPY requirements.txt .
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# Копируем проект
COPY . .

# Собираем статические файлы (если нужно)
# RUN python manage.py collectstatic --noinput

# Команда для запуска (может быть изменена в зависимости от ваших нужд)
CMD ["gunicorn", "--bind", "0.0.0.0:8000", "excel.wsgi:application"]