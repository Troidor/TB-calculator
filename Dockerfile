# Use a slim Python image
FROM python:3.11-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

# Install LibreOffice and required libs
RUN apt-get update && apt-get install -y --no-install-recommends \
    libreoffice-core libreoffice-common libreoffice-calc libreoffice-writer \
    fonts-dejavu-core fontconfig \
    wget unzip \
  && apt-get clean \
  && rm -rf /var/lib/apt/lists/*

# Create app dir
WORKDIR /app

# Copy requirements first (for better caching)
COPY requirements.txt /app/
RUN pip install --upgrade pip
RUN pip install --no-cache-dir -r requirements.txt

# Copy app files
COPY . /app

# Ensure excel folder and workbook exist
# (you should have excel/TB_v18.xlsx in your repo)
RUN ls -la /app/excel || true

# Expose port
EXPOSE 8080

# Fly provides PORT env var at runtime; use gunicorn
CMD ["gunicorn", "app:app", "-b", "0.0.0.0:8080", "--workers", "1", "--threads", "4"]
