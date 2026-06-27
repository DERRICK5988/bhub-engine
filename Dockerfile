FROM python:3.11-slim

# LibreOffice (xlsx recalc + PDF) and fonts
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-calc fonts-dejavu fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY app/ ./

EXPOSE 8000
CMD ["uvicorn", "main:app", "--host", "0.0.0.0", "--port", "8000"]
