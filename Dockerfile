FROM python:3.11-slim

# Install system deps (fonts & basic utilities)
RUN apt-get update && apt-get install -y --no-install-recommends \
    fonts-noto-cjk \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY requirements.txt /app/
RUN pip install --no-cache-dir -r requirements.txt

# Copy app and optional fonts folder (if you bundle fonts)
COPY . /app

EXPOSE 8501

ENV STREAMLIT_SERVER_HEADLESS=true
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
