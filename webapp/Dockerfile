FROM python:3.11-slim
ENV DEBIAN_FRONTEND=noninteractive
WORKDIR /app

# Install system packages required for PPTX -> images conversion
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
       libreoffice \
       imagemagick \
       ghostscript \
       fonts-dejavu-core \
    && rm -rf /var/lib/apt/lists/*

# Install Python requirements
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy repository and expose port
COPY .. /app
EXPOSE 8000

CMD ["uvicorn", "webapp.app:app", "--host", "0.0.0.0", "--port", "8000"]
