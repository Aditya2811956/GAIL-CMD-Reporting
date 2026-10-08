FROM python:3.11

# Install LibreOffice, fontconfig, and Carlito (metrically identical to Calibri)
RUN apt-get update && \
    apt-get install -y --no-install-recommends libreoffice-writer fontconfig fonts-crosextra-carlito fonts-crosextra-caladea && \
    apt-get clean && rm -rf /var/lib/apt/lists/*

# Install Times New Roman + fontconfig alias (Calibri → Carlito)
COPY fonts/ /usr/share/fonts/truetype/msfonts/
RUN cp /usr/share/fonts/truetype/msfonts/local.conf /etc/fonts/conf.d/99-ms-fonts.conf && \
    fc-cache -f

# Python deps
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

WORKDIR /app

# Copy app source
COPY . /app

EXPOSE 8001

# Run with --network host so the container can reach ChromaDB at 127.0.0.1:8001
CMD ["uvicorn", "app:app", "--host", "0.0.0.0", "--port", "8001"]
