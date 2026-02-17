FROM python:3.11-slim

# Install LibreOffice and dependencies
RUN apt-get update && apt-get install -y \
    libreoffice \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy requirements and install
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy app files
COPY . .

# Set environment variables
ENV PORT=8501

# Expose port
EXPOSE 8501

# Run streamlit
ENTRYPOINT ["sh", "-c", "streamlit run app.py --server.port=$PORT --server.address=0.0.0.0"]
