# Word MCP Server Dockerfile
# syntax=docker/dockerfile:1

# Use official Python runtime
FROM python:3.11-slim

# Set working directory
WORKDIR /app

# Install build dependencies and LibreOffice for PDF conversion
RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        build-essential \
        libreoffice-writer \
    && rm -rf /var/lib/apt/lists/*

# Copy project files
COPY . /app

# Install Python dependencies
RUN pip install --no-cache-dir .

# Expose port for HTTP transport
EXPOSE 8000

# Default command
ENTRYPOINT ["word_mcp_server"]
