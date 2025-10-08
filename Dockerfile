# Use Node.js 18 with Debian base for better package support
FROM node:20-bookworm

# Install system dependencies for image processing and Python
RUN apt-get update && apt-get install -y --no-install-recommends \
    libvips-dev \
    libraw-bin \
    dcraw \
    ghostscript \
    imagemagick \
    libmagick++-dev \
    libreoffice \
    calibre \
    fonts-dejavu fonts-liberation locales \
    python3 \
    python3-pip \
    python3-venv \
    && rm -rf /var/lib/apt/lists/*

# Install Python dependencies in virtual environment
COPY requirements.txt ./
RUN python3 -m venv /opt/venv
ENV PATH="/opt/venv/bin:$PATH"
RUN pip install --no-cache-dir -r requirements.txt

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json ./
COPY tsconfig.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm install

# Copy source code and Python scripts
COPY src/ ./src/
COPY scripts/ ./scripts/
COPY viewers/ ./viewers/

# Build the application
RUN npm run build

# Remove dev dependencies to reduce image size
RUN npm prune --production

# Create non-root user for security with proper home directory
RUN groupadd -r appuser && useradd -r -g appuser -d /home/appuser -m appuser
RUN chown -R appuser:appuser /app

# Create necessary directories for LibreOffice and set permissions
RUN mkdir -p /home/appuser/.cache/dconf \
    && mkdir -p /home/appuser/.config/libreoffice \
    && mkdir -p /tmp/libreoffice \
    && chown -R appuser:appuser /home/appuser \
    && chmod -R 755 /home/appuser

# Set environment variables for LibreOffice
ENV HOME=/home/appuser
ENV TMPDIR=/tmp
ENV DCONF_PROFILE=/dev/null

USER appuser

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/health || exit 1

# Start the application with garbage collection enabled
CMD ["node", "--expose-gc", "dist/server.js"]
