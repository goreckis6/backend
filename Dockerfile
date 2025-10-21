# Backend Dockerfile
FROM node:20-slim

WORKDIR /app

# Install system dependencies step by step
RUN apt-get update

# Install basic packages first
RUN apt-get install -y \
    libreoffice \
    imagemagick \
    ghostscript \
    python3 \
    python3-pip

# Install Sharp dependencies
RUN apt-get install -y \
    libvips-dev \
    libglib2.0-dev \
    libcairo2-dev \
    libpango1.0-dev \
    libgdk-pixbuf2.0-dev \
    libffi-dev

# Install image format libraries
RUN apt-get install -y \
    libjpeg-dev \
    libpng-dev \
    libtiff-dev \
    libwebp-dev \
    libexif-dev \
    libraw-dev \
    build-essential \
    libffi-dev

# Install Python packages for RAW processing and CSV conversion
RUN pip3 install --no-cache-dir rawpy Pillow pandas python-docx openpyxl xlsxwriter

# Create Python virtual environment for consistency
RUN python3 -m venv /opt/venv
RUN /opt/venv/bin/pip install --no-cache-dir rawpy Pillow pandas python-docx openpyxl xlsxwriter

# Clean up
RUN apt-get clean && rm -rf /var/lib/apt/lists/*

# Copy package files first for better caching
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production --no-audit --no-fund

# Copy source code
COPY . .

# Build the application
RUN npm run build

# Create non-root user
RUN groupadd -g 1001 nodejs
RUN useradd -r -u 1001 -g nodejs nodejs

# Change ownership of the app directory
RUN chown -R nodejs:nodejs /app
USER nodejs

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD node -e "require('http').get('http://localhost:3000/health', (res) => { process.exit(res.statusCode === 200 ? 0 : 1) })"

# Start the application
CMD ["node", "dist/server.js"]