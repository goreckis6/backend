FROM node:20-bookworm

# Install system dependencies for file conversion
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
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package*.json ./

# Install production dependencies only
RUN npm ci --only=production

# Copy application code
COPY server.js ./

# Expose port
EXPOSE 10000

# Start the application
CMD ["node", "server.js"]