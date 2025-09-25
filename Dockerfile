# Use Node.js 18 with Debian base for better package support
FROM node:18-bullseye-slim

# Install system dependencies for image processing
RUN apt-get update && apt-get install -y \
    libvips-dev \
    libraw-bin \
    dcraw \
    ghostscript \
    imagemagick \
    libmagick++-dev \
    && rm -rf /var/lib/apt/lists/*

# Set working directory
WORKDIR /app

# Copy package files
COPY package.json ./
COPY tsconfig.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm install

# Copy source code
COPY src/ ./src/

# Build the application
RUN npm run build

# Remove dev dependencies to reduce image size
RUN npm prune --production

# Create non-root user for security
RUN groupadd -r appuser && useradd -r -g appuser appuser
RUN chown -R appuser:appuser /app
USER appuser

# Expose port
EXPOSE 3000

# Health check
HEALTHCHECK --interval=30s --timeout=3s --start-period=5s --retries=3 \
  CMD curl -f http://localhost:3000/health || exit 1

# Start the application with garbage collection enabled
CMD ["node", "--expose-gc", "dist/server.js"]
