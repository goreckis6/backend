# Backend Dockerfile
FROM node:20-alpine

WORKDIR /app

# Install system dependencies for file conversions
RUN apk add --no-cache \
    libreoffice \
    imagemagick \
    ghostscript \
    calibre \
    python3 \
    py3-pip \
    dcraw \
    exiftool \
    && pip3 install rawpy

# Copy package files first for better caching
COPY package*.json ./

# Install dependencies
RUN npm ci --only=production --no-audit --no-fund

# Copy source code
COPY . .

# Build the application
RUN npm run build

# Create non-root user
RUN addgroup -g 1001 -S nodejs
RUN adduser -S nodejs -u 1001

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