# Backend Dockerfile
FROM node:18-bullseye-slim

# Install system dependencies
RUN apt-get update && apt-get install -y \
    libvips-dev \
    libraw-bin \
    dcraw \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy package files
COPY package*.json ./
COPY tsconfig.json ./

# Install all dependencies (including dev dependencies for build)
RUN npm ci

# Copy source code
COPY src ./src

# Build the application
RUN npm run build

# Remove dev dependencies and source code to reduce image size
RUN npm prune --production && \
    rm -rf src tsconfig.json

# Create non-root user
RUN groupadd -r nodejs && useradd -r -g nodejs nodejs
RUN chown -R nodejs:nodejs /app

USER nodejs

EXPOSE 3000

CMD ["npm", "start"]

