# Backend Dockerfile
FROM node:18-alpine

# Install system dependencies
RUN apk add --no-cache \
    vips-dev \
    libraw \
    dcraw \
    build-base \
    python3 \
    make \
    g++

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
RUN addgroup -g 1001 -S nodejs && \
    adduser -S nodejs -u 1001

USER nodejs

EXPOSE 3000

CMD ["npm", "start"]

