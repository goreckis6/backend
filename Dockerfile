# Backend Dockerfile
FROM node:18-bullseye

RUN apt-get update && apt-get install -y \
    libvips-dev \
    libraw-bin \
    dcraw \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

COPY package*.json ./
COPY tsconfig.json ./
COPY src ./src

RUN npm install
RUN npm run build

CMD ["npm", "start"]

