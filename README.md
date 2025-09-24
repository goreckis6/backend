# Morphy Backend

RAW-aware image conversion API for Morphy.

## Features

- RAW image format support (CR2, CR3, NEF, ARW, DNG, etc.)
- Image format conversion (WebP, PNG, JPEG, ICO)
- Image resizing and quality adjustment
- Rate limiting and security headers
- Docker support

## Environment Variables

Set these environment variables in your deployment:

```bash
PORT=3000
NODE_ENV=production
FRONTEND_URL=https://morphy-1-ulvv.onrender.com
```

## Deployment on Render.com

1. Connect your GitHub repository to Render
2. Create a new Web Service
3. Set the following:
   - **Build Command**: `npm install && npm run build`
   - **Start Command**: `npm start`
   - **Environment**: `Node`
   - **Node Version**: `18`
   - **Dockerfile Path**: `./Dockerfile` (if using Docker deployment)

### Docker Deployment (Recommended)
- Render will automatically detect and use the Dockerfile
- The Dockerfile includes all necessary system dependencies for image processing
- Optimized for production with security best practices

## API Endpoints

- `GET /health` - Health check with system info
- `GET /api/status` - Simple status check
- `POST /api/convert` - Convert image files

## Development

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
npm start
```
