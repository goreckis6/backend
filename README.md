# Morphy Backend API

A Docker-based backend service for file conversion, specifically optimized for RAW image processing.

## Features

- **RAW Image Support**: Handles CR2, DNG, NEF, ARW, and other RAW formats
- **Multiple Output Formats**: WebP, PNG, JPEG, ICO
- **Batch Processing**: Convert multiple files at once
- **High Performance**: Uses Sharp for fast image processing
- **Docker Ready**: Optimized container with all dependencies

## API Endpoints

### Health Check
```
GET /health
```

### Single File Conversion
```
POST /api/convert
Content-Type: multipart/form-data

Parameters:
- file: The image file to convert
- quality: 'high' | 'medium' | 'low' (default: 'high')
- lossless: 'true' | 'false' (default: 'false')
- format: 'webp' | 'png' | 'jpeg' | 'ico' (default: 'webp')
- width: Optional width for resizing
- height: Optional height for resizing
- iconSize: Size for ICO conversion (default: '16')
```

### Batch Conversion
```
POST /api/convert/batch
Content-Type: multipart/form-data

Parameters:
- files: Array of image files (max 10)
- quality: 'high' | 'medium' | 'low' (default: 'high')
- lossless: 'true' | 'false' (default: 'false')
- format: 'webp' | 'png' | 'jpeg' (default: 'webp')
```

## Development

```bash
# Install dependencies
npm install

# Run in development mode
npm run dev

# Build for production
npm run build

# Start production server
npm start
```

## Docker

```bash
# Build image
docker build -t morphy-backend .

# Run container
docker run -p 3000:3000 morphy-backend
```

## Environment Variables

- `PORT`: Server port (default: 3000)
- `FRONTEND_URL`: Frontend URL for CORS (default: http://localhost:5173)
- `NODE_ENV`: Environment (development/production)
