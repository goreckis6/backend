const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const rateLimit = require('express-rate-limit');
const multer = require('multer');

const app = express();
const PORT = process.env.PORT || 10000;

// CORS configuration
app.use(helmet());
app.use(cors({
  origin: process.env.FRONTEND_URL || 'https://morphy-1-ulvv.onrender.com',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: false
}));

// Handle preflight requests
app.options('*', cors({
  origin: process.env.FRONTEND_URL || 'https://morphy-1-ulvv.onrender.com',
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization'],
  credentials: false
}));

const limiter = rateLimit({
  windowMs: 15 * 60 * 1000,
  max: 60,
  message: 'Too many requests from this IP, please try again later.'
});
app.use('/api/', limiter);

app.use(express.json({ limit: '10mb' }));
app.use(express.urlencoded({ extended: true, limit: '10mb' }));

const upload = multer({
  storage: multer.memoryStorage(),
  limits: {
    fileSize: 100 * 1024 * 1024,
    files: 1
  }
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Basic convert endpoint - just returns success for now
app.post('/api/convert', upload.single('file'), async (req, res) => {
  try {
    console.log('=== CONVERSION REQUEST START ===');
    console.log('File received:', req.file ? req.file.originalname : 'No file');
    console.log('Request options:', req.body);
    
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }
    
    // For now, just return a simple response to test CORS
    res.json({
      message: 'Conversion endpoint working',
      filename: req.file.originalname,
      size: req.file.size,
      options: req.body
    });
    
    console.log('=== CONVERSION REQUEST END (TEST) ===');
  } catch (error) {
    console.error('Conversion error:', error);
    res.status(500).json({ error: 'Conversion failed' });
  }
});

// Error handler
app.use((err, req, res, next) => {
  console.error('Unhandled error:', err);
  res.status(500).json({ error: 'Internal server error' });
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Morphy backend running on port ${PORT}`);
});

module.exports = app;
