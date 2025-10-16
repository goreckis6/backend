import express from 'express';
import { body, validationResult } from 'express-validator';
import { AuthService } from '../services/authService.js';
import { authenticateToken } from '../middleware/auth.js';

const router = express.Router();

// Validation middleware
const validateRegister = [
  body('email')
    .isEmail()
    .normalizeEmail()
    .withMessage('Please provide a valid email address'),
  body('password')
    .isLength({ min: 6 })
    .withMessage('Password must be at least 6 characters long'),
  body('name')
    .optional()
    .trim()
    .isLength({ min: 2, max: 50 })
    .withMessage('Name must be between 2 and 50 characters')
];

const validateLogin = [
  body('email')
    .isEmail()
    .normalizeEmail()
    .withMessage('Please provide a valid email address'),
  body('password')
    .notEmpty()
    .withMessage('Password is required')
];

const validateProfileUpdate = [
  body('name')
    .optional()
    .trim()
    .isLength({ min: 2, max: 50 })
    .withMessage('Name must be between 2 and 50 characters'),
  body('email')
    .optional()
    .isEmail()
    .normalizeEmail()
    .withMessage('Please provide a valid email address')
];

const validatePasswordChange = [
  body('currentPassword')
    .notEmpty()
    .withMessage('Current password is required'),
  body('newPassword')
    .isLength({ min: 6 })
    .withMessage('New password must be at least 6 characters long')
];

// Helper function to handle validation errors
const handleValidationErrors = (req: express.Request, res: express.Response, next: express.NextFunction) => {
  const errors = validationResult(req);
  if (!errors.isEmpty()) {
    return res.status(400).json({
      success: false,
      error: 'Validation failed',
      details: errors.array()
    });
  }
  next();
};

// POST /api/auth/register
router.post('/register', validateRegister, handleValidationErrors, async (req, res) => {
  try {
    const { email, password, name } = req.body;
    
    const result = await AuthService.register({ email, password, name });
    
    if (result.success) {
      res.status(201).json({
        success: true,
        message: 'User registered successfully',
        user: result.user,
        token: result.token
      });
    } else {
      res.status(400).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Register route error:', error);
    res.status(500).json({
      success: false,
      error: 'Registration failed'
    });
  }
});

// POST /api/auth/login
router.post('/login', validateLogin, handleValidationErrors, async (req, res) => {
  try {
    const { email, password } = req.body;
    
    const result = await AuthService.login({ email, password });
    
    if (result.success) {
      res.json({
        success: true,
        message: 'Login successful',
        user: result.user,
        token: result.token
      });
    } else {
      res.status(401).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Login route error:', error);
    res.status(500).json({
      success: false,
      error: 'Login failed'
    });
  }
});

// GET /api/auth/me
router.get('/me', authenticateToken, async (req, res) => {
  try {
    res.json({
      success: true,
      user: req.user
    });
  } catch (error) {
    console.error('Get user profile error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get user profile'
    });
  }
});

// PUT /api/auth/profile
router.put('/profile', authenticateToken, validateProfileUpdate, handleValidationErrors, async (req, res) => {
  try {
    const { name, email } = req.body;
    
    const result = await AuthService.updateProfile(req.user!.id, { name, email });
    
    if (result.success) {
      res.json({
        success: true,
        message: 'Profile updated successfully',
        user: result.user
      });
    } else {
      res.status(400).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Update profile error:', error);
    res.status(500).json({
      success: false,
      error: 'Profile update failed'
    });
  }
});

// PUT /api/auth/change-password
router.put('/change-password', authenticateToken, validatePasswordChange, handleValidationErrors, async (req, res) => {
  try {
    const { currentPassword, newPassword } = req.body;
    
    const result = await AuthService.changePassword(req.user!.id, currentPassword, newPassword);
    
    if (result.success) {
      res.json({
        success: true,
        message: 'Password changed successfully',
        user: result.user
      });
    } else {
      res.status(400).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Change password error:', error);
    res.status(500).json({
      success: false,
      error: 'Password change failed'
    });
  }
});

// GET /api/auth/stats
router.get('/stats', authenticateToken, async (req, res) => {
  try {
    const result = await AuthService.getUserStats(req.user!.id);
    
    if (result.success) {
      res.json({
        success: true,
        stats: result.stats,
        recentConversions: result.recentConversions
      });
    } else {
      res.status(500).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Get user stats error:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get user statistics'
    });
  }
});

// POST /api/auth/verify-token
router.post('/verify-token', async (req, res) => {
  try {
    const { token } = req.body;
    
    if (!token) {
      return res.status(400).json({
        success: false,
        error: 'Token is required'
      });
    }
    
    const result = await AuthService.getUserByToken(token);
    
    if (result.success) {
      res.json({
        success: true,
        user: result.user
      });
    } else {
      res.status(401).json({
        success: false,
        error: result.error
      });
    }
  } catch (error) {
    console.error('Verify token error:', error);
    res.status(500).json({
      success: false,
      error: 'Token verification failed'
    });
  }
});

export default router;
