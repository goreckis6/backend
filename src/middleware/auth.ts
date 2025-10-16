import { Request, Response, NextFunction } from 'express';
import { AuthService } from '../services/authService.js';

// Extend Request interface to include user
declare global {
  namespace Express {
    interface Request {
      user?: {
        id: number;
        email: string;
        name?: string;
        isActive: boolean;
      };
    }
  }
}

// Authentication middleware
export const authenticateToken = async (req: Request, res: Response, next: NextFunction) => {
  try {
    const authHeader = req.headers.authorization;
    const token = authHeader && authHeader.split(' ')[1]; // Bearer TOKEN

    if (!token) {
      return res.status(401).json({ 
        success: false,
        error: 'Access token required' 
      });
    }

    const result = await AuthService.getUserByToken(token);
    
    if (!result.success || !result.user) {
      return res.status(401).json({ 
        success: false,
        error: result.error || 'Invalid token' 
      });
    }

    req.user = result.user;
    next();
  } catch (error) {
    console.error('Authentication middleware error:', error);
    return res.status(500).json({ 
      success: false,
      error: 'Authentication failed' 
    });
  }
};

// Optional authentication middleware (doesn't fail if no token)
export const optionalAuth = async (req: Request, res: Response, next: NextFunction) => {
  try {
    const authHeader = req.headers.authorization;
    const token = authHeader && authHeader.split(' ')[1];

    if (token) {
      const result = await AuthService.getUserByToken(token);
      if (result.success && result.user) {
        req.user = result.user;
      }
    }

    next();
  } catch (error) {
    console.error('Optional auth middleware error:', error);
    // Continue without authentication
    next();
  }
};

// Admin middleware (placeholder for future admin features)
export const requireAdmin = (req: Request, res: Response, next: NextFunction) => {
  // For now, just check if user is authenticated
  // In the future, you can add admin role checking
  if (!req.user) {
    return res.status(401).json({ 
      success: false,
      error: 'Admin access required' 
    });
  }

  // TODO: Add admin role check when you implement roles
  next();
};
