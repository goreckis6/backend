import { Request, Response, NextFunction } from 'express';
import { AnonymousConversionService } from '../services/anonymousConversionService.js';

/**
 * Extract real IP address from request
 * Handles proxies, load balancers, and CDNs
 */
export const getClientIP = (req: Request): string => {
  // Check various headers that might contain the real IP
  const headers = [
    'x-forwarded-for',
    'x-real-ip',
    'x-client-ip',
    'cf-connecting-ip', // Cloudflare
    'x-cluster-client-ip',
    'x-forwarded',
    'forwarded-for',
    'forwarded'
  ];

  for (const header of headers) {
    const value = req.headers[header] as string;
    if (value) {
      // x-forwarded-for can contain multiple IPs, take the first one
      const ip = value.split(',')[0].trim();
      if (ip && ip !== 'unknown') {
        return ip;
      }
    }
  }

  // Fallback to connection remote address
  return req.connection.remoteAddress || req.socket.remoteAddress || '127.0.0.1';
};

/**
 * Middleware to check conversion limits for anonymous users
 */
export const checkConversionLimits = async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Skip limit checking for authenticated users
    const authHeader = req.headers.authorization;
    if (authHeader && authHeader.startsWith('Bearer ')) {
      // User is authenticated, skip limits
      return next();
    }

    // Get client IP address
    const clientIP = getClientIP(req);
    
    // Check if IP can perform conversion
    const canConvert = await AnonymousConversionService.canConvert(clientIP);
    
    if (!canConvert) {
      const status = await AnonymousConversionService.getConversionStatus(clientIP);
      
      return res.status(429).json({
        success: false,
        error: 'Conversion limit reached',
        message: status.message,
        details: {
          remainingConversions: status.remainingConversions,
          usedConversions: status.usedConversions,
          limit: status.limit,
          suggestion: 'Register for unlimited conversions'
        }
      });
    }

    // Add IP to request for later use
    (req as any).clientIP = clientIP;
    
    next();
  } catch (error) {
    console.error('Error checking conversion limits:', error);
    // In case of error, allow the request to proceed (fail open)
    next();
  }
};

/**
 * Middleware to record conversion after successful conversion
 */
export const recordConversion = async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Skip recording for authenticated users
    const authHeader = req.headers.authorization;
    if (authHeader && authHeader.startsWith('Bearer ')) {
      return next();
    }

    // Get client IP address
    const clientIP = (req as any).clientIP || getClientIP(req);
    const userAgent = req.headers['user-agent'];

    // Record the conversion
    await AnonymousConversionService.recordConversion(clientIP, userAgent);
    
    next();
  } catch (error) {
    console.error('Error recording conversion:', error);
    // Don't fail the request if recording fails
    next();
  }
};

/**
 * Get conversion status endpoint handler
 */
export const getConversionStatus = async (req: Request, res: Response) => {
  try {
    const clientIP = getClientIP(req);
    const status = await AnonymousConversionService.getConversionStatus(clientIP);
    
    res.json({
      success: true,
      data: status
    });
  } catch (error) {
    console.error('Error getting conversion status:', error);
    res.status(500).json({
      success: false,
      error: 'Failed to get conversion status'
    });
  }
};
