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
      console.log('🔍 Authenticated user - skipping conversion limits');
      return next();
    }

    // Get client IP address
    const clientIP = getClientIP(req);
    
    console.log('🔍 Checking conversion limits for anonymous user:', {
      clientIP,
      endpoint: req.path,
      method: req.method
    });
    
    // Check if IP can perform conversion
    const canConvert = await AnonymousConversionService.canConvert(clientIP);
    
    console.log('🔍 Conversion limit check result:', {
      clientIP,
      canConvert
    });
    
    if (!canConvert) {
      const status = await AnonymousConversionService.getConversionStatus(clientIP);
      
      console.log('❌ Conversion limit reached for IP:', {
        clientIP,
        usedConversions: status.usedConversions,
        limit: status.limit,
        remainingConversions: status.remainingConversions
      });
      
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
    
    console.log('✅ Conversion limit check passed for IP:', clientIP);
    next();
  } catch (error) {
    console.error('❌ Error checking conversion limits:', error);
    // In case of error, allow the request to proceed (fail open)
    next();
  }
};

/**
 * Middleware to record conversion after successful conversion
 */
export const recordConversion = async (req: Request, res: Response, next: NextFunction) => {
  try {
    // Only record for anonymous users (no Bearer token)
    const authHeader = req.headers.authorization;
    if (authHeader && authHeader.startsWith('Bearer ')) {
      // Skip recording for authenticated users - they have unlimited conversions
      return next();
    }

    // Get client IP address
    const clientIP = (req as any).clientIP || getClientIP(req);
    const userAgent = req.headers['user-agent'];

    console.log('🔍 Recording conversion for anonymous user:', {
      clientIP,
      userAgent: userAgent ? userAgent.substring(0, 50) + '...' : 'unknown'
    });

    // Record the conversion for anonymous user
    await AnonymousConversionService.recordConversion(clientIP, userAgent);
    
    console.log('✅ Conversion recorded successfully for IP:', clientIP);
    next();
  } catch (error) {
    console.error('❌ Error recording conversion:', error);
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
    
    console.log('🔍 Getting conversion status for IP:', {
      clientIP,
      headers: {
        'x-forwarded-for': req.headers['x-forwarded-for'],
        'x-real-ip': req.headers['x-real-ip'],
        'cf-connecting-ip': req.headers['cf-connecting-ip'],
        'user-agent': req.headers['user-agent']?.substring(0, 50) + '...'
      },
      connection: {
        remoteAddress: req.connection.remoteAddress,
        socketRemoteAddress: req.socket.remoteAddress
      }
    });
    
    const status = await AnonymousConversionService.getConversionStatus(clientIP);
    
    console.log('🔍 Conversion status result:', {
      clientIP,
      status
    });
    
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
