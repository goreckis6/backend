import bcrypt from 'bcryptjs';
import jwt from 'jsonwebtoken';
import { User } from '../database/index.js';
import { DatabaseService } from './databaseService.js';

export interface LoginCredentials {
  email: string;
  password: string;
}

export interface RegisterData {
  email: string;
  password: string;
  name?: string;
}

export interface AuthResult {
  success: boolean;
  user?: {
    id: number;
    email: string;
    name?: string;
    isActive: boolean;
  };
  token?: string;
  error?: string;
}

export interface JwtPayload {
  userId: number;
  email: string;
  iat?: number;
  exp?: number;
}

export class AuthService {
  private static readonly JWT_SECRET = process.env.JWT_SECRET || 'fallback-secret-key';
  private static readonly JWT_EXPIRES_IN = '7d';
  private static readonly SALT_ROUNDS = 12;

  // Hash password
  static async hashPassword(password: string): Promise<string> {
    try {
      console.log('🔍 Hashing password...');
      const hashed = await bcrypt.hash(password, this.SALT_ROUNDS);
      console.log('✅ Password hashed successfully');
      return hashed;
    } catch (error) {
      console.error('❌ Password hashing failed:', error);
      throw error;
    }
  }

  // Verify password
  static async verifyPassword(password: string, hashedPassword: string): Promise<boolean> {
    return await bcrypt.compare(password, hashedPassword);
  }

  // Generate JWT token
  static generateToken(payload: JwtPayload): string {
    return jwt.sign(payload, this.JWT_SECRET, { expiresIn: this.JWT_EXPIRES_IN });
  }

  // Verify JWT token
  static verifyToken(token: string): JwtPayload | null {
    try {
      return jwt.verify(token, this.JWT_SECRET) as JwtPayload;
    } catch (error) {
      return null;
    }
  }

  // Register new user
  static async register(data: RegisterData): Promise<AuthResult> {
    try {
      console.log('🔍 Registering user:', { email: data.email, name: data.name });
      
      // Check if user already exists
      const existingUser = await DatabaseService.findUserByEmail(data.email);
      if (existingUser) {
        console.log('❌ User already exists:', data.email);
        return {
          success: false,
          error: 'User with this email already exists'
        };
      }

      // Hash password
      const hashedPassword = await this.hashPassword(data.password);
      console.log('✅ Password hashed successfully');

      // Create user
      const user = await DatabaseService.createUser({
        email: data.email,
        password: hashedPassword,
        name: data.name
      });
      console.log('✅ User created successfully:', user.id);

      // Generate token
      const token = this.generateToken({
        userId: user.id,
        email: user.email
      });

      return {
        success: true,
        user: {
          id: user.id,
          email: user.email,
          name: user.name,
          isActive: user.isActive
        },
        token
      };
    } catch (error) {
      console.error('❌ Registration error:', error);
      console.error('❌ Error details:', error instanceof Error ? error.message : 'Unknown error');
      console.error('❌ Stack trace:', error instanceof Error ? error.stack : 'No stack trace');
      return {
        success: false,
        error: `Registration failed: ${error instanceof Error ? error.message : 'Unknown error'}`
      };
    }
  }

  // Login user
  static async login(credentials: LoginCredentials): Promise<AuthResult> {
    try {
      // Find user by email
      const user = await DatabaseService.findUserByEmail(credentials.email);
      if (!user) {
        return {
          success: false,
          error: 'Invalid email or password'
        };
      }

      // Check if user is active
      if (!user.isActive) {
        return {
          success: false,
          error: 'Account is deactivated. Please contact support.'
        };
      }

      // Verify password
      const isValidPassword = await this.verifyPassword(credentials.password, user.password);
      if (!isValidPassword) {
        return {
          success: false,
          error: 'Invalid email or password'
        };
      }

      // Generate token
      const token = this.generateToken({
        userId: user.id,
        email: user.email
      });

      return {
        success: true,
        user: {
          id: user.id,
          email: user.email,
          name: user.name,
          isActive: user.isActive
        },
        token
      };
    } catch (error) {
      console.error('Login error:', error);
      return {
        success: false,
        error: 'Login failed. Please try again.'
      };
    }
  }

  // Get user by token
  static async getUserByToken(token: string): Promise<AuthResult> {
    try {
      const payload = this.verifyToken(token);
      if (!payload) {
        return {
          success: false,
          error: 'Invalid or expired token'
        };
      }

      const user = await DatabaseService.findUserById(payload.userId);
      if (!user || !user.isActive) {
        return {
          success: false,
          error: 'User not found or inactive'
        };
      }

      return {
        success: true,
        user: {
          id: user.id,
          email: user.email,
          name: user.name,
          isActive: user.isActive
        }
      };
    } catch (error) {
      console.error('Token verification error:', error);
      return {
        success: false,
        error: 'Token verification failed'
      };
    }
  }

  // Update user profile
  static async updateProfile(userId: number, updates: { name?: string; email?: string }): Promise<AuthResult> {
    try {
      const user = await DatabaseService.updateUser(userId, updates);
      if (!user) {
        return {
          success: false,
          error: 'User not found'
        };
      }

      return {
        success: true,
        user: {
          id: user.id,
          email: user.email,
          name: user.name,
          isActive: user.isActive
        }
      };
    } catch (error) {
      console.error('Profile update error:', error);
      return {
        success: false,
        error: 'Profile update failed'
      };
    }
  }

  // Change password
  static async changePassword(userId: number, currentPassword: string, newPassword: string): Promise<AuthResult> {
    try {
      const user = await DatabaseService.findUserById(userId);
      if (!user) {
        return {
          success: false,
          error: 'User not found'
        };
      }

      // Verify current password
      const isValidPassword = await this.verifyPassword(currentPassword, user.password);
      if (!isValidPassword) {
        return {
          success: false,
          error: 'Current password is incorrect'
        };
      }

      // Hash new password
      const hashedNewPassword = await this.hashPassword(newPassword);

      // Update password
      await DatabaseService.updateUser(userId, { password: hashedNewPassword });

      return {
        success: true,
        user: {
          id: user.id,
          email: user.email,
          name: user.name,
          isActive: user.isActive
        }
      };
    } catch (error) {
      console.error('Password change error:', error);
      return {
        success: false,
        error: 'Password change failed'
      };
    }
  }

  // Get user statistics
  static async getUserStats(userId: number) {
    try {
      const stats = await DatabaseService.getConversionStats(userId);
      const conversions = await DatabaseService.findConversionsByUser(userId, 10);
      
      return {
        success: true,
        stats,
        recentConversions: conversions
      };
    } catch (error) {
      console.error('Get user stats error:', error);
      return {
        success: false,
        error: 'Failed to get user statistics'
      };
    }
  }
}
