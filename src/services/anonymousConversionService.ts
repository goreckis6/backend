import { AnonymousConversion } from '../database/models/AnonymousConversion.js';
import { Op } from 'sequelize';

export interface AnonymousConversionStatus {
  canConvert: boolean;
  remainingConversions: number;
  usedConversions: number;
  limit: number;
  message: string;
}

export class AnonymousConversionService {
  private static readonly FREE_CONVERSION_LIMIT = 5;
  private static readonly CLEANUP_DAYS = 30; // Clean up old records after 30 days

  /**
   * Get conversion status for an IP address
   */
  static async getConversionStatus(ipAddress: string): Promise<AnonymousConversionStatus> {
    try {
      const record = await AnonymousConversion.findOne({
        where: { ipAddress }
      });

      const usedConversions = record?.conversionCount || 0;
      const remainingConversions = Math.max(0, this.FREE_CONVERSION_LIMIT - usedConversions);
      const canConvert = remainingConversions > 0;

      let message = '';
      if (canConvert) {
        message = `${remainingConversions} free conversions remaining`;
      } else {
        message = `You've used all ${this.FREE_CONVERSION_LIMIT} free conversions. Register for unlimited conversions!`;
      }

      return {
        canConvert,
        remainingConversions,
        usedConversions,
        limit: this.FREE_CONVERSION_LIMIT,
        message
      };
    } catch (error) {
      console.error('Error getting conversion status:', error);
      // In case of error, allow conversion (fail open)
      return {
        canConvert: true,
        remainingConversions: this.FREE_CONVERSION_LIMIT,
        usedConversions: 0,
        limit: this.FREE_CONVERSION_LIMIT,
        message: `${this.FREE_CONVERSION_LIMIT} free conversions remaining`
      };
    }
  }

  /**
   * Record a conversion for an IP address
   */
  static async recordConversion(ipAddress: string, userAgent?: string): Promise<boolean> {
    try {
      const [record, created] = await AnonymousConversion.findOrCreate({
        where: { ipAddress },
        defaults: {
          ipAddress,
          userAgent,
          conversionCount: 1,
          lastConversionAt: new Date()
        }
      });

      if (!created) {
        // Update existing record
        await record.update({
          conversionCount: record.conversionCount + 1,
          lastConversionAt: new Date(),
          userAgent: userAgent || record.userAgent // Update user agent if provided
        });
      }

      console.log(`Recorded conversion for IP ${ipAddress}. Total: ${record.conversionCount + (created ? 0 : 1)}`);
      return true;
    } catch (error) {
      console.error('Error recording conversion:', error);
      return false;
    }
  }

  /**
   * Check if an IP can perform a conversion
   */
  static async canConvert(ipAddress: string): Promise<boolean> {
    const status = await this.getConversionStatus(ipAddress);
    return status.canConvert;
  }

  /**
   * Get conversion statistics for admin purposes
   */
  static async getConversionStats(): Promise<{
    totalAnonymousUsers: number;
    totalConversions: number;
    activeUsers: number;
    topIPs: Array<{ ipAddress: string; conversionCount: number; lastConversionAt: Date }>;
  }> {
    try {
      const totalAnonymousUsers = await AnonymousConversion.count();
      
      const totalConversions = await AnonymousConversion.sum('conversionCount') || 0;
      
      // Active users in the last 7 days
      const sevenDaysAgo = new Date();
      sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);
      
      const activeUsers = await AnonymousConversion.count({
        where: {
          lastConversionAt: {
            [Op.gte]: sevenDaysAgo
          }
        }
      });

      // Top IPs by conversion count
      const topIPs = await AnonymousConversion.findAll({
        attributes: ['ipAddress', 'conversionCount', 'lastConversionAt'],
        order: [['conversionCount', 'DESC']],
        limit: 10
      });

      return {
        totalAnonymousUsers,
        totalConversions,
        activeUsers,
        topIPs: topIPs.map(record => ({
          ipAddress: record.ipAddress,
          conversionCount: record.conversionCount,
          lastConversionAt: record.lastConversionAt
        }))
      };
    } catch (error) {
      console.error('Error getting conversion stats:', error);
      return {
        totalAnonymousUsers: 0,
        totalConversions: 0,
        activeUsers: 0,
        topIPs: []
      };
    }
  }

  /**
   * Clean up old conversion records
   */
  static async cleanupOldRecords(): Promise<number> {
    try {
      const cutoffDate = new Date();
      cutoffDate.setDate(cutoffDate.getDate() - this.CLEANUP_DAYS);

      const deletedCount = await AnonymousConversion.destroy({
        where: {
          lastConversionAt: {
            [Op.lt]: cutoffDate
          }
        }
      });

      console.log(`Cleaned up ${deletedCount} old anonymous conversion records`);
      return deletedCount;
    } catch (error) {
      console.error('Error cleaning up old records:', error);
      return 0;
    }
  }

  /**
   * Reset conversion count for an IP (for testing or manual intervention)
   */
  static async resetConversionsForIP(ipAddress: string): Promise<boolean> {
    try {
      const record = await AnonymousConversion.findOne({
        where: { ipAddress }
      });

      if (record) {
        await record.update({
          conversionCount: 0,
          lastConversionAt: new Date()
        });
        console.log(`Reset conversion count for IP ${ipAddress}`);
        return true;
      }

      return false;
    } catch (error) {
      console.error('Error resetting conversions for IP:', error);
      return false;
    }
  }
}
