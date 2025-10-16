import { User, Conversion } from '../database/index.js';
import { Op } from 'sequelize';

export class DatabaseService {
  // User operations
  static async createUser(userData: {
    email: string;
    password: string;
    name?: string;
  }) {
    return await User.create(userData);
  }

  static async findUserByEmail(email: string) {
    return await User.findOne({ where: { email } });
  }

  static async findUserById(id: number) {
    return await User.findByPk(id);
  }

  static async updateUser(id: number, updates: { name?: string; email?: string; password?: string }) {
    const user = await User.findByPk(id);
    if (!user) return null;

    await user.update(updates);
    return user;
  }

  // Conversion operations
  static async createConversion(conversionData: {
    originalFilename: string;
    convertedFilename: string;
    originalFormat: string;
    convertedFormat: string;
    fileSize: number;
    conversionTime?: number;
    status: 'completed' | 'failed' | 'processing';
    errorMessage?: string;
    ipAddress?: string;
    userAgent?: string;
    userId?: number;
  }) {
    return await Conversion.create(conversionData);
  }

  static async getConversionStats(userId?: number) {
    const whereClause = userId ? { userId } : {};
    
    const [total, completed, failed] = await Promise.all([
      Conversion.count({ where: whereClause }),
      Conversion.count({ where: { ...whereClause, status: 'completed' } }),
      Conversion.count({ where: { ...whereClause, status: 'failed' } })
    ]);

    return {
      total,
      completed,
      failed,
      successRate: total > 0 ? Math.round((completed / total) * 100) : 0
    };
  }

  static async getPopularFormats(limit: number = 10) {
    return await Conversion.findAll({
      where: { status: 'completed' },
      attributes: [
        'originalFormat',
        'convertedFormat',
        [Conversion.sequelize!.fn('COUNT', '*'), 'count']
      ],
      group: ['originalFormat', 'convertedFormat'],
      order: [[Conversion.sequelize!.fn('COUNT', '*'), 'DESC']],
      limit,
      raw: true
    });
  }

  static async getRecentConversions(limit: number = 10) {
    return await Conversion.findAll({
      order: [['createdAt', 'DESC']],
      limit,
      include: [
        {
          model: User,
          as: 'user',
          attributes: ['id', 'email', 'name']
        }
      ]
    });
  }

  static async findConversionsByUser(userId: number, limit: number = 10) {
    return await Conversion.findAll({
      where: { userId },
      order: [['createdAt', 'DESC']],
      limit
    });
  }

  // Cleanup operations
  static async cleanupOldFailedConversions(daysOld: number = 30) {
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysOld);
    
    const deletedCount = await Conversion.destroy({
      where: {
        status: 'failed',
        createdAt: {
          [Op.lt]: cutoffDate
        }
      }
    });

    return deletedCount;
  }

  // Get user dashboard statistics
  static async getUserDashboardStats(userId: number) {
    try {
      const [
        totalConversions,
        completedConversions,
        failedConversions,
        todayConversions,
        weekConversions,
        monthConversions,
        popularFormats
      ] = await Promise.all([
        Conversion.count({ where: { userId } }),
        Conversion.count({ where: { userId, status: 'completed' } }),
        Conversion.count({ where: { userId, status: 'failed' } }),
        Conversion.count({ 
          where: { 
            userId,
            createdAt: {
              [Op.gte]: new Date(new Date().setHours(0, 0, 0, 0))
            }
          }
        }),
        Conversion.count({ 
          where: { 
            userId,
            createdAt: {
              [Op.gte]: new Date(Date.now() - 7 * 24 * 60 * 60 * 1000)
            }
          }
        }),
        Conversion.count({ 
          where: { 
            userId,
            createdAt: {
              [Op.gte]: new Date(Date.now() - 30 * 24 * 60 * 60 * 1000)
            }
          }
        }),
        Conversion.findAll({
          where: { userId, status: 'completed' },
          attributes: [
            'originalFormat',
            'convertedFormat',
            [Conversion.sequelize!.fn('COUNT', '*'), 'count']
          ],
          group: ['originalFormat', 'convertedFormat'],
          order: [[Conversion.sequelize!.fn('COUNT', '*'), 'DESC']],
          limit: 5,
          raw: true
        })
      ]);

      const avgConversionTime = await Conversion.findOne({
        where: { userId, status: 'completed' },
        attributes: [
          [Conversion.sequelize!.fn('AVG', Conversion.sequelize!.col('conversionTime')), 'avgTime']
        ],
        raw: true
      });

      return {
        success: true,
        dashboard: {
          totalConversions,
          completedConversions,
          failedConversions,
          successRate: totalConversions > 0 ? Math.round((completedConversions / totalConversions) * 100) : 0,
          todayConversions,
          weekConversions,
          monthConversions,
          avgConversionTime: Math.round((avgConversionTime as any)?.avgTime || 0),
          popularFormats
        }
      };
    } catch (error) {
      console.error('Get user dashboard stats error:', error);
      return {
        success: false,
        error: 'Failed to get user dashboard statistics'
      };
    }
  }
}
