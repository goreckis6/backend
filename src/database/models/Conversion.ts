import { DataTypes, Model, Optional } from 'sequelize';
import sequelize from '../connection.js';

interface ConversionAttributes {
  id: number;
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
  createdAt: Date;
  updatedAt: Date;
}

interface ConversionCreationAttributes extends Optional<ConversionAttributes, 'id' | 'conversionTime' | 'errorMessage' | 'ipAddress' | 'userAgent' | 'userId' | 'createdAt' | 'updatedAt'> {}

export class Conversion extends Model<ConversionAttributes, ConversionCreationAttributes> implements ConversionAttributes {
  declare id: number;
  declare originalFilename: string;
  declare convertedFilename: string;
  declare originalFormat: string;
  declare convertedFormat: string;
  declare fileSize: number;
  declare conversionTime?: number;
  declare status: 'completed' | 'failed' | 'processing';
  declare errorMessage?: string;
  declare ipAddress?: string;
  declare userAgent?: string;
  declare userId?: number;
  declare readonly createdAt: Date;
  declare readonly updatedAt: Date;
}

Conversion.init(
  {
    id: {
      type: DataTypes.INTEGER,
      autoIncrement: true,
      primaryKey: true,
    },
    originalFilename: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    convertedFilename: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    originalFormat: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    convertedFormat: {
      type: DataTypes.STRING,
      allowNull: false,
    },
    fileSize: {
      type: DataTypes.INTEGER,
      allowNull: false,
    },
    conversionTime: {
      type: DataTypes.FLOAT,
      allowNull: true,
    },
    status: {
      type: DataTypes.ENUM('completed', 'failed', 'processing'),
      allowNull: false,
      defaultValue: 'processing',
    },
    errorMessage: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    ipAddress: {
      type: DataTypes.STRING,
      allowNull: true,
    },
    userAgent: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    userId: {
      type: DataTypes.INTEGER,
      allowNull: true,
      references: {
        model: 'users',
        key: 'id',
      },
    },
    createdAt: {
      type: DataTypes.DATE,
      allowNull: false,
    },
    updatedAt: {
      type: DataTypes.DATE,
      allowNull: false,
    },
  },
  {
    sequelize,
    modelName: 'Conversion',
    tableName: 'conversions',
    timestamps: true,
  }
);

// Note: Associations will be defined after all models are loaded

export default Conversion;
