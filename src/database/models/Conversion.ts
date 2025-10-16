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
  public id!: number;
  public originalFilename!: string;
  public convertedFilename!: string;
  public originalFormat!: string;
  public convertedFormat!: string;
  public fileSize!: number;
  public conversionTime?: number;
  public status!: 'completed' | 'failed' | 'processing';
  public errorMessage?: string;
  public ipAddress?: string;
  public userAgent?: string;
  public userId?: number;
  public readonly createdAt!: Date;
  public readonly updatedAt!: Date;
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
  },
  {
    sequelize,
    modelName: 'Conversion',
    tableName: 'conversions',
    timestamps: true,
  }
);

export default Conversion;
