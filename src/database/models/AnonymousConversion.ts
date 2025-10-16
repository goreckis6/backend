import { DataTypes, Model, Optional } from 'sequelize';
import sequelize from '../connection.js';

interface AnonymousConversionAttributes {
  id: number;
  ipAddress: string;
  userAgent?: string;
  conversionCount: number;
  lastConversionAt: Date;
  createdAt: Date;
  updatedAt: Date;
}

interface AnonymousConversionCreationAttributes extends Optional<AnonymousConversionAttributes, 'id' | 'userAgent' | 'conversionCount' | 'createdAt' | 'updatedAt'> {}

export class AnonymousConversion extends Model<AnonymousConversionAttributes, AnonymousConversionCreationAttributes> implements AnonymousConversionAttributes {
  declare id: number;
  declare ipAddress: string;
  declare userAgent?: string;
  declare conversionCount: number;
  declare lastConversionAt: Date;
  declare readonly createdAt: Date;
  declare readonly updatedAt: Date;
}

AnonymousConversion.init(
  {
    id: {
      type: DataTypes.INTEGER,
      autoIncrement: true,
      primaryKey: true,
    },
    ipAddress: {
      type: DataTypes.STRING,
      allowNull: false,
      unique: true,
      validate: {
        isIP: true,
      },
    },
    userAgent: {
      type: DataTypes.TEXT,
      allowNull: true,
    },
    conversionCount: {
      type: DataTypes.INTEGER,
      allowNull: false,
      defaultValue: 0,
      validate: {
        min: 0,
      },
    },
    lastConversionAt: {
      type: DataTypes.DATE,
      allowNull: false,
      defaultValue: DataTypes.NOW,
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
    modelName: 'AnonymousConversion',
    tableName: 'anonymous_conversions',
    timestamps: true,
    indexes: [
      {
        unique: true,
        fields: ['ipAddress'],
      },
      {
        fields: ['lastConversionAt'],
      },
    ],
  }
);

export default AnonymousConversion;
