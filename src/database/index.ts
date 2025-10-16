import sequelize from './connection.js';
import { User } from './models/User.js';
import { Conversion } from './models/Conversion.js';

// Define associations
Conversion.belongsTo(User, { foreignKey: 'userId', as: 'user' });
User.hasMany(Conversion, { foreignKey: 'userId', as: 'conversions' });

// Initialize database
export const initializeDatabase = async () => {
  try {
    await sequelize.authenticate();
    console.log('✅ Database connection established successfully');
    
    await sequelize.sync({ alter: false }); // Set to true if you want to alter tables
    console.log('✅ Database synchronized successfully');
    
    return true;
  } catch (error) {
    console.error('❌ Database connection failed:', error);
    throw error;
  }
};

// Close database connection
export const closeDatabase = async () => {
  try {
    await sequelize.close();
    console.log('✅ Database connection closed');
  } catch (error) {
    console.error('❌ Error closing database connection:', error);
  }
};

export { sequelize, User, Conversion };
export default sequelize;
