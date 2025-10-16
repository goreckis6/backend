import sequelize from './connection.js';
import { User } from './models/User.js';
import { Conversion } from './models/Conversion.js';
import { AnonymousConversion } from './models/AnonymousConversion.js';

// Define associations
// Note: Associations are defined in the model files themselves

// Initialize database
export const initializeDatabase = async () => {
  try {
    console.log('🔍 Attempting to connect to database...');
    await sequelize.authenticate();
    console.log('✅ Database connection established successfully');
    
    console.log('🔍 Synchronizing database tables...');
    await sequelize.sync({ force: false }); // Create tables if they don't exist
    console.log('✅ Database synchronized successfully');
    
    // Test if tables exist
    const tables = await sequelize.getQueryInterface().showAllTables();
    console.log('📋 Available tables:', tables);
    
    return true;
  } catch (error) {
    console.error('❌ Database connection failed:', error);
    console.error('❌ Error details:', error instanceof Error ? error.message : 'Unknown error');
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

export { sequelize, User, Conversion, AnonymousConversion };
export default sequelize;
