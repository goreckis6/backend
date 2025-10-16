import sequelize from './connection.js';
import { User } from './models/User.js';
import { Conversion } from './models/Conversion.js';
import { AnonymousConversion } from './models/AnonymousConversion.js';
import { up as runMigrations } from './migrations.js';

// Define associations after models are imported
Conversion.belongsTo(User, { foreignKey: 'userId', as: 'user' });
User.hasMany(Conversion, { foreignKey: 'userId', as: 'conversions' });

// Initialize database
export const initializeDatabase = async () => {
  try {
    console.log('🔍 Attempting to connect to database...');
    await sequelize.authenticate();
    console.log('✅ Database connection established successfully');
    
    // Check if this is the first deployment (no users table or missing columns)
    const usersTableExists = await sequelize.getQueryInterface().tableExists('users');
    let needsForceSync = false;
    
    if (usersTableExists) {
      try {
        // Test if the table has the correct schema by trying to query it
        await User.findOne({ limit: 1 });
        console.log('✅ Users table schema is correct');
      } catch (error) {
        console.log('⚠️ Users table schema mismatch, will use migrations');
        needsForceSync = false; // Use migrations instead of force sync
      }
    } else {
      console.log('🆕 First deployment - creating tables');
      needsForceSync = true;
    }
    
    if (needsForceSync) {
      console.log('🔍 Creating tables (first deployment)...');
      await sequelize.sync({ force: true });
    } else {
      console.log('🔍 Synchronizing database tables...');
      await sequelize.sync({ alter: true }); // Alter tables to match model schema without dropping data
      
      // Run migrations to ensure all columns exist
      await runMigrations(sequelize.getQueryInterface());
    }
    
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
