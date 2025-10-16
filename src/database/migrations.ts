import { QueryInterface, DataTypes } from 'sequelize';

// Migration to ensure all required columns exist
export const up = async (queryInterface: QueryInterface): Promise<void> => {
  try {
    console.log('🔍 Running database migrations...');

    // Check if users table exists and has required columns
    const usersTableExists = await queryInterface.tableExists('users');
    if (usersTableExists) {
      const tableDescription = await queryInterface.describeTable('users');
      
      // Add missing columns if they don't exist
      if (!tableDescription.name) {
        await queryInterface.addColumn('users', 'name', {
          type: DataTypes.STRING,
          allowNull: true
        });
        console.log('✅ Added name column to users table');
      }

      if (!tableDescription.isActive) {
        await queryInterface.addColumn('users', 'isActive', {
          type: DataTypes.BOOLEAN,
          allowNull: false,
          defaultValue: true
        });
        console.log('✅ Added isActive column to users table');
      }

      if (!tableDescription.createdAt) {
        await queryInterface.addColumn('users', 'createdAt', {
          type: DataTypes.DATE,
          allowNull: false,
          defaultValue: DataTypes.NOW
        });
        console.log('✅ Added createdAt column to users table');
      }

      if (!tableDescription.updatedAt) {
        await queryInterface.addColumn('users', 'updatedAt', {
          type: DataTypes.DATE,
          allowNull: false,
          defaultValue: DataTypes.NOW
        });
        console.log('✅ Added updatedAt column to users table');
      }
    }

    console.log('✅ Database migrations completed successfully');
  } catch (error) {
    console.error('❌ Migration error:', error);
    throw error;
  }
};

export const down = async (queryInterface: QueryInterface): Promise<void> => {
  // Rollback migrations if needed
  console.log('⚠️ Rolling back migrations...');
};
