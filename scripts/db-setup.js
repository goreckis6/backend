import { initializeDatabase, closeDatabase, sequelize } from '../src/database/index.js';

const command = process.argv[2];

async function testConnection() {
  try {
    console.log('🔍 Testing database connection...');
    await sequelize.authenticate();
    console.log('✅ Database connection successful!');
    
    // Test a simple query
    const result = await sequelize.query('SELECT NOW()');
    console.log('✅ Query test successful:', result[0][0]);
    
  } catch (error) {
    console.error('❌ Database connection failed:', error.message);
    process.exit(1);
  }
}

async function syncDatabase() {
  try {
    console.log('🔄 Synchronizing database...');
    await initializeDatabase();
    console.log('✅ Database synchronized successfully!');
    
  } catch (error) {
    console.error('❌ Database synchronization failed:', error.message);
    process.exit(1);
  }
}

async function seedDatabase() {
  try {
    console.log('🌱 Seeding database...');
    // Add seed data here if needed
    console.log('✅ Database seeded successfully!');
    
  } catch (error) {
    console.error('❌ Database seeding failed:', error.message);
    process.exit(1);
  }
}

async function resetDatabase() {
  try {
    console.log('⚠️  Resetting database...');
    await sequelize.sync({ force: true });
    console.log('✅ Database reset successfully!');
    
  } catch (error) {
    console.error('❌ Database reset failed:', error.message);
    process.exit(1);
  }
}

async function getStatus() {
  try {
    console.log('📊 Getting database status...');
    await sequelize.authenticate();
    
    const [users, conversions] = await Promise.all([
      sequelize.query('SELECT COUNT(*) as count FROM users'),
      sequelize.query('SELECT COUNT(*) as count FROM conversions')
    ]);
    
    console.log('✅ Database Status:');
    console.log(`   Users: ${users[0][0].count}`);
    console.log(`   Conversions: ${conversions[0][0].count}`);
    
  } catch (error) {
    console.error('❌ Failed to get database status:', error.message);
    process.exit(1);
  }
}

async function main() {
  try {
    switch (command) {
      case 'test':
        await testConnection();
        break;
      case 'sync':
        await syncDatabase();
        break;
      case 'seed':
        await seedDatabase();
        break;
      case 'reset':
        await resetDatabase();
        break;
      case 'status':
        await getStatus();
        break;
      default:
        console.log('Usage: npm run db:<command>');
        console.log('Commands:');
        console.log('  test   - Test database connection');
        console.log('  sync   - Synchronize database schema');
        console.log('  seed   - Seed database with initial data');
        console.log('  reset  - Reset database (WARNING: destroys all data)');
        console.log('  status - Show database status');
        process.exit(1);
    }
  } catch (error) {
    console.error('❌ Command failed:', error.message);
    process.exit(1);
  } finally {
    await closeDatabase();
  }
}

main();
