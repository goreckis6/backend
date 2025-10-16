// Simple test script to test registration
// Using built-in fetch (Node.js 18+)

async function testRegistration() {
  try {
    console.log('🔍 Testing registration...');
    
    const response = await fetch('https://morphy-2-n2tb.onrender.com/api/auth/register', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        email: 'test@example.com',
        password: 'testpassword123',
        name: 'Test User'
      })
    });
    
    const result = await response.json();
    console.log('📊 Response status:', response.status);
    console.log('📊 Response body:', result);
    
    if (result.success) {
      console.log('✅ Registration successful!');
    } else {
      console.log('❌ Registration failed:', result.error);
    }
    
  } catch (error) {
    console.error('❌ Test error:', error.message);
  }
}

testRegistration();
