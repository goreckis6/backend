module.exports = {
  apps: [{
    name: 'morphy-backend',
    script: 'dist/server.js',
    instances: 1,
    autorestart: true,
    watch: false,
    max_memory_restart: '1G',
    env: {
      NODE_ENV: 'production',
      PORT: 3000,
      FRONTEND_URL: 'https://morphyhub.com',
      FRONTEND_URL_ALT: 'https://www.morphyhub.com'
    }
  }]
};
