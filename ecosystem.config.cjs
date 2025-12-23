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
      FRONTEND_URL: 'https://formipeek.com',
      FRONTEND_URL_ALT: 'https://www.formipeek.com'
    }
  }]
};
