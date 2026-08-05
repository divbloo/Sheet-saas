const path = require("path");

module.exports = {
  apps: [
    {
      name: "sheet-saas",
      cwd: path.join(__dirname, "backend"),
      script: "server.js",
      interpreter: "node",
      instances: 1,
      exec_mode: "fork",
      autorestart: true,
      watch: false,
      max_restarts: 10,
      restart_delay: 5000,
      env: {
        NODE_ENV: "production",
        PORT: "5000",
      },
    },
  ],
};
