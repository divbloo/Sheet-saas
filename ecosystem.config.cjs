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
      wait_ready: true,
      listen_timeout: 20000,
      kill_timeout: 20000,
      max_restarts: 10,
      restart_delay: 5000,
      max_memory_restart: "500M",
      time: true,
      merge_logs: true,
      env: {
        NODE_ENV: "production",
        PORT: "5000",
      },
    },
  ],
};
