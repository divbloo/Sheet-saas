const { spawn } = require("node:child_process");

const isWindows = process.platform === "win32";
const npmCommand = isWindows ? "npm" : "npm";

const processes = [
  {
    name: "backend",
    args: ["--prefix", "backend", "start"],
  },
  {
    name: "frontend",
    args: ["--prefix", "frontend", "run", "dev"],
  },
];

const children = processes.map(({ name, args }) => {
  const child = spawn(npmCommand, args, {
    cwd: process.cwd(),
    stdio: "inherit",
    shell: isWindows,
  });

  child.on("exit", (code, signal) => {
    if (signal) {
      console.log(`${name} stopped with ${signal}`);
      return;
    }

    if (code && code !== 0) {
      console.error(`${name} exited with code ${code}`);
      stopAll();
      process.exit(code);
    }
  });

  return child;
});

const stopAll = () => {
  children.forEach((child) => {
    if (!child.killed) {
      child.kill();
    }
  });
};

process.on("SIGINT", () => {
  stopAll();
  process.exit(0);
});

process.on("SIGTERM", () => {
  stopAll();
  process.exit(0);
});
