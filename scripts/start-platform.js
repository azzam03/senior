const fs = require("fs");
const path = require("path");
const { spawn, spawnSync } = require("child_process");

const root = path.resolve(__dirname, "..");
const nodeModules = path.join(root, "node_modules");
const databaseDir = path.join(root, "database");
const backupDir = path.join(databaseDir, "backups");

function command(name) {
  return process.platform === "win32" ? `${name}.cmd` : name;
}

fs.mkdirSync(databaseDir, { recursive: true });
fs.mkdirSync(backupDir, { recursive: true });

if (!fs.existsSync(nodeModules)) {
  console.log("Dependencies not found. Running npm install...");
  const install = spawnSync(command("npm"), ["install"], {
    cwd: root,
    stdio: "inherit",
    shell: false,
  });
  if (install.status !== 0) {
    process.exit(install.status || 1);
  }
}

const child = spawn(process.execPath, ["--no-warnings", "server.js"], {
  cwd: root,
  stdio: "inherit",
  env: { ...process.env },
});

child.on("exit", (code) => process.exit(code || 0));
