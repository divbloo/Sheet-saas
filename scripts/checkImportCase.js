const fs = require("node:fs");
const path = require("node:path");
const { spawnSync } = require("node:child_process");

const SOURCE_EXTENSIONS = new Set([".js", ".jsx", ".mjs", ".cjs"]);
const RESOLUTION_SUFFIXES = ["", ".js", ".jsx", ".mjs", ".cjs", ".json"];

const hasExactPathFrom = (sourceDirectory, targetPath) => {
  const relativePath = path.relative(sourceDirectory, targetPath);
  const segments = relativePath.split(path.sep).filter(Boolean);
  let currentPath = sourceDirectory;

  for (const segment of segments) {
    if (segment === "..") {
      currentPath = path.dirname(currentPath);
      continue;
    }
    let entries;
    try {
      entries = fs.readdirSync(currentPath);
    } catch {
      return false;
    }
    if (!entries.includes(segment)) return false;
    currentPath = path.join(currentPath, segment);
  }

  return true;
};

const findCaseSensitiveTarget = (sourceDirectory, specifier) => {
  const cleanSpecifier = specifier.split(/[?#]/, 1)[0];
  const baseTarget = path.resolve(sourceDirectory, cleanSpecifier);
  const candidates = [
    ...RESOLUTION_SUFFIXES.map((suffix) => `${baseTarget}${suffix}`),
    ...RESOLUTION_SUFFIXES.slice(1).map((suffix) => path.join(baseTarget, `index${suffix}`)),
  ];

  return candidates.find((candidate) => (
    hasExactPathFrom(sourceDirectory, candidate) && fs.statSync(candidate).isFile()
  )) || null;
};

const extractRelativeImports = (source) => {
  const imports = new Set();
  const patterns = [
    /\b(?:require|import)\s*\(\s*["']([^"']+)["']\s*\)/g,
    /\b(?:import|export)\s+(?:[^"']*?\s+from\s+)?["']([^"']+)["']/g,
  ];

  patterns.forEach((pattern) => {
    for (const match of source.matchAll(pattern)) {
      if (match[1].startsWith(".")) imports.add(match[1]);
    }
  });

  return [...imports];
};

const getTrackedSourceFiles = (rootDirectory) => {
  const result = spawnSync("git", ["ls-files", "-z"], {
    cwd: rootDirectory,
    encoding: "utf8",
  });
  if (result.status !== 0) {
    throw new Error(result.stderr || "Unable to list tracked files");
  }

  return result.stdout
    .split("\0")
    .filter(Boolean)
    .filter((filename) => SOURCE_EXTENSIONS.has(path.extname(filename)));
};

const checkImportCase = (rootDirectory = process.cwd()) => {
  const failures = [];

  getTrackedSourceFiles(rootDirectory).forEach((filename) => {
    const absoluteFilename = path.resolve(rootDirectory, filename);
    const source = fs.readFileSync(absoluteFilename, "utf8");

    extractRelativeImports(source).forEach((specifier) => {
      if (!findCaseSensitiveTarget(path.dirname(absoluteFilename), specifier)) {
        failures.push(`${filename}: ${specifier}`);
      }
    });
  });

  return failures;
};

if (require.main === module) {
  const failures = checkImportCase();
  if (failures.length > 0) {
    console.error("Linux-incompatible relative imports found:");
    failures.forEach((failure) => console.error(`- ${failure}`));
    process.exitCode = 1;
  } else {
    console.log("All relative imports use Linux-compatible filename casing.");
  }
}

module.exports = {
  checkImportCase,
  extractRelativeImports,
  findCaseSensitiveTarget,
};
