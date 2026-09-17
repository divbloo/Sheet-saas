const DEFAULT_PORT = 5000;

const requireValue = (environment, name) => {
  const value = String(environment[name] || "").trim();
  if (!value) throw new Error(`Missing required environment variable: ${name}`);
  return value;
};

const parsePort = (value) => {
  if (value === undefined || value === "") return DEFAULT_PORT;

  const port = Number(value);
  if (!Number.isInteger(port) || port < 1 || port > 65535) {
    throw new Error("PORT must be an integer between 1 and 65535");
  }
  return port;
};

const parseTrustProxy = (value, isProduction) => {
  if (value === undefined || value === "") return isProduction ? 1 : false;
  if (String(value).toLowerCase() === "false") return false;

  const hops = Number(value);
  if (!Number.isInteger(hops) || hops < 0 || hops > 10) {
    throw new Error("TRUST_PROXY must be false or an integer between 0 and 10");
  }
  return hops;
};

const validateRuntimeEnvironment = (environment = process.env) => {
  const nodeEnv = String(environment.NODE_ENV || "development").trim();
  const isProduction = nodeEnv === "production";
  const mongoUri = requireValue(environment, "MONGO_URI");
  const jwtSecret = requireValue(environment, "JWT_SECRET");

  if (isProduction && jwtSecret.length < 32) {
    throw new Error("JWT_SECRET must be at least 32 characters in production");
  }
  if (isProduction) requireValue(environment, "FRONTEND_URL");

  return {
    isProduction,
    jwtSecret,
    mongoUri,
    nodeEnv,
    port: parsePort(environment.PORT),
    trustProxy: parseTrustProxy(environment.TRUST_PROXY, isProduction),
  };
};

const getReadiness = (databaseReadyState) => {
  const isReady = databaseReadyState === 1;
  return { isReady, status: isReady ? "ok" : "unavailable" };
};

const createGracefulShutdown = ({
  closeRealtime,
  closeHttp,
  closeDatabase,
  logger = console,
}) => {
  let shutdownPromise = null;

  return (signal) => {
    if (shutdownPromise) return shutdownPromise;

    shutdownPromise = (async () => {
      logger.info(`Received ${signal}; shutting down gracefully`);
      try {
        await closeRealtime();
        await closeHttp();
        await closeDatabase();
        logger.info("Graceful shutdown complete");
      } catch (error) {
        logger.error("Graceful shutdown failed", error);
        throw error;
      }
    })();

    return shutdownPromise;
  };
};

module.exports = {
  createGracefulShutdown,
  getReadiness,
  validateRuntimeEnvironment,
};
