import express, { type Express } from "express";
import cors from "cors";
import pinoHttp from "pino-http";
import { createProxyMiddleware } from "http-proxy-middleware";
import router from "./routes";
import { logger } from "./lib/logger";

const app: Express = express();

app.use(
  pinoHttp({
    logger,
    serializers: {
      req(req) {
        return {
          id: req.id,
          method: req.method,
          url: req.url?.split("?")[0],
        };
      },
      res(res) {
        return {
          statusCode: res.statusCode,
        };
      },
    },
  }),
);
app.use(cors());

// Body parsing only for /api routes — do NOT apply globally or the stream
// gets consumed before http-proxy-middleware can forward it to Flask.
app.use("/api", express.json(), express.urlencoded({ extended: true }), router);

// Proxy all non-/api requests to the Flask ContractVault app.
// Must come after /api so body parsing never runs on proxied requests.
app.use(
  "/",
  createProxyMiddleware({
    target: "http://localhost:8000",
    changeOrigin: true,
    ws: true,
  }),
);

export default app;
