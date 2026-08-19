import { createApp } from "./app.js";
import { config } from "./config.js";
import { ensureStorageDir } from "./store/sessions.js";
import { log } from "./logging.js";

await ensureStorageDir();
const app = createApp();
app.listen(config.port, () => {
  log.info(`${config.productName} server listening`, {
    port: config.port,
    env: config.env,
    metadataProvider: config.metadataProvider,
  });
});
