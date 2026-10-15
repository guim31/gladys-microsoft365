// -----------------------------------------------------------------------------
// Entry point: builds the real Gladys client and hands it to src/app.js, which
// registers every handler before connect(). The SDK reads the supervisor's
// environment (GLADYS_HOST_API_URL, GLADYS_INTEGRATION_TOKEN,
// GLADYS_INTEGRATION_SELECTOR) by itself.
// -----------------------------------------------------------------------------

import { GladysIntegration, logger } from '@gladysassistant/integration-sdk';
import { registerHandlers } from './src/app.js';

const gladys = new GladysIntegration();
const app = registerHandlers(gladys);

// SIGTERM/SIGINT from the supervisor: stop the timer, then the SDK disconnects.
gladys.handleShutdown(() => {
  app.stop();
});

logger.info('Starting the Microsoft 365 integration');
gladys.connect().catch((err) => {
  logger.error('Initial connection failed', err);
  process.exit(1);
});
