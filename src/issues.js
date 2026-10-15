// -----------------------------------------------------------------------------
// Repeated failures, logged once: a warning when a problem starts (or its
// message changes), debug while it lasts, one info line when it is over.
// -----------------------------------------------------------------------------

/**
 * @param {{ debug: Function, info: Function, warn: Function }} logger
 */
export function createIssueLog(logger) {
  /** @type {Map<string, string>} the message logged, by problem key */
  const ongoing = new Map();
  return {
    /** A failure: a warning the first time (or when it changes), debug after. */
    fail(key, message) {
      if (ongoing.get(key) === message) {
        logger.debug(`${message} (still failing)`);
        return;
      }
      ongoing.set(key, message);
      logger.warn(message);
    },

    /** A success: one info line when it ends a failure, nothing otherwise. */
    ok(key, message) {
      if (ongoing.delete(key) && message) {
        logger.info(message);
      }
    },

    /** Forget a problem without a recovery line (its subject is gone). */
    forget(key) {
      ongoing.delete(key);
    },

    isFailing(key) {
      return ongoing.has(key);
    },
  };
}
