/**
 * Always-on console logging for the add-in (independent of Settings → Verbose debug trace).
 * Use for high-signal lifecycle events; keep volume reasonable.
 */

/**
 * @param {string} category e.g. "translate", "word", "llm"
 * @param {string} message
 * @param {unknown} [detail]
 */
export function appLog(category, message, detail) {
  const tag = `[Legal AI] [${category}]`;
  if (detail !== undefined) {
    console.info(tag, message, detail);
  } else {
    console.info(tag, message);
  }
}

/**
 * @param {string} category
 * @param {string} message
 * @param {unknown} [detail]
 */
export function appWarn(category, message, detail) {
  const tag = `[Legal AI] [${category}]`;
  if (detail !== undefined) {
    console.warn(tag, message, detail);
  } else {
    console.warn(tag, message);
  }
}
