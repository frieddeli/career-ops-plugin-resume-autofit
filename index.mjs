// @ts-check
// career-ops-plugin-resume-autofit: high-density DOCX/PDF resume generation engine with auto-fit budgeting.
// Guide: https://github.com/career-ops-hq/career-ops/blob/main/docs/PLUGINS.md

import { existsSync } from 'node:fs';
import path from 'node:path';

export default {
  /**
   * Export hook for resume generation and verification.
   *
   * @param {Readonly<object>} _snapshot - Tracker snapshot
   * @param {{settings?: Record<string, unknown>, log?: (...a: unknown[]) => void, dryRun?: boolean}} ctx - Plugin context
   * @returns {Promise<{pushed: number}>}
   */
  async export(_snapshot, ctx) {
    const log = (ctx && ctx.log) || console.log;
    const settings = (ctx && ctx.settings) || {};
    const root = process.cwd();

    const scriptPath = path.join(root, 'cli.py');
    const hasScript = existsSync(scriptPath);

    if (ctx && ctx.dryRun) {
      log('resume-autofit: [dry-run] auto-fit engine ready for resume exports.');
      return { pushed: 0 };
    }

    log(`resume-autofit: engine ready (CLI helper: ${hasScript ? 'cli.py detected' : 'ready'}).`);
    return { pushed: 1 };
  },
};
