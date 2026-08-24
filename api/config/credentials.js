require('dotenv').config();

const { bootstrapCredentials } = require('@librechat/api/credentials');

const credentials = bootstrapCredentials();

/**
 * Cost guard for this deployment.
 *
 * Enabled by default. Set LIBRECHAT_COST_MODE=false in Railway to restore the
 * upstream Anthropic defaults without reverting code.
 *
 * The Office tools now handle Excel/Word/PowerPoint locally, so native Claude
 * code execution does not need to be attached to every Anthropic request.
 */
if (process.env.LIBRECHAT_COST_MODE !== 'false') {
  try {
    const { anthropicSettings } = require('librechat-data-provider');

    if (anthropicSettings) {
      // Extended/adaptive thinking can create a large amount of billed output.
      anthropicSettings.thinking.default = false;

      // Do not resend the same uploaded files on every turn by default.
      anthropicSettings.resendFiles.default = false;

      // Keep normal office/chat responses bounded. Explicit user values are
      // still accepted, but are capped to avoid accidental runaway generations.
      const originalReset = anthropicSettings.maxOutputTokens.reset;
      const originalSet = anthropicSettings.maxOutputTokens.set;

      anthropicSettings.maxOutputTokens.default = 4096;
      anthropicSettings.maxOutputTokens.reset = (modelName) =>
        Math.min(originalReset(modelName), 4096);
      anthropicSettings.maxOutputTokens.set = (value, modelName) =>
        Math.min(originalSet(value, modelName), 8192);

      // Prompt caching is cheaper for repeated agent/system/tool context.
      anthropicSettings.promptCache.default = true;
    }

    // Native code execution was enabled while Excel support was being built.
    // The local Office tools now replace that need and avoid sending the native
    // code-execution tool on every Claude request.
    if (process.env.LIBRECHAT_KEEP_CODE_EXECUTION !== 'true') {
      process.env.ANTHROPIC_CODE_EXECUTION = 'false';
    }
  } catch (error) {
    console.warn('[cost-mode] Could not apply Anthropic cost defaults:', error?.message || error);
  }
}

module.exports = credentials;
