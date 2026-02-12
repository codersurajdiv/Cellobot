/**
 * Resolve frontend model selector value to provider and model ID.
 */
function resolveModel(modelValue) {
  const val = (modelValue || 'claude-sonnet').toLowerCase();
  switch (val) {
    case 'claude-opus':
      return { provider: 'claude', modelId: 'claude-opus-4-20250918' };
    case 'claude-sonnet':
    case 'claude':
      return { provider: 'claude', modelId: 'claude-sonnet-4-5-20250514' };
    case 'openai-4o':
      return { provider: 'openai', modelId: 'gpt-4o' };
    case 'openai-4o-mini':
    case 'openai':
      return { provider: 'openai', modelId: 'gpt-4o-mini' };
    default:
      return { provider: 'claude', modelId: 'claude-sonnet-4-5-20250514' };
  }
}

module.exports = { resolveModel };
