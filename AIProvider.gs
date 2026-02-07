/**
 * AI Provider Integration for Google Apps Script
 * Supports multiple AI providers: Claude, OpenAI, Google Gemini, DeepSeek
 *
 * This file contains only generic AI infrastructure.
 * Resume-specific AI functions are in ResumeAnalyzer.gs
 */

// ==================== AI Provider Configuration ====================

const AI_PROVIDERS = {
  claude: {
    name: 'Claude (Anthropic)',
    apiUrl: 'https://api.anthropic.com/v1/messages',
    defaultModel: 'claude-sonnet-4-20250514',
    models: ['claude-sonnet-4-20250514', 'claude-3-5-sonnet-20241022', 'claude-3-haiku-20240307']
  },
  openai: {
    name: 'GPT (OpenAI)',
    apiUrl: 'https://api.openai.com/v1/chat/completions',
    defaultModel: 'gpt-4o',
    models: ['gpt-4o', 'gpt-4o-mini', 'gpt-4-turbo']
  },
  gemini: {
    name: 'Gemini (Google)',
    apiUrl: 'https://generativelanguage.googleapis.com/v1beta/models/',
    defaultModel: 'gemini-1.5-pro',
    models: ['gemini-1.5-pro', 'gemini-1.5-flash', 'gemini-2.0-flash']
  },
  deepseek: {
    name: 'DeepSeek',
    apiUrl: 'https://api.deepseek.com/v1/chat/completions',
    defaultModel: 'deepseek-chat',
    models: ['deepseek-chat', 'deepseek-reasoner']
  }
};

// ==================== Settings Management ====================

/**
 * Gets the current AI provider setting.
 */
function getAIProvider() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('AI_PROVIDER') || 'claude';
}

/**
 * Sets the AI provider.
 */
function setAIProvider(provider) {
  if (!AI_PROVIDERS[provider]) {
    throw new Error('Invalid AI provider: ' + provider);
  }
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('AI_PROVIDER', provider);
  return { success: true, provider: provider };
}

/**
 * Gets the API key for a specific provider.
 */
function getProviderApiKey(provider) {
  const userProperties = PropertiesService.getUserProperties();
  const key = userProperties.getProperty(`API_KEY_${provider.toUpperCase()}`);
  // Fallback to old key name for Claude
  if (!key && provider === 'claude') {
    return userProperties.getProperty('CLAUDE_API_KEY') || '';
  }
  return key || '';
}

/**
 * Saves the API key for a specific provider.
 */
function saveProviderApiKey(provider, apiKey) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(`API_KEY_${provider.toUpperCase()}`, apiKey);
  // Also save to old key name for Claude (backward compatibility)
  if (provider === 'claude') {
    userProperties.setProperty('CLAUDE_API_KEY', apiKey);
  }
  return { success: true, message: `${AI_PROVIDERS[provider].name} API key saved!` };
}

/**
 * Gets all AI settings for the UI.
 */
function getAISettings() {
  const currentProvider = getAIProvider();
  const providers = Object.keys(AI_PROVIDERS).map(key => ({
    id: key,
    name: AI_PROVIDERS[key].name,
    hasKey: !!getProviderApiKey(key),
    models: AI_PROVIDERS[key].models
  }));

  return {
    currentProvider: currentProvider,
    providers: providers
  };
}

/**
 * Checks if the current AI provider has an API key configured.
 * @returns {boolean} True if the current provider has an API key set
 */
function hasApiKey() {
  const provider = getAIProvider();
  const apiKey = getProviderApiKey(provider);
  return !!apiKey;
}

// ==================== Generic AI Call ====================

/**
 * Makes a request to the configured AI provider.
 * @param {string} systemPrompt - The system/context prompt
 * @param {string} userMessage - The user message/query
 * @returns {string} The AI response text
 */
function callAI(systemPrompt, userMessage) {
  const provider = getAIProvider();
  const apiKey = getProviderApiKey(provider);

  if (!apiKey) {
    throw new Error(`${AI_PROVIDERS[provider].name} API key not configured. Please set it in the add-on settings.`);
  }

  switch (provider) {
    case 'claude':
      return callClaude(systemPrompt, userMessage, apiKey);
    case 'openai':
      return callOpenAI(systemPrompt, userMessage, apiKey);
    case 'gemini':
      return callGemini(systemPrompt, userMessage, apiKey);
    case 'deepseek':
      return callDeepSeek(systemPrompt, userMessage, apiKey);
    default:
      throw new Error('Unknown AI provider: ' + provider);
  }
}

// ==================== Provider Implementations ====================

/**
 * Makes a request to Claude API.
 */
function callClaude(systemPrompt, userMessage, apiKey) {
  const config = AI_PROVIDERS.claude;

  const payload = {
    model: config.defaultModel,
    max_tokens: 4096,
    system: systemPrompt,
    messages: [
      { role: 'user', content: userMessage }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'x-api-key': apiKey,
      'anthropic-version': '2023-06-01'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(config.apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const error = JSON.parse(responseText);
    throw new Error(error.error?.message || `Claude API error: ${responseCode}`);
  }

  const data = JSON.parse(responseText);
  return data.content[0].text;
}

/**
 * Makes a request to OpenAI API.
 */
function callOpenAI(systemPrompt, userMessage, apiKey) {
  const config = AI_PROVIDERS.openai;

  const payload = {
    model: config.defaultModel,
    max_tokens: 4096,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userMessage }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(config.apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const error = JSON.parse(responseText);
    throw new Error(error.error?.message || `OpenAI API error: ${responseCode}`);
  }

  const data = JSON.parse(responseText);
  return data.choices[0].message.content;
}

/**
 * Makes a request to Google Gemini API.
 */
function callGemini(systemPrompt, userMessage, apiKey) {
  const config = AI_PROVIDERS.gemini;
  const apiUrl = `${config.apiUrl}${config.defaultModel}:generateContent?key=${apiKey}`;

  const payload = {
    contents: [
      {
        parts: [
          { text: systemPrompt + '\n\n' + userMessage }
        ]
      }
    ],
    generationConfig: {
      maxOutputTokens: 4096
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const error = JSON.parse(responseText);
    throw new Error(error.error?.message || `Gemini API error: ${responseCode}`);
  }

  const data = JSON.parse(responseText);
  return data.candidates[0].content.parts[0].text;
}

/**
 * Makes a request to DeepSeek API (OpenAI-compatible).
 */
function callDeepSeek(systemPrompt, userMessage, apiKey) {
  const config = AI_PROVIDERS.deepseek;

  const payload = {
    model: config.defaultModel,
    max_tokens: 4096,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userMessage }
    ]
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(config.apiUrl, options);
  const responseCode = response.getResponseCode();
  const responseText = response.getContentText();

  if (responseCode !== 200) {
    const error = JSON.parse(responseText);
    throw new Error(error.error?.message || `DeepSeek API error: ${responseCode}`);
  }

  const data = JSON.parse(responseText);
  return data.choices[0].message.content;
}
