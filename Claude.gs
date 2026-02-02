/**
 * AI Integration for Google Apps Script
 * Supports multiple AI providers: Claude, OpenAI, Google Gemini
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

// ==================== Generic AI Call ====================

/**
 * Makes a request to the configured AI provider.
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

// ==================== Resume Analysis Functions ====================

/**
 * Analyzes how well a resume matches a job description.
 * Returns a match score and analysis.
 */
function analyzeResumeMatch(resumeData, jobDescription) {
  const systemPrompt = `You are an expert resume analyst and career coach. Your job is to objectively analyze how well a candidate's resume matches a job description.

Be honest and accurate in your assessment. Do not inflate scores to make the candidate feel good.

You must respond with valid JSON only, no other text.`;

  const userMessage = `Analyze this resume against the job description below.

RESUME:
${formatResumeForAnalysis(resumeData)}

JOB DESCRIPTION:
${jobDescription}

Provide your analysis as JSON with this exact structure:
{
  "matchScore": <number 0-100>,
  "summary": "<2-3 sentence overall assessment>",
  "companyName": "<company name from job posting>",
  "positionTitle": "<job title/position from job posting>",
  "strengths": ["<strength 1>", "<strength 2>", ...],
  "gaps": ["<gap 1>", "<gap 2>", ...],
  "keywordsMatched": ["<keyword 1>", "<keyword 2>", ...],
  "keywordsMissing": ["<keyword 1>", "<keyword 2>", ...],
  "recommendation": "<should they apply? honest advice>"
}`;

  const response = callAI(systemPrompt, userMessage);

  // Parse the JSON response
  try {
    // Handle potential markdown code blocks
    let jsonStr = response.trim();
    if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/```json?\n?/g, '').replace(/```$/g, '').trim();
    }
    return JSON.parse(jsonStr);
  } catch (e) {
    console.error('Failed to parse AI response:', response);
    throw new Error('Failed to parse analysis results');
  }
}

/**
 * Tailors a resume to better match a job description.
 * Maintains authenticity - only rephrases and reorganizes, doesn't fabricate.
 */
function tailorResume(resumeData, jobDescription, analysisResults) {
  const systemPrompt = `You are an expert resume writer who helps candidates present their authentic experience in the best light for specific roles.

CRITICAL RULES:
1. NEVER invent or fabricate skills, experiences, or achievements
2. NEVER add anything that isn't already in the resume
3. You CAN rephrase bullet points to better highlight relevant aspects
4. You CAN reorder bullet points to put most relevant first
5. You CAN adjust language to mirror the job description's terminology (where truthful)
6. You CAN make the summary/objective more targeted
7. You CAN update the candidate's headline/title to match the target position (e.g., if applying for "Senior Software Engineer", update their title line accordingly)
8. Maintain a natural, human voice - not robotic or overly formal
9. Keep the candidate's personality and authentic voice

The goal is to help the candidate's REAL qualifications shine through, not to create a fake persona.

You must respond with valid JSON only, no other text.`;

  const userMessage = `Tailor this resume for the job below. Remember: only rephrase and reorganize - do NOT add fake content.

CURRENT RESUME:
${formatResumeForAnalysis(resumeData)}

JOB DESCRIPTION:
${jobDescription}

ANALYSIS (for context):
- Match Score: ${analysisResults.matchScore}%
- Key gaps: ${analysisResults.gaps?.join(', ') || 'None identified'}
- Keywords to incorporate (if truthfully applicable): ${analysisResults.keywordsMissing?.join(', ') || 'None'}

Provide tailored content as JSON with this structure:
{
  "changes": [
    {
      "section": "<section name>",
      "original": "<original text>",
      "tailored": "<rephrased text>",
      "reason": "<why this change helps>"
    }
  ],
  "reorderSuggestions": [
    {
      "section": "<section name>",
      "suggestion": "<what to move and why>"
    }
  ],
  "overallNotes": "<any other advice for this application>"
}

Only include changes where rephrasing genuinely improves the match. Quality over quantity.`;

  const response = callAI(systemPrompt, userMessage);

  // Parse the JSON response
  try {
    let jsonStr = response.trim();
    if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/```json?\n?/g, '').replace(/```$/g, '').trim();
    }
    return JSON.parse(jsonStr);
  } catch (e) {
    console.error('Failed to parse AI response:', response);
    throw new Error('Failed to parse tailoring results');
  }
}

/**
 * Formats resume data for AI analysis.
 */
function formatResumeForAnalysis(resumeData) {
  let formatted = '';

  for (const section in resumeData.sections) {
    formatted += `\n## ${section}\n`;
    resumeData.sections[section].forEach(item => {
      formatted += `- ${item.content}\n`;
    });
  }

  return formatted.trim();
}
