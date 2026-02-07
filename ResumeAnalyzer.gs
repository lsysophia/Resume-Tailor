/**
 * Resume Analyzer - High-level functions for the sidebar UI
 */

// Default minimum match score required for tailoring
const DEFAULT_MIN_MATCH_SCORE = 80;

// ==================== User Profile Management ====================

/**
 * Gets the user's profile settings.
 * @returns {Object} User profile with industry, careerStage, etc.
 */
function getUserProfile() {
  const userProperties = PropertiesService.getUserProperties();
  const profileStr = userProperties.getProperty('USER_PROFILE');

  if (!profileStr) {
    return null;
  }

  try {
    return JSON.parse(profileStr);
  } catch (e) {
    return null;
  }
}

/**
 * Saves the user's profile settings.
 * @param {Object} profile - The profile object to save
 * @returns {Object} Result with success status
 */
function saveUserProfile(profile) {
  try {
    const userProperties = PropertiesService.getUserProperties();

    // Validate required fields
    if (!profile.industry || !profile.careerStage) {
      return {
        success: false,
        error: 'Industry and career stage are required.'
      };
    }

    const profileData = {
      industry: profile.industry,
      careerStage: profile.careerStage,
      specialization: profile.specialization || '',
      targetRoleType: profile.targetRoleType || '',
      updatedAt: new Date().toISOString()
    };

    userProperties.setProperty('USER_PROFILE', JSON.stringify(profileData));

    return {
      success: true,
      profile: profileData
    };
  } catch (e) {
    console.error('Error saving user profile:', e);
    return {
      success: false,
      error: e.message || 'Failed to save profile.'
    };
  }
}

/**
 * Checks if the user has completed the intake process.
 * @returns {boolean} True if profile exists
 */
function hasUserProfile() {
  return getUserProfile() !== null;
}

/**
 * Formats user profile context for AI prompts.
 * @returns {string} Formatted context string
 */
function formatUserProfileForAI() {
  const profile = getUserProfile();
  if (!profile) return '';

  let context = `\nCANDIDATE CONTEXT:`;
  context += `\n- Industry: ${profile.industry}`;
  context += `\n- Career Stage: ${profile.careerStage}`;

  if (profile.specialization) {
    context += `\n- Specialization/Focus: ${profile.specialization}`;
  }

  if (profile.targetRoleType) {
    context += `\n- Target Role Type: ${profile.targetRoleType}`;
  }

  return context;
}

// ==================== Basic Analysis (No AI Required) ====================

/**
 * Common technical skills and keywords to look for in job descriptions.
 * Organized by category for better matching.
 */
const COMMON_SKILLS = {
  languages: ['javascript', 'typescript', 'python', 'java', 'c++', 'c#', 'go', 'rust', 'ruby', 'php', 'swift', 'kotlin', 'scala', 'r', 'sql', 'html', 'css'],
  frameworks: ['react', 'angular', 'vue', 'node.js', 'express', 'django', 'flask', 'spring', 'rails', 'laravel', '.net', 'next.js', 'nuxt', 'svelte'],
  databases: ['postgresql', 'mysql', 'mongodb', 'redis', 'elasticsearch', 'dynamodb', 'oracle', 'sql server', 'sqlite', 'cassandra'],
  cloud: ['aws', 'azure', 'gcp', 'google cloud', 'heroku', 'digitalocean', 'cloudflare'],
  devops: ['docker', 'kubernetes', 'jenkins', 'ci/cd', 'terraform', 'ansible', 'github actions', 'gitlab ci'],
  tools: ['git', 'jira', 'confluence', 'slack', 'figma', 'sketch', 'postman', 'webpack', 'babel'],
  concepts: ['agile', 'scrum', 'rest api', 'graphql', 'microservices', 'tdd', 'unit testing', 'integration testing'],
  soft: ['leadership', 'communication', 'teamwork', 'problem-solving', 'analytical', 'detail-oriented', 'self-motivated']
};

/**
 * Extracts keywords from job description text.
 * @param {string} jobDescription - The job description text
 * @returns {Object} Extracted keywords by category
 */
function extractKeywordsFromJob(jobDescription) {
  const lowerJob = jobDescription.toLowerCase();
  const found = {
    skills: [],
    requirements: [],
    niceToHave: []
  };

  // Find all known skills mentioned in the job description
  for (const category in COMMON_SKILLS) {
    for (const skill of COMMON_SKILLS[category]) {
      // Use word boundary matching to avoid partial matches
      const regex = new RegExp('\\b' + skill.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
      if (regex.test(lowerJob)) {
        found.skills.push(skill);
      }
    }
  }

  // Extract years of experience requirements
  const yearsMatch = lowerJob.match(/(\d+)\+?\s*(?:years?|yrs?)\s*(?:of\s+)?(?:experience|exp)/gi);
  if (yearsMatch) {
    found.requirements.push(...yearsMatch.map(m => m.trim()));
  }

  // Extract degree requirements
  const degreePatterns = [
    /bachelor'?s?\s*(?:degree)?/gi,
    /master'?s?\s*(?:degree)?/gi,
    /ph\.?d\.?/gi,
    /computer science/gi,
    /software engineering/gi
  ];
  for (const pattern of degreePatterns) {
    const matches = lowerJob.match(pattern);
    if (matches) {
      found.requirements.push(...matches.map(m => m.trim()));
    }
  }

  // Remove duplicates
  found.skills = [...new Set(found.skills)];
  found.requirements = [...new Set(found.requirements)];

  return found;
}

/**
 * Performs basic resume analysis without AI.
 * Uses keyword matching to calculate a match score.
 * @param {Object} resumeData - The parsed resume data
 * @param {string} jobDescription - The job description text
 * @returns {Object} Analysis results
 */
function analyzeResumeMatchBasic(resumeData, jobDescription) {
  // Extract all resume text
  let resumeText = '';
  for (const section in resumeData.sections) {
    for (const item of resumeData.sections[section]) {
      resumeText += ' ' + item.content;
    }
  }
  const lowerResume = resumeText.toLowerCase();

  // Extract keywords from job description
  const jobKeywords = extractKeywordsFromJob(jobDescription);

  // Match skills against resume
  const matched = [];
  const missing = [];

  for (const skill of jobKeywords.skills) {
    const regex = new RegExp('\\b' + skill.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '\\b', 'i');
    if (regex.test(lowerResume)) {
      matched.push(skill);
    } else {
      missing.push(skill);
    }
  }

  // Calculate match score based on keyword overlap
  const totalKeywords = jobKeywords.skills.length;
  let matchScore = 0;

  if (totalKeywords > 0) {
    matchScore = Math.round((matched.length / totalKeywords) * 100);
  } else {
    // If no specific skills found, give a neutral score
    matchScore = 50;
  }

  // Cap score between reasonable bounds
  matchScore = Math.max(20, Math.min(95, matchScore));

  // Extract company and position from job description (basic extraction)
  const companyMatch = jobDescription.match(/(?:at|@|company[:\s]+|employer[:\s]+)([A-Z][A-Za-z0-9\s&]+?)(?:\.|,|\n|$)/);
  const positionMatch = jobDescription.match(/(?:position|role|title|job)[:\s]+([A-Za-z\s]+?)(?:\.|,|\n|$)/i) ||
                        jobDescription.match(/^([A-Z][A-Za-z\s]+(?:Engineer|Developer|Manager|Designer|Analyst|Specialist|Lead|Director))/m);

  return {
    matchScore: matchScore,
    summary: `Based on keyword analysis, your resume matches ${matched.length} of ${totalKeywords} key skills mentioned in the job description.`,
    companyName: companyMatch ? companyMatch[1].trim() : 'Company',
    positionTitle: positionMatch ? positionMatch[1].trim() : 'Position',
    strengths: matched.length > 0
      ? [`Resume contains ${matched.length} relevant keywords`, 'Skills alignment with job requirements']
      : ['Unable to determine specific strengths without AI analysis'],
    gaps: missing.length > 0
      ? [`${missing.length} skills from job description not found in resume`]
      : [],
    keywordsMatched: matched,
    keywordsToAddToSkills: [], // Can't determine without AI
    keywordsMissing: missing,
    skillsInferredFromExperience: [], // Can't determine without AI
    recommendation: matchScore >= 60
      ? 'Your resume shows keyword alignment with this role. For deeper insights and tailored suggestions, configure an AI provider in settings.'
      : 'Consider highlighting more relevant skills. For personalized recommendations, configure an AI provider in settings.',
    isBasicAnalysis: true // Flag to indicate this was done without AI
  };
}

/**
 * Gets the minimum match score threshold from user settings.
 */
function getMinMatchScore() {
  const userProperties = PropertiesService.getUserProperties();
  const stored = userProperties.getProperty('MIN_MATCH_SCORE');
  return stored ? parseInt(stored, 10) : DEFAULT_MIN_MATCH_SCORE;
}

/**
 * Saves the minimum match score threshold to user settings.
 */
function setMinMatchScore(score) {
  const userProperties = PropertiesService.getUserProperties();
  const value = Math.max(50, Math.min(100, parseInt(score, 10))); // Clamp between 50-100
  userProperties.setProperty('MIN_MATCH_SCORE', value.toString());
  return { success: true, minScore: value };
}

// Store the current working copy ID during a session
let currentCopyId = null;

/**
 * Main function called from the sidebar to analyze a job description.
 * Accepts either a URL or plain text job description.
 * @param {string} input - The job posting URL or text
 * @returns {Object} Analysis results
 */
function analyzeJob(input) {
  try {
    // Validate inputs
    if (!input || input.trim().length < 10) {
      return {
        success: false,
        error: 'Please provide a job description or URL.'
      };
    }

    let jobDescription = input.trim();

    // Check if input is a URL
    if (isUrl(jobDescription)) {
      const fetchResult = fetchJobDescription(jobDescription);
      if (!fetchResult.success) {
        return fetchResult;
      }
      jobDescription = fetchResult.jobDescription;
    }

    // Validate we have enough content
    if (jobDescription.length < 50) {
      return {
        success: false,
        error: 'Could not extract enough job description content. Please try pasting the job description directly.'
      };
    }

    if (!hasTemplate()) {
      return {
        success: false,
        error: 'Please set your resume as the template first (Add-ons > Resume Tailor > Set This Doc as Template).'
      };
    }

    // Get resume data from template
    const resumeData = getResumeData();

    if (Object.keys(resumeData.sections).length === 0) {
      return {
        success: false,
        error: 'Could not parse resume content. Please make sure your resume has clear section headers.'
      };
    }

    // Use AI analysis if available, otherwise fall back to basic keyword matching
    let analysis;
    if (hasApiKey()) {
      // Full AI-powered analysis
      analysis = analyzeResumeMatch(resumeData, jobDescription);
    } else {
      // Basic keyword matching (no AI required)
      analysis = analyzeResumeMatchBasic(resumeData, jobDescription);
    }

    // Get candidate name from resume
    const candidateName = getCandidateName(resumeData.docId);

    return {
      success: true,
      analysis: analysis,
      candidateName: candidateName,
      canTailor: analysis.matchScore >= getMinMatchScore(),
      minScoreRequired: getMinMatchScore()
    };

  } catch (error) {
    console.error('Analysis error:', error);
    return {
      success: false,
      error: error.message || 'An error occurred during analysis.'
    };
  }
}

/**
 * Tailors the resume based on the job description.
 * Creates a copy of the template and applies changes to it.
 * @param {string} jobDescription - The job posting text
 * @param {Object} analysisResults - Previous analysis results
 * @param {string} candidateName - The candidate's name from their resume
 * @returns {Object} Tailoring suggestions and copy info
 */
function tailorForJob(jobDescription, analysisResults, candidateName, additionalContext) {
  try {
    if (!hasApiKey()) {
      return {
        success: false,
        error: 'Please configure your Claude API key first.'
      };
    }

    if (!hasTemplate()) {
      return {
        success: false,
        error: 'Please set your resume as the template first.'
      };
    }

    // Note: We allow tailoring even below threshold - the UI shows a warning but lets user proceed

    // Create a copy of the template resume with proper naming
    const companyName = analysisResults.companyName || 'Company';
    const positionTitle = analysisResults.positionTitle || 'Position';
    const copyInfo = createResumeCopy(candidateName, companyName, positionTitle);

    if (!copyInfo.success) {
      return {
        success: false,
        error: 'Failed to create a copy of your resume.'
      };
    }

    // Get the tailoring suggestions (pass additional context if provided)
    const resumeData = getResumeData();
    const tailoring = tailorResume(resumeData, jobDescription, analysisResults, additionalContext || '');

    // Store pending changes for this copy so sidebar can show them when opened in the new doc
    if (tailoring.changes && tailoring.changes.length > 0) {
      storePendingChanges(copyInfo.copyId, tailoring.changes);
    }

    return {
      success: true,
      tailoring: tailoring,
      copy: copyInfo
    };

  } catch (error) {
    console.error('Tailoring error:', error);
    return {
      success: false,
      error: error.message || 'An error occurred during tailoring.'
    };
  }
}

/**
 * Applies a single tailored change to the copy document.
 * @param {string} copyId - The document ID of the copy
 * @param {string} section - Section name
 * @param {string} original - Original text to find
 * @param {string} tailored - New text to apply
 * @returns {Object} Result
 */
function applyChange(copyId, section, original, tailored) {
  try {
    if (!copyId) {
      return {
        success: false,
        error: 'No document ID provided for changes.'
      };
    }

    const result = applyChangeToDoc(copyId, original, tailored);
    return result;

  } catch (error) {
    console.error('Apply change error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Applies all changes to the copy document at once.
 * @param {string} copyId - The document ID of the copy
 * @param {Array} changes - Array of change objects
 * @returns {Object} Result
 */
function applyAllChanges(copyId, changes) {
  try {
    if (!copyId) {
      return {
        success: false,
        error: 'No document ID provided for changes.'
      };
    }

    let appliedCount = 0;
    let failedCount = 0;
    const results = [];

    for (const change of changes) {
      const result = applyChangeToDoc(copyId, change.original, change.tailored);
      results.push({
        section: change.section,
        success: result.success,
        message: result.message
      });

      if (result.success) {
        appliedCount++;
      } else {
        failedCount++;
      }
    }

    return {
      success: true,
      appliedCount: appliedCount,
      failedCount: failedCount,
      results: results,
      message: `Applied ${appliedCount} of ${changes.length} changes.`
    };

  } catch (error) {
    console.error('Apply all changes error:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Gets a preview of the resume structure.
 */
function getResumePreview() {
  try {
    const templateInfo = getTemplateInfo();

    if (!templateInfo.hasTemplate) {
      return {
        success: false,
        error: 'No template resume set.'
      };
    }

    const resumeData = getResumeData();
    const sections = Object.keys(resumeData.sections);

    return {
      success: true,
      templateName: templateInfo.templateName,
      sectionCount: sections.length,
      sections: sections,
      itemCounts: sections.reduce((acc, section) => {
        acc[section] = resumeData.sections[section].length;
        return acc;
      }, {}),
      totalItems: resumeData.elements.length
    };
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

// ==================== URL Handling ====================

/**
 * Checks if a string looks like a URL.
 */
function isUrl(str) {
  const urlPattern = /^(https?:\/\/)/i;
  return urlPattern.test(str.trim());
}

/**
 * Fetches job description from a URL.
 * @param {string} url - The job posting URL
 * @returns {Object} Result with jobDescription or error
 */
function fetchJobDescription(url) {
  try {
    // Fetch the page
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'User-Agent': 'Mozilla/5.0 (compatible; GoogleAppsScript)'
      }
    });

    const responseCode = response.getResponseCode();
    if (responseCode !== 200) {
      return {
        success: false,
        error: `Could not fetch the URL (status ${responseCode}). Please paste the job description directly.`
      };
    }

    const html = response.getContentText();

    // Extract text content from HTML
    const text = extractTextFromHtml(html);

    if (!text || text.length < 100) {
      return {
        success: false,
        error: 'Could not extract job description from URL. The page might require login or use JavaScript. Please paste the job description directly.'
      };
    }

    return {
      success: true,
      jobDescription: text
    };

  } catch (error) {
    console.error('URL fetch error:', error);
    return {
      success: false,
      error: 'Failed to fetch URL: ' + error.message + '. Please paste the job description directly.'
    };
  }
}

/**
 * Extracts readable text from HTML, removing scripts, styles, and tags.
 */
function extractTextFromHtml(html) {
  // Remove script and style content
  let text = html.replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '');
  text = text.replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '');
  text = text.replace(/<noscript[^>]*>[\s\S]*?<\/noscript>/gi, '');

  // Remove HTML comments
  text = text.replace(/<!--[\s\S]*?-->/g, '');

  // Replace common block elements with newlines
  text = text.replace(/<\/(p|div|h[1-6]|li|tr|br)[^>]*>/gi, '\n');
  text = text.replace(/<br[^>]*>/gi, '\n');

  // Remove all remaining HTML tags
  text = text.replace(/<[^>]+>/g, ' ');

  // Decode HTML entities
  text = text.replace(/&nbsp;/gi, ' ');
  text = text.replace(/&amp;/gi, '&');
  text = text.replace(/&lt;/gi, '<');
  text = text.replace(/&gt;/gi, '>');
  text = text.replace(/&quot;/gi, '"');
  text = text.replace(/&#39;/gi, "'");
  text = text.replace(/&rsquo;/gi, "'");
  text = text.replace(/&lsquo;/gi, "'");
  text = text.replace(/&rdquo;/gi, '"');
  text = text.replace(/&ldquo;/gi, '"');
  text = text.replace(/&mdash;/gi, '—');
  text = text.replace(/&ndash;/gi, '–');
  text = text.replace(/&bull;/gi, '•');

  // Clean up whitespace
  text = text.replace(/[ \t]+/g, ' ');
  text = text.replace(/\n\s*\n/g, '\n\n');
  text = text.trim();

  // Limit length to avoid API limits (keep first ~15000 chars which is plenty for a job description)
  if (text.length > 15000) {
    text = text.substring(0, 15000);
  }

  return text;
}

// ==================== Pending Changes Management ====================

/**
 * Stores pending changes for a tailored document copy.
 * This allows the sidebar to show the review UI when opened in the new doc.
 * @param {string} docId - The tailored document ID
 * @param {Array} changes - Array of change objects
 */
function storePendingChanges(docId, changes) {
  const userProperties = PropertiesService.getUserProperties();
  const data = {
    changes: changes,
    timestamp: new Date().toISOString()
  };
  userProperties.setProperty('PENDING_CHANGES_' + docId, JSON.stringify(data));
}

/**
 * Gets pending changes for the current document.
 * @returns {Object|null} Pending changes data or null if none
 */
function getPendingChangesForCurrentDoc() {
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) return null;

    const docId = doc.getId();
    const userProperties = PropertiesService.getUserProperties();
    const dataStr = userProperties.getProperty('PENDING_CHANGES_' + docId);

    if (!dataStr) return null;

    const data = JSON.parse(dataStr);
    return {
      docId: docId,
      docName: doc.getName(),
      changes: data.changes,
      timestamp: data.timestamp
    };
  } catch (e) {
    console.error('Error getting pending changes:', e);
    return null;
  }
}

/**
 * Clears pending changes for a document after they've been processed.
 * @param {string} docId - The document ID to clear
 */
function clearPendingChanges(docId) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty('PENDING_CHANGES_' + docId);
}

// ==================== AI-Powered Resume Analysis ====================

/**
 * Analyzes how well a resume matches a job description.
 * Returns a match score and analysis.
 *
 * IMPORTANT: This function detects skills not just from explicit skill listings,
 * but also infers them from work experience, projects, and achievements.
 */
function analyzeResumeMatch(resumeData, jobDescription) {
  // Get user profile context if available
  const profileContext = formatUserProfileForAI();

  const systemPrompt = `You are an expert resume analyst and career coach with deep technical knowledge. Your job is to objectively analyze how well a candidate's resume matches a job description.

Be honest and accurate in your assessment. Do not inflate scores to make the candidate feel good.
${profileContext ? `
CANDIDATE BACKGROUND:
Consider this context about the candidate when analyzing their fit:
${profileContext}

Use this context to:
- Understand their industry norms and terminology expectations
- Assess fit based on their career stage (entry-level candidates need different things than senior)
- Consider if they're a career changer who may have transferable skills
- Tailor your recommendations to their specific situation` : ''}

CRITICAL - TECHNICAL PRECISION:
You must be precise about technical terminology. Similar-sounding terms are often DIFFERENT things:
- "service-based architecture" (microservices/SOA) ≠ "service layer pattern" (code organization)
- "distributed systems" ≠ "client-server architecture"
- "CI/CD pipelines" ≠ "automated testing"
- "event-driven architecture" ≠ "using webhooks"
- "container orchestration" ≠ "using Docker"

Only claim a match when the candidate's experience ACTUALLY matches what the job is asking for, not just sounds similar.

CRITICAL: When analyzing skills, you must distinguish between:
1. Skills explicitly listed in the Skills section
2. Skills demonstrated through experience but NOT in Skills section (implicit skills)
3. Skills truly missing (not demonstrated anywhere)

For scoring, give credit for both explicit AND implicit skills, but ONLY when there's a genuine match.

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
  "keywordsMatched": ["<keywords explicitly in Skills section>", ...],
  "keywordsToAddToSkills": ["<keywords demonstrated in experience but NOT in Skills section>", ...],
  "keywordsMissing": ["<keywords truly not demonstrated anywhere in resume>", ...],
  "skillsInferredFromExperience": ["<skill inferred from work descriptions>", ...],
  "recommendation": "<should they apply? honest advice>"
}

IMPORTANT distinctions:
- keywordsMatched: Skills from job description that ARE in the candidate's Skills section
- keywordsToAddToSkills: Skills the candidate CLEARLY demonstrates in experience/projects but NOT in Skills section
  - ONLY include if there's EXACT or very close terminology match
  - "PostgreSQL" in experience → add "PostgreSQL" ✓
  - "built REST APIs" → add "REST APIs" ✓
  - "organized code into services" → add "service-based architecture" ✗ (different concept!)
- keywordsMissing: Skills from job description the candidate does NOT demonstrate anywhere

CRITICAL - BE CONSERVATIVE:
- When in doubt, put skills in keywordsMissing rather than keywordsToAddToSkills
- Do NOT infer architectural patterns unless explicitly described
- "Using X" is not the same as "expertise in X"
- Similar terminology does NOT mean same skill`;

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
 * Also suggests adding skills to the Skills section when the candidate has demonstrable experience.
 *
 * @param {Object} resumeData - The parsed resume data
 * @param {string} jobDescription - The job description text
 * @param {Object} analysisResults - Results from the analysis step
 * @param {string} additionalContext - Optional user-provided context about unlisted experience
 */
function tailorResume(resumeData, jobDescription, analysisResults, additionalContext) {
  // Get user profile context if available
  const profileContext = formatUserProfileForAI();

  const systemPrompt = `You are an expert resume writer with deep technical knowledge who helps candidates present their authentic experience in the best light for specific roles.
${profileContext ? `
CANDIDATE BACKGROUND:
${profileContext}

Use this context to inform your tailoring approach:
- Match the tone and terminology conventions of their industry
- For entry-level: emphasize potential, projects, education, eagerness to learn
- For mid-career: emphasize achievements, growth trajectory, expertise depth
- For senior/executive: emphasize leadership, strategic impact, business outcomes
- For career changers: emphasize transferable skills and relevant crossover experience
- For those returning to workforce: focus on continuous learning and updated skills` : ''}

CRITICAL RULES - AUTHENTICITY IS PARAMOUNT:
1. NEVER invent or fabricate skills, experiences, or achievements
2. Do NOT add terms/keywords just because they appear in the job description
3. Only mention skills/technologies that are EXPLICITLY demonstrated in the resume OR CONFIDENTLY confirmed by the candidate
4. You CAN rephrase bullet points to better highlight relevant aspects THAT ALREADY EXIST
5. You CAN reorder bullet points to put most relevant first
6. You CAN adjust language to use similar terminology WHERE THE CANDIDATE DEMONSTRABLY HAS THE SKILL
7. You CAN make the summary/objective more targeted using ONLY proven qualifications
8. Maintain a natural, human voice - not robotic or overly formal
9. Keep the candidate's personality and authentic voice

CRITICAL - TECHNICAL PRECISION:
Be precise about technical terminology. Similar-sounding terms are often DIFFERENT things:
- "service-based architecture" (microservices/SOA) ≠ "service layer pattern" (code organization)
- "distributed systems" ≠ "client-server architecture"
- "CI/CD pipelines" ≠ "automated testing"
- "event-driven architecture" ≠ "using webhooks"
- "container orchestration" ≠ "using Docker"

Do NOT substitute one term for another just because they sound similar!

STRICT EVIDENCE REQUIREMENT:
- Every change must be backed by specific evidence from the resume or candidate's additional context
- Do NOT add buzzwords like "automation", "scalable", "enterprise" unless the resume shows concrete examples
- If a job wants "automation experience" but the resume doesn't show it, do NOT add it

HANDLING CANDIDATE'S ADDITIONAL CONTEXT:
- If the candidate expresses uncertainty ("I'm not sure if...", "I don't know if this counts", "maybe"), do NOT treat it as confirmation
- Only use additional context when the candidate clearly confirms having the skill
- When in doubt, err on the side of NOT adding the skill

SKILL ADDITIONS - EXTREMELY CONSERVATIVE:
You may suggest adding skills to the Skills section ONLY when:
- The EXACT skill term is explicitly mentioned in the resume's experience/projects section, OR
- The candidate CONFIDENTLY confirms having the skill (not "I think" or "maybe")
- The terminology is an EXACT or very close match (not just similar-sounding)
- Never suggest skills just because the job description mentions them

The goal is to help the candidate's REAL qualifications shine through, not to keyword-stuff or create a fake persona.

You must respond with valid JSON only, no other text.`;

  // Build additional context section if provided
  const additionalContextSection = additionalContext && additionalContext.trim()
    ? `\nADDITIONAL CONTEXT FROM CANDIDATE:
The candidate provided this information about experience not fully reflected in their resume:
"${additionalContext}"

IMPORTANT: Parse this carefully:
- If they express uncertainty ("not sure", "maybe", "I think", "don't know if this counts"), do NOT treat it as confirmation
- Only use confident statements as evidence for adding skills
- If the terminology they use is different from the job posting's terminology, they might be describing a DIFFERENT thing`
    : '';

  const userMessage = `Tailor this resume for the job below.

STRICT REQUIREMENT: Only make changes backed by evidence from the resume or candidate's additional context. Do NOT add keywords/skills just because the job description mentions them.

CURRENT RESUME:
${formatResumeForAnalysis(resumeData)}

JOB DESCRIPTION:
${jobDescription}
${additionalContextSection}

ANALYSIS (for context):
- Match Score: ${analysisResults.matchScore}%
- Key gaps: ${analysisResults.gaps?.join(', ') || 'None identified'}
- Missing keywords (do NOT add these unless candidate has proven experience): ${analysisResults.keywordsMissing?.join(', ') || 'None'}

Provide tailored content as JSON with this structure:
{
  "changes": [
    {
      "section": "<section name>",
      "original": "<original text>",
      "tailored": "<rephrased text>",
      "reason": "<why this change helps>",
      "evidence": "<what in the resume/context proves this change is authentic>"
    }
  ],
  "contentToRemove": [
    {
      "section": "<section name>",
      "content": "<exact text to remove>",
      "reason": "<why removing this helps - be specific about irrelevance to the target role>"
    }
  ],
  "skillsToAdd": [
    {
      "skill": "<skill/keyword to add>",
      "category": "<suggested category in Skills section>",
      "justification": "<SPECIFIC quote from resume/context with EXACT terminology match>"
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

CRITICAL REQUIREMENTS:
- For "changes": Every tailored text must be provably true based on the resume or additional context
- For "skillsToAdd":
  - The justification MUST cite EXACT text from the resume or additional context
  - The skill term must EXACTLY match what's demonstrated (not a similar-sounding concept)
  - If the candidate expressed uncertainty, do NOT add the skill
  - If you can't cite verbatim evidence, do NOT suggest the skill
  - Example: "organized business logic into services" does NOT justify adding "service-based architecture" (different concepts!)
- For "contentToRemove": Only suggest if clearly irrelevant to the target role
- CRITICAL: The "content" field in contentToRemove must be copied EXACTLY from the resume - character for character.

Quality over quantity. When in doubt, leave it out. Return empty arrays if nothing meets the strict evidence bar.`;

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
 * Regenerates a single tailoring suggestion based on user feedback.
 * Used when the AI's suggestion is inaccurate or embellished.
 *
 * @param {string} original - The original text from the resume
 * @param {string} currentTailored - The current AI-generated suggestion
 * @param {string} feedback - User's feedback about what to change
 * @param {string} jobDescription - The job description for context
 * @returns {Object} Result with new tailored text and reason
 */
function rephraseChange(original, currentTailored, feedback, jobDescription) {
  try {
    if (!hasApiKey()) {
      return {
        success: false,
        error: 'Please configure your AI API key first.'
      };
    }

    const systemPrompt = `You are an expert resume writer helping to refine a resume suggestion based on user feedback.

CRITICAL RULES:
1. The user has provided feedback about a suggestion that was inaccurate or embellished
2. You MUST respect the user's feedback completely - if they say they don't have experience with something, DO NOT include it
3. Only use skills, technologies, and experiences that are provably in the original text
4. Do NOT add any terms or skills that are not explicitly mentioned in the original resume text
5. Keep the suggestion authentic to the candidate's actual experience
6. Maintain a natural, professional tone

You must respond with valid JSON only, no other text.`;

    const userMessage = `The user wants to refine this resume suggestion.

ORIGINAL TEXT FROM RESUME:
${original}

CURRENT AI SUGGESTION (needs refinement):
${currentTailored}

USER FEEDBACK - THIS IS CRITICAL, YOU MUST FOLLOW IT:
${feedback}

JOB DESCRIPTION (for context only - do NOT add skills from here that aren't in the original):
${jobDescription ? jobDescription.substring(0, 2000) : 'Not provided'}

Based on the user's feedback, provide a revised suggestion. Remember:
- If the user says they don't have experience with something, DO NOT mention it
- Only include skills/technologies actually mentioned in the original text
- The suggestion should be authentic to the candidate's real experience

Respond with JSON:
{
  "tailored": "<your revised suggestion that respects the user's feedback>",
  "reason": "<brief explanation of the improvement and how you addressed the feedback>"
}`;

    const response = callAI(systemPrompt, userMessage);

    // Parse the JSON response
    let jsonStr = response.trim();
    if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/```json?\n?/g, '').replace(/```$/g, '').trim();
    }

    const result = JSON.parse(jsonStr);

    return {
      success: true,
      tailored: result.tailored,
      reason: result.reason
    };

  } catch (error) {
    console.error('Rephrase error:', error);
    return {
      success: false,
      error: error.message || 'Failed to regenerate suggestion'
    };
  }
}

/**
 * Uses AI to determine the correct category/subcategory for placing a skill in the resume.
 * Analyzes the resume structure to find where the skill fits best.
 *
 * @param {string} skill - The skill to place (e.g., "LLM", "Kubernetes", "React")
 * @param {Object} resumeData - The parsed resume data with sections
 * @returns {Object} Result with category info and placement instructions
 */
function determineSkillCategory(skill, resumeData) {
  const systemPrompt = `You are an expert resume formatter. Your job is to determine exactly where a new skill should be placed in an existing resume's Skills section.

You must analyze the resume's skill organization and determine the MOST appropriate category/subcategory for the new skill.

You must respond with valid JSON only, no other text.`;

  const userMessage = `Given this resume structure, determine where to place the skill "${skill}".

RESUME SKILLS SECTION:
${formatSkillsSectionForAnalysis(resumeData)}

FULL RESUME CONTEXT:
${formatResumeForAnalysis(resumeData)}

Analyze the resume's skill categories and determine the best placement for "${skill}".

Respond with JSON:
{
  "skill": "${skill}",
  "recommendedCategory": "<exact category name as it appears in resume, e.g., 'Integration & Services' or 'Frontend Technologies'>",
  "reasoning": "<brief explanation of why this category fits>",
  "existingSkillsInCategory": ["<list of skills already in that category>"],
  "alternativeCategories": ["<other possible categories if primary doesn't fit>"]
}

IMPORTANT:
- Use the EXACT category name as it appears in the resume (match capitalization and punctuation)
- If the resume has subcategories under "Skills" like "Programming Languages:", "Frontend:", etc., use those exact subcategory names
- If no good category exists, set recommendedCategory to "Skills" or "Technical Skills" (whichever exists)`;

  const response = callAI(systemPrompt, userMessage);

  try {
    let jsonStr = response.trim();
    if (jsonStr.startsWith('```')) {
      jsonStr = jsonStr.replace(/```json?\n?/g, '').replace(/```$/g, '').trim();
    }
    return JSON.parse(jsonStr);
  } catch (e) {
    console.error('Failed to parse AI skill category response:', response);
    // Return a default response
    return {
      skill: skill,
      recommendedCategory: 'Skills',
      reasoning: 'Could not determine specific category',
      existingSkillsInCategory: [],
      alternativeCategories: []
    };
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

/**
 * Formats just the skills section of the resume for category analysis.
 */
function formatSkillsSectionForAnalysis(resumeData) {
  let formatted = '';
  const skillKeywords = ['SKILLS', 'TECHNICAL', 'COMPETENCIES', 'EXPERTISE', 'TECHNOLOGIES'];

  for (const section in resumeData.sections) {
    const upperSection = section.toUpperCase();
    if (skillKeywords.some(kw => upperSection.includes(kw))) {
      formatted += `\n## ${section}\n`;
      resumeData.sections[section].forEach(item => {
        formatted += `- ${item.content}\n`;
      });
    }
  }

  // If no skills section found, return full resume context
  if (!formatted) {
    return 'No explicit skills section found. Skills may be listed within other sections.';
  }

  return formatted.trim();
}
