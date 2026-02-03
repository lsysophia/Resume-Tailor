/**
 * Resume Analyzer - High-level functions for the sidebar UI
 */

// Default minimum match score required for tailoring
const DEFAULT_MIN_MATCH_SCORE = 80;

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

    if (!hasApiKey()) {
      return {
        success: false,
        error: 'Please configure your AI API key first (Add-ons > Resume Tailor > AI Settings).'
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

    // Analyze the match (Claude extracts company and position)
    const analysis = analyzeResumeMatch(resumeData, jobDescription);

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
function tailorForJob(jobDescription, analysisResults, candidateName) {
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

    // Get the tailoring suggestions
    const resumeData = getResumeData();
    const tailoring = tailorResume(resumeData, jobDescription, analysisResults);

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
  const systemPrompt = `You are an expert resume analyst and career coach. Your job is to objectively analyze how well a candidate's resume matches a job description.

Be honest and accurate in your assessment. Do not inflate scores to make the candidate feel good.

CRITICAL: When analyzing skills, you must:
1. Look at BOTH explicit skill listings AND infer skills from experience descriptions
2. If someone describes "building AI-powered features" or "integrating ChatGPT", they have AI/LLM skills
3. If someone "developed REST APIs", they have API development skills
4. If someone "deployed to AWS", they have cloud/AWS skills
5. Technologies mentioned in project descriptions count as skills the candidate has
6. Don't mark a skill as "missing" if the candidate demonstrates it through their work experience

For "keywordsMissing", only include skills/technologies that:
- Are required or strongly preferred in the job description
- Are NOT demonstrated anywhere in the resume (neither in skills section NOR in experience)
- Would genuinely strengthen the candidate's application if added

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
  "skillsInferredFromExperience": ["<skill inferred from work descriptions>", ...],
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
 * Also suggests adding skills to the Skills section when the candidate has demonstrable experience.
 */
function tailorResume(resumeData, jobDescription, analysisResults) {
  const systemPrompt = `You are an expert resume writer who helps candidates present their authentic experience in the best light for specific roles.

CRITICAL RULES:
1. NEVER invent or fabricate skills, experiences, or achievements
2. Do NOT add fake content - only surface skills the candidate demonstrably has
3. You CAN rephrase bullet points to better highlight relevant aspects
4. You CAN reorder bullet points to put most relevant first
5. You CAN adjust language to mirror the job description's terminology (where truthful)
6. You CAN make the summary/objective more targeted
7. You CAN update the candidate's headline/title to match the target position
8. Maintain a natural, human voice - not robotic or overly formal
9. Keep the candidate's personality and authentic voice

IMPORTANT - SKILL SECTION UPDATES:
You SHOULD suggest adding skills/keywords to the Skills section when:
- The skill is mentioned/required in the job description
- The candidate DEMONSTRABLY has the skill based on their experience

Examples of valid skill additions (not fabrications - making implicit skills explicit):
- If resume mentions "ChatGPT" or "built AI features" → suggest adding "LLM", "AI Tooling", "Generative AI"
- If resume mentions "PostgreSQL" in experience but not in skills → suggest adding "PostgreSQL" to skills
- If resume mentions "deployed to AWS" → suggest adding "AWS" if not already in skills
- If resume mentions "REST API development" → suggest adding "REST APIs" or "API Design"

The goal is to help the candidate's REAL qualifications shine through, not to create a fake persona.

You must respond with valid JSON only, no other text.`;

  const userMessage = `Tailor this resume for the job below. Remember: do NOT add fake content - only rephrase and surface skills the candidate demonstrably has.

CURRENT RESUME:
${formatResumeForAnalysis(resumeData)}

JOB DESCRIPTION:
${jobDescription}

ANALYSIS (for context):
- Match Score: ${analysisResults.matchScore}%
- Key gaps: ${analysisResults.gaps?.join(', ') || 'None identified'}
- Keywords from job posting to incorporate: ${analysisResults.keywordsMissing?.join(', ') || 'None'}
- Skills inferred from experience: ${analysisResults.skillsInferredFromExperience?.join(', ') || 'None identified'}

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
  "skillsToAdd": [
    {
      "skill": "<skill/keyword to add>",
      "category": "<suggested category in Skills section, e.g., 'Integration & Services', 'Cloud & DevOps'>",
      "justification": "<what in the resume proves the candidate has this skill>"
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

IMPORTANT for skillsToAdd:
- Only suggest skills the candidate demonstrably has based on their experience
- Look for related technologies: ChatGPT implies LLM/AI, "deployed to AWS" implies AWS, etc.
- Check if skills mentioned in experience are missing from the Skills section
- Match category names to the resume's existing skill categories if possible

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
