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
