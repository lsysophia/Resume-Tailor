/**
 * Resume Tailor - Google Docs Add-on
 * Makes a copy of your master resume and tailors it to match job descriptions.
 */

// ==================== Menu & Triggers ====================

/**
 * Creates the add-on menu when the document opens.
 */
function onOpen() {
  DocumentApp.getUi()
    .createAddonMenu()
    .addItem('Open Resume Tailor', 'showSidebar')
    .addSeparator()
    .addSubMenu(DocumentApp.getUi().createMenu('Set Template')
      .addItem('Use Current Document', 'setCurrentDocAsTemplate')
      .addItem('Select from Drive...', 'showTemplatePickerDialog'))
    .addItem('AI Settings', 'showApiKeyDialog')
    .addItem('Help', 'showHelp')
    .addToUi();
}

/**
 * Runs when the add-on is installed.
 */
function onInstall() {
  onOpen();
}

/**
 * Homepage trigger for the add-on.
 */
function onHomepage() {
  return createHomepageCard();
}

// ==================== Sidebar ====================

/**
 * Shows the main sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setTitle('Resume Tailor')
    .setWidth(400);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Shows the AI settings dialog.
 */
function showApiKeyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ApiKeyDialog')
    .setWidth(450)
    .setHeight(480);
  DocumentApp.getUi().showModalDialog(html, 'AI Settings');
}

/**
 * Shows the template picker dialog.
 */
function showTemplatePickerDialog() {
  const html = HtmlService.createHtmlOutputFromFile('TemplatePicker')
    .setWidth(500)
    .setHeight(450);
  DocumentApp.getUi().showModalDialog(html, 'Select Resume Template');
}

/**
 * Gets recent Google Docs for the template picker.
 * @returns {Array} List of recent docs with id, name, and modifiedDate
 */
function getRecentDocs() {
  const docs = [];
  const files = DriveApp.searchFiles(
    "mimeType = 'application/vnd.google-apps.document' and trashed = false"
  );

  let count = 0;
  const maxDocs = 20; // Limit to 20 most recent

  while (files.hasNext() && count < maxDocs) {
    const file = files.next();
    docs.push({
      id: file.getId(),
      name: file.getName(),
      modifiedDate: file.getLastUpdated().toLocaleDateString(),
      url: file.getUrl()
    });
    count++;
  }

  return docs;
}

/**
 * Searches Google Docs by name.
 * @param {string} query - Search query
 * @returns {Array} List of matching docs
 */
function searchDocs(query) {
  const docs = [];
  const searchQuery = `mimeType = 'application/vnd.google-apps.document' and trashed = false and name contains '${query.replace(/'/g, "\\'")}'`;
  const files = DriveApp.searchFiles(searchQuery);

  let count = 0;
  const maxDocs = 20;

  while (files.hasNext() && count < maxDocs) {
    const file = files.next();
    docs.push({
      id: file.getId(),
      name: file.getName(),
      modifiedDate: file.getLastUpdated().toLocaleDateString(),
      url: file.getUrl()
    });
    count++;
  }

  return docs;
}

/**
 * Sets a specific document as the template by ID.
 * @param {string} docId - The document ID to set as template
 * @returns {Object} Result with success status
 */
function setDocAsTemplate(docId) {
  try {
    const file = DriveApp.getFileById(docId);
    const docName = file.getName();

    const userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('TEMPLATE_RESUME_ID', docId);
    userProperties.setProperty('TEMPLATE_RESUME_NAME', docName);

    return {
      success: true,
      docId: docId,
      docName: docName,
      message: `"${docName}" is now set as your template.`
    };
  } catch (e) {
    return {
      success: false,
      error: 'Could not access the selected document: ' + e.message
    };
  }
}

/**
 * Shows help information.
 */
function showHelp() {
  const ui = DocumentApp.getUi();
  ui.alert(
    'Resume Tailor Help',
    'How to use Resume Tailor:\n\n' +
    '1. Configure your AI provider in "AI Settings" (Claude, GPT, or Gemini)\n\n' +
    '2. Open your master resume and click "Set This Doc as Template"\n\n' +
    '3. Click "Open Resume Tailor" and paste a job description or URL\n\n' +
    '4. Click "Analyze" to see your match score and gaps\n\n' +
    '5. Click "Create Tailored Copy" to generate suggestions\n\n' +
    '6. Review changes in the document and accept/reject each one\n\n' +
    'Note: Your original resume is NEVER modified - we always work on a copy!',
    ui.ButtonSet.OK
  );
}

// ==================== API Key Management ====================

/**
 * Saves the Claude API key to user properties (legacy function).
 * New code should use saveProviderApiKey() from Claude.gs
 */
function saveApiKey(apiKey) {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('CLAUDE_API_KEY', apiKey);
  userProperties.setProperty('API_KEY_CLAUDE', apiKey);
  return { success: true, message: 'API key saved successfully!' };
}

/**
 * Gets the Claude API key from user properties (legacy function).
 * New code should use getProviderApiKey() from Claude.gs
 */
function getApiKey() {
  const userProperties = PropertiesService.getUserProperties();
  return userProperties.getProperty('CLAUDE_API_KEY') || '';
}

/**
 * Checks if API key is configured for the current AI provider.
 */
function hasApiKey() {
  const provider = getAIProvider();
  return !!getProviderApiKey(provider);
}

// ==================== Template Resume Management ====================

/**
 * Sets the current document as the template resume.
 */
function setCurrentDocAsTemplate() {
  const doc = DocumentApp.getActiveDocument();
  const docId = doc.getId();
  const docName = doc.getName();

  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('TEMPLATE_RESUME_ID', docId);
  userProperties.setProperty('TEMPLATE_RESUME_NAME', docName);

  const ui = DocumentApp.getUi();
  ui.alert(
    'Template Set',
    `"${docName}" is now set as your master resume template.\n\n` +
    'When you tailor your resume for a job, a copy will be created and the original will remain unchanged.',
    ui.ButtonSet.OK
  );

  return { success: true, docId: docId, docName: docName };
}

/**
 * Gets the template resume info.
 */
function getTemplateInfo() {
  const userProperties = PropertiesService.getUserProperties();
  const templateId = userProperties.getProperty('TEMPLATE_RESUME_ID');
  const templateName = userProperties.getProperty('TEMPLATE_RESUME_NAME');

  if (!templateId) {
    return { hasTemplate: false };
  }

  // Verify the template still exists and is accessible
  try {
    const file = DriveApp.getFileById(templateId);
    return {
      hasTemplate: true,
      templateId: templateId,
      templateName: file.getName() // Get current name in case it was renamed
    };
  } catch (e) {
    // Template no longer accessible
    userProperties.deleteProperty('TEMPLATE_RESUME_ID');
    userProperties.deleteProperty('TEMPLATE_RESUME_NAME');
    return { hasTemplate: false };
  }
}

/**
 * Checks if template is set.
 */
function hasTemplate() {
  return getTemplateInfo().hasTemplate;
}

// ==================== Resume Data ====================

/**
 * Reads resume data from a Google Doc.
 * Parses the document structure to identify sections and content.
 */
function getResumeData(docId) {
  // If no docId provided, use the template
  if (!docId) {
    const templateInfo = getTemplateInfo();
    if (!templateInfo.hasTemplate) {
      throw new Error('No template resume set. Please set a template first.');
    }
    docId = templateInfo.templateId;
  }

  const doc = DocumentApp.openById(docId);
  const body = doc.getBody();

  const resume = {
    docId: docId,
    docName: doc.getName(),
    sections: {},
    elements: [] // Store element details for later modification
  };

  let currentSection = 'Header';
  resume.sections[currentSection] = [];

  const numChildren = body.getNumChildren();

  for (let i = 0; i < numChildren; i++) {
    const element = body.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();

      // Check if this is a section header
      if (text && (
        heading !== DocumentApp.ParagraphHeading.NORMAL ||
        text === text.toUpperCase() ||
        isSectionHeader(text)
      )) {
        currentSection = normalizeSectionName(text);
        if (!resume.sections[currentSection]) {
          resume.sections[currentSection] = [];
        }
      } else if (text) {
        // Regular content
        resume.sections[currentSection].push({
          type: 'paragraph',
          index: i,
          content: text
        });
      }

      resume.elements.push({
        type: 'paragraph',
        index: i,
        section: currentSection,
        content: text,
        heading: heading.toString()
      });

    } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
      const listItem = element.asListItem();
      const text = listItem.getText().trim();

      if (text) {
        resume.sections[currentSection].push({
          type: 'listItem',
          index: i,
          content: text,
          nestingLevel: listItem.getNestingLevel()
        });
      }

      resume.elements.push({
        type: 'listItem',
        index: i,
        section: currentSection,
        content: text,
        nestingLevel: listItem.getNestingLevel()
      });

    } else if (elementType === DocumentApp.ElementType.TABLE) {
      const table = element.asTable();
      const tableContent = extractTableContent(table);

      resume.sections[currentSection].push({
        type: 'table',
        index: i,
        content: tableContent
      });

      resume.elements.push({
        type: 'table',
        index: i,
        section: currentSection,
        content: tableContent
      });
    }
  }

  return resume;
}

/**
 * Checks if text looks like a section header.
 */
function isSectionHeader(text) {
  const normalizedText = text.trim().toUpperCase();

  // Common section header keywords
  const sectionKeywords = [
    'SUMMARY', 'OBJECTIVE', 'PROFILE', 'ABOUT',
    'EXPERIENCE', 'EMPLOYMENT', 'WORK HISTORY', 'CAREER',
    'EDUCATION', 'ACADEMIC', 'QUALIFICATIONS',
    'SKILLS', 'COMPETENCIES', 'EXPERTISE', 'TECHNOLOGIES',
    'PROJECTS', 'PORTFOLIO',
    'CERTIFICATIONS', 'CERTIFICATES', 'LICENSES', 'CREDENTIALS',
    'AWARDS', 'HONORS', 'ACHIEVEMENTS', 'ACCOMPLISHMENTS',
    'PUBLICATIONS', 'PRESENTATIONS', 'PAPERS',
    'VOLUNTEER', 'COMMUNITY', 'LEADERSHIP',
    'LANGUAGES',
    'INTERESTS', 'HOBBIES', 'ACTIVITIES',
    'REFERENCES'
  ];

  // Check if the text contains any section keyword
  return sectionKeywords.some(keyword => normalizedText.includes(keyword));
}

/**
 * Normalizes section names for consistency.
 */
function normalizeSectionName(text) {
  return text.replace(/[:\-]/g, '').trim();
}

/**
 * Extracts content from a table.
 */
function extractTableContent(table) {
  const rows = [];
  const numRows = table.getNumRows();

  for (let r = 0; r < numRows; r++) {
    const row = table.getRow(r);
    const cells = [];
    const numCells = row.getNumCells();

    for (let c = 0; c < numCells; c++) {
      cells.push(row.getCell(c).getText().trim());
    }
    rows.push(cells.join(' | '));
  }

  return rows.join('\n');
}

// ==================== Resume Copy & Modification ====================

/**
 * Creates a copy of the template resume for tailoring.
 * @param {string} candidateName - The candidate's name from their resume
 * @param {string} companyName - The company name from the job posting
 * @param {string} positionTitle - The position title from the job posting
 * @returns {Object} Info about the created copy
 */
function createResumeCopy(candidateName, companyName, positionTitle) {
  const templateInfo = getTemplateInfo();

  if (!templateInfo.hasTemplate) {
    throw new Error('No template resume set. Please set a template first.');
  }

  const templateFile = DriveApp.getFileById(templateInfo.templateId);
  const parentFolder = templateFile.getParents().hasNext()
    ? templateFile.getParents().next()
    : DriveApp.getRootFolder();

  // Clean up the name components
  const cleanName = (candidateName || 'Resume').replace(/[^a-zA-Z0-9\s\-]/g, '').trim().substring(0, 30);
  const cleanCompany = (companyName || 'Company').replace(/[^a-zA-Z0-9\s\-]/g, '').trim().substring(0, 30);
  const cleanPosition = (positionTitle || 'Position').replace(/[^a-zA-Z0-9\s\-]/g, '').trim().substring(0, 30);

  // Format: Name - Company - Position
  const copyName = `${cleanName} - ${cleanCompany} - ${cleanPosition}`;

  const copy = templateFile.makeCopy(copyName, parentFolder);

  return {
    success: true,
    copyId: copy.getId(),
    copyName: copy.getName(),
    copyUrl: copy.getUrl()
  };
}

/**
 * Extracts the candidate's name from the resume.
 * Checks: 1) Document header section, 2) First heading in body, 3) First line of body text.
 */
function getCandidateName(docId) {
  try {
    const doc = DocumentApp.openById(docId);

    // First: Check the document header section (many resumes put name in header)
    const header = doc.getHeader();
    if (header) {
      const headerName = extractNameFromSection(header);
      if (headerName) {
        return headerName;
      }
    }

    // Second: Check the body content
    const body = doc.getBody();
    const bodyName = extractNameFromSection(body);
    if (bodyName) {
      return bodyName;
    }

    return null;
  } catch (e) {
    console.error('Error getting candidate name:', e);
    return null;
  }
}

/**
 * Extracts a name from a document section (header or body).
 * @param {HeaderSection|Body} section - The section to search
 * @returns {string|null} The extracted name or null
 */
function extractNameFromSection(section) {
  // Approach 1: Use getParagraphs() to get all paragraphs directly
  // This is more reliable than iterating through children
  try {
    const paragraphs = section.getParagraphs();

    // First pass: Look for headings (names are often styled as headings)
    for (let i = 0; i < Math.min(paragraphs.length, 5); i++) {
      const para = paragraphs[i];
      const text = para.getText().trim();
      const heading = para.getHeading();

      if (text && heading !== DocumentApp.ParagraphHeading.NORMAL) {
        if (!isSectionHeader(text) && text.length < 50) {
          return text;
        }
      }
    }

    // Second pass: Look for any paragraph that looks like a name
    for (let i = 0; i < Math.min(paragraphs.length, 10); i++) {
      const text = paragraphs[i].getText().trim();
      if (text && text.length > 1 && text.length < 50 && !isSectionHeader(text)) {
        // Skip if it looks like contact info (email, phone, address, URL)
        if (!text.includes('@') &&
            !text.match(/^\+?\d[\d\s\-()]+$/) &&
            !text.match(/^\d+\s/) &&
            !text.match(/^https?:\/\//) &&
            !text.toLowerCase().includes('linkedin') &&
            !text.toLowerCase().includes('github')) {
          return text;
        }
      }
    }
  } catch (e) {
    // getParagraphs() might not be available, fall back to child iteration
    console.log('getParagraphs failed, trying child iteration:', e);
  }

  // Approach 2: Fall back to getText() and parse lines
  try {
    const allText = section.getText();
    if (allText) {
      const lines = allText.split('\n').map(l => l.trim()).filter(l => l.length > 0);
      for (const line of lines.slice(0, 10)) {
        if (line.length > 1 && line.length < 50 && !isSectionHeader(line)) {
          // Skip contact info
          if (!line.includes('@') &&
              !line.match(/^\+?\d[\d\s\-()]+$/) &&
              !line.match(/^\d+\s/) &&
              !line.match(/^https?:\/\//) &&
              !line.toLowerCase().includes('linkedin') &&
              !line.toLowerCase().includes('github')) {
            return line;
          }
        }
      }
    }
  } catch (e) {
    console.log('getText failed:', e);
  }

  // Approach 3: Original child iteration method
  const numChildren = section.getNumChildren();

  // First pass: Look for the first heading (names are often styled as headings)
  for (let i = 0; i < Math.min(numChildren, 5); i++) {
    const element = section.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const para = element.asParagraph();
      const text = para.getText().trim();
      const heading = para.getHeading();

      // If it's a heading and NOT a section header like "EXPERIENCE", it's likely the name
      if (text && heading !== DocumentApp.ParagraphHeading.NORMAL) {
        if (!isSectionHeader(text) && text.length < 50) {
          return text;
        }
      }
    }
  }

  // Second pass: Look for the first non-empty, non-section-header text
  for (let i = 0; i < Math.min(numChildren, 10); i++) {
    const element = section.getChild(i);
    const elementType = element.getType();

    if (elementType === DocumentApp.ElementType.PARAGRAPH) {
      const text = element.asParagraph().getText().trim();
      // Skip empty lines, section headers, and text that looks like contact info
      if (text && text.length > 1 && text.length < 50 && !isSectionHeader(text)) {
        // Skip if it looks like an email, phone, or address
        if (!text.includes('@') && !text.match(/^\+?\d[\d\s\-()]+$/) && !text.match(/^\d+\s/)) {
          return text;
        }
      }
    } else if (elementType === DocumentApp.ElementType.TABLE) {
      // Name might be in a table (common resume format)
      const table = element.asTable();
      if (table.getNumRows() > 0 && table.getRow(0).getNumCells() > 0) {
        const firstCellText = table.getRow(0).getCell(0).getText().trim();
        // Check if first cell looks like a name (not too long, not a section header)
        if (firstCellText && firstCellText.length > 1 && firstCellText.length < 50 &&
            !isSectionHeader(firstCellText) && !firstCellText.includes('@')) {
          return firstCellText;
        }
      }
    }
  }

  return null;
}

/**
 * Applies a single change to a document, preserving original formatting.
 * @param {string} docId - The document ID to modify
 * @param {string} originalText - The text to find and replace
 * @param {string} newText - The replacement text
 */
function applyChangeToDoc(docId, originalText, newText) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Try exact replacement first
    const found = body.findText(escapeRegex(originalText));

    if (found) {
      const element = found.getElement();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();

      // Get the parent paragraph or list item
      const parent = element.getParent();

      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH ||
          parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        const text = parent.editAsText();

        // Capture formatting from the start of the original text
        const formatting = captureFormatting(text, start);

        // Replace text
        text.deleteText(start, end);
        text.insertText(start, newText);

        // Apply captured formatting to the new text
        applyFormatting(text, start, start + newText.length - 1, formatting);
      }

      doc.saveAndClose();
      return { success: true, message: 'Change applied successfully' };
    }

    // If exact match not found, try to find the element containing similar text
    const normalizedOriginal = originalText.toLowerCase().replace(/\s+/g, ' ').trim();
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const elementType = element.getType();

      if (elementType === DocumentApp.ElementType.PARAGRAPH ||
          elementType === DocumentApp.ElementType.LIST_ITEM) {
        const textElement = element.asText();
        const text = textElement.getText();
        const normalizedText = text.toLowerCase().replace(/\s+/g, ' ').trim();

        // Check for substantial overlap
        if (normalizedText.includes(normalizedOriginal.substring(0, 30)) ||
            normalizedOriginal.includes(normalizedText.substring(0, 30))) {

          // Capture formatting from the start of the element
          const formatting = captureFormatting(textElement, 0);

          // Replace text
          textElement.setText(newText);

          // Apply captured formatting to the entire new text
          applyFormatting(textElement, 0, newText.length - 1, formatting);

          doc.saveAndClose();
          return { success: true, message: 'Change applied successfully' };
        }
      }
    }

    doc.saveAndClose();
    return {
      success: false,
      message: 'Could not find the original text to replace. You may need to apply this change manually.'
    };

  } catch (error) {
    console.error('Error applying change:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Captures text formatting attributes at a given position.
 * @param {Text} textElement - The text element to get formatting from
 * @param {number} offset - The character offset to sample formatting from
 * @returns {Object} Formatting attributes
 */
function captureFormatting(textElement, offset) {
  return {
    fontFamily: textElement.getFontFamily(offset),
    fontSize: textElement.getFontSize(offset),
    bold: textElement.isBold(offset),
    italic: textElement.isItalic(offset),
    underline: textElement.isUnderline(offset),
    strikethrough: textElement.isStrikethrough(offset),
    foregroundColor: textElement.getForegroundColor(offset)
  };
}

/**
 * Applies formatting attributes to a range of text.
 * @param {Text} textElement - The text element to format
 * @param {number} start - Start offset
 * @param {number} end - End offset (inclusive)
 * @param {Object} formatting - Formatting attributes to apply
 */
function applyFormatting(textElement, start, end, formatting) {
  if (end < start) return;

  if (formatting.fontFamily) {
    textElement.setFontFamily(start, end, formatting.fontFamily);
  }
  if (formatting.fontSize) {
    textElement.setFontSize(start, end, formatting.fontSize);
  }
  if (formatting.bold !== null) {
    textElement.setBold(start, end, formatting.bold);
  }
  if (formatting.italic !== null) {
    textElement.setItalic(start, end, formatting.italic);
  }
  if (formatting.underline !== null) {
    textElement.setUnderline(start, end, formatting.underline);
  }
  if (formatting.strikethrough !== null) {
    textElement.setStrikethrough(start, end, formatting.strikethrough);
  }
  if (formatting.foregroundColor) {
    textElement.setForegroundColor(start, end, formatting.foregroundColor);
  }
}

/**
 * Escapes special regex characters in a string.
 */
function escapeRegex(string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Applies a suggested change visually to a document (strikethrough old, green new).
 * The accept/reject buttons in the sidebar will finalize or revert these changes.
 *
 * @param {string} docId - The document ID to modify
 * @param {string} originalText - The text to find
 * @param {string} newText - The suggested replacement
 * @param {string} reason - Why this change is suggested
 * @returns {Object} Result with success status
 */
function applySuggestedChangeVisual(docId, originalText, newText, reason) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Try to find the exact text
    const found = body.findText(escapeRegex(originalText));

    if (found) {
      const element = found.getElement();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();
      const parent = element.getParent();

      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH ||
          parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        const text = parent.editAsText();

        // Capture original formatting
        const formatting = captureFormatting(text, start);

        // Apply strikethrough and red color to original text (marking for deletion)
        text.setStrikethrough(start, end, true);
        text.setForegroundColor(start, end, '#ea4335'); // Red for deletion

        // Insert the new text right after the original
        const insertPos = end + 1;
        text.insertText(insertPos, newText);

        // Apply formatting to new text: green color, preserve other formatting
        const newEnd = insertPos + newText.length - 1;
        if (formatting.fontFamily) text.setFontFamily(insertPos, newEnd, formatting.fontFamily);
        if (formatting.fontSize) text.setFontSize(insertPos, newEnd, formatting.fontSize);
        text.setBold(insertPos, newEnd, formatting.bold || false);
        text.setItalic(insertPos, newEnd, formatting.italic || false);
        text.setForegroundColor(insertPos, newEnd, '#34a853'); // Green for addition
        text.setStrikethrough(insertPos, newEnd, false);
      }

      doc.saveAndClose();
      return { success: true, message: 'Suggestion applied' };
    }

    // Fallback: try to find similar text using fuzzy match
    const normalizedOriginal = originalText.toLowerCase().replace(/\s+/g, ' ').trim();
    const numChildren = body.getNumChildren();

    for (let i = 0; i < numChildren; i++) {
      const element = body.getChild(i);
      const elementType = element.getType();

      if (elementType === DocumentApp.ElementType.PARAGRAPH ||
          elementType === DocumentApp.ElementType.LIST_ITEM) {
        const textElement = element.asText();
        const text = textElement.getText();
        const normalizedText = text.toLowerCase().replace(/\s+/g, ' ').trim();

        if (normalizedText.includes(normalizedOriginal.substring(0, 30)) ||
            normalizedOriginal.includes(normalizedText.substring(0, 30))) {

          const formatting = captureFormatting(textElement, 0);
          const originalLength = text.length;

          textElement.setStrikethrough(0, originalLength - 1, true);
          textElement.setForegroundColor(0, originalLength - 1, '#ea4335');
          textElement.appendText(newText);

          const newStart = originalLength;
          const newEnd = newStart + newText.length - 1;
          if (formatting.fontFamily) textElement.setFontFamily(newStart, newEnd, formatting.fontFamily);
          if (formatting.fontSize) textElement.setFontSize(newStart, newEnd, formatting.fontSize);
          textElement.setBold(newStart, newEnd, formatting.bold || false);
          textElement.setItalic(newStart, newEnd, formatting.italic || false);
          textElement.setForegroundColor(newStart, newEnd, '#34a853');
          textElement.setStrikethrough(newStart, newEnd, false);

          doc.saveAndClose();
          return { success: true, message: 'Suggestion applied visually' };
        }
      }
    }

    doc.saveAndClose();
    return { success: false, message: 'Could not find text to suggest change for.' };

  } catch (error) {
    console.error('Error applying visual suggestion:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Accepts a suggested change (removes strikethrough text, keeps green text).
 * @param {string} docId - The document ID
 * @param {string} originalText - The strikethrough text to remove
 * @param {string} newText - The green text to keep
 */
function acceptSuggestedChange(docId, originalText, newText) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Find the combined text (original with strikethrough + new text)
    const combinedPattern = escapeRegex(originalText + newText);
    const found = body.findText(combinedPattern);

    if (found) {
      const element = found.getElement();
      const start = found.getStartOffset();
      const parent = element.getParent();

      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH ||
          parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        const text = parent.editAsText();

        // Get formatting from the new text
        const newTextStart = start + originalText.length;
        const formatting = captureFormatting(text, newTextStart);

        // Delete the original (strikethrough) text
        text.deleteText(start, start + originalText.length - 1);

        // Reset the color of the remaining (new) text to black
        text.setForegroundColor(start, start + newText.length - 1, null);
      }

      doc.saveAndClose();
      return { success: true, message: 'Change accepted' };
    }

    doc.saveAndClose();
    return { success: false, message: 'Could not find the suggested change to accept.' };

  } catch (error) {
    console.error('Error accepting suggestion:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Rejects a suggested change (removes green text, keeps and restores original).
 * @param {string} docId - The document ID
 * @param {string} originalText - The strikethrough text to restore
 * @param {string} newText - The green text to remove
 */
function rejectSuggestedChange(docId, originalText, newText) {
  try {
    const doc = DocumentApp.openById(docId);
    const body = doc.getBody();

    // Find the combined text
    const combinedPattern = escapeRegex(originalText + newText);
    const found = body.findText(combinedPattern);

    if (found) {
      const element = found.getElement();
      const start = found.getStartOffset();
      const parent = element.getParent();

      if (parent.getType() === DocumentApp.ElementType.PARAGRAPH ||
          parent.getType() === DocumentApp.ElementType.LIST_ITEM) {
        const text = parent.editAsText();

        // Delete the new (green) text
        const newTextStart = start + originalText.length;
        text.deleteText(newTextStart, newTextStart + newText.length - 1);

        // Remove strikethrough and reset color on original text
        text.setStrikethrough(start, start + originalText.length - 1, false);
        text.setForegroundColor(start, start + originalText.length - 1, null);
      }

      doc.saveAndClose();
      return { success: true, message: 'Change rejected' };
    }

    doc.saveAndClose();
    return { success: false, message: 'Could not find the suggested change to reject.' };

  } catch (error) {
    console.error('Error rejecting suggestion:', error);
    return { success: false, message: error.message };
  }
}

/**
 * Applies all suggested changes visually to a document.
 * @param {string} docId - The document ID
 * @param {Array} changes - Array of change objects with original, tailored, reason
 * @returns {Object} Result with counts
 */
function applyAllSuggestionsVisual(docId, changes) {
  let applied = 0;
  let failed = 0;

  for (const change of changes) {
    const result = applySuggestedChangeVisual(docId, change.original, change.tailored, change.reason);
    if (result.success) {
      applied++;
    } else {
      failed++;
    }
  }

  return {
    success: true,
    applied: applied,
    failed: failed,
    message: `Applied ${applied} of ${changes.length} suggestions to the document.`
  };
}

// ==================== Card UI (for Add-on homepage) ====================

function createHomepageCard() {
  const card = CardService.newCardBuilder();

  card.setHeader(CardService.newCardHeader()
    .setTitle('Resume Tailor')
    .setSubtitle('Customize your resume for any job'));

  const section = CardService.newCardSection();

  section.addWidget(CardService.newTextParagraph()
    .setText('Analyze job descriptions and create tailored copies of your resume - while keeping your authentic voice.'));

  section.addWidget(CardService.newTextButton()
    .setText('Open Resume Tailor')
    .setOnClickAction(CardService.newAction().setFunctionName('showSidebar')));

  card.addSection(section);

  return card.build();
}
