# Resume Tailor - Google Docs Add-on

A Google Docs add-on that uses Claude AI to analyze job descriptions and create tailored copies of your resume - while keeping your authentic voice.

## Features

- **Match Analysis**: Analyzes how well your resume matches a job description (0-100% score)
- **Honest Assessment**: Identifies strengths, gaps, and missing keywords
- **Smart Tailoring**: Only activates for 80%+ matches (jobs worth pursuing)
- **Non-Destructive**: Creates a COPY of your resume - your original stays unchanged
- **Authentic Voice**: Rephrases and reorganizes content - never fabricates skills or experience
- **One-Click Apply**: Review and apply changes directly to the tailored copy

## How It Works

1. Open your master resume Google Doc
2. Set it as your template (Add-ons > Resume Tailor > Set This Doc as Template)
3. Paste a job description into the sidebar
4. Get an instant match score and analysis
5. If 80%+ match, click "Create Tailored Copy"
6. Review and apply suggested changes to your new copy

## Setup Instructions

### Option 1: Manual Setup (Easiest)

1. **Create or open your resume** in Google Docs

2. **Open the Script Editor**:
   - Go to `Extensions > Apps Script`

3. **Copy the code files**:
   - Delete the default `Code.gs` content
   - Copy the contents of each `.gs` file from this repo into new files in the script editor
   - Create HTML files for `Sidebar.html` and `ApiKeyDialog.html`

4. **Update the manifest**:
   - Click on `Project Settings` (gear icon)
   - Check "Show 'appsscript.json' manifest file"
   - Replace `appsscript.json` contents with the one from this repo

5. **Save and refresh** your Google Doc

6. **Configure API Key**:
   - Go to `Add-ons > Resume Tailor > Set API Key`
   - Enter your Claude API key from [console.anthropic.com](https://console.anthropic.com)

7. **Set your template**:
   - Go to `Add-ons > Resume Tailor > Set This Doc as Template`

### Option 2: Using clasp (For Developers)

1. **Install clasp** (Google's CLI for Apps Script):
   ```bash
   npm install -g @google/clasp
   clasp login
   ```

2. **Create a new Apps Script project**:
   ```bash
   clasp create --type docs --title "Resume Tailor"
   ```

3. **Copy the `.clasp.json`** file to this directory (contains your script ID)

4. **Push the code**:
   ```bash
   clasp push
   ```

5. **Open the document**:
   ```bash
   clasp open --addon
   ```

## Resume Format

The add-on works with standard resume formatting in Google Docs:

**Tips:**
- Use headings or ALL CAPS for section headers (e.g., "EXPERIENCE" or "Experience")
- Use bullet points for achievements and responsibilities
- Structure your resume with clear sections like Summary, Experience, Skills, Education
- The add-on is flexible with formatting and will auto-detect sections

**Example sections it recognizes:**
- Summary / Objective / Profile
- Experience / Work Experience / Employment
- Skills / Technical Skills / Core Competencies
- Education
- Projects
- Certifications
- Awards / Achievements

## API Key

You need a Claude API key from Anthropic:

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Create an account or sign in
3. Generate an API key
4. Enter it in the add-on via `Set API Key`

Your API key is stored securely in your Google account's user properties and never shared.

## Privacy & Data

- Your original resume is NEVER modified
- A new copy is created for each tailored version
- Job descriptions are sent to Claude API for analysis
- No data is stored on external servers
- API key is stored in your Google account only

## The 80% Rule

The add-on only offers to tailor your resume when the match score is 80% or higher. This is intentional:

- **Below 80%**: The job may not be a good fit. Consider other opportunities.
- **80% and above**: You're a strong candidate! Tailoring can help you stand out.

This prevents you from wasting time on jobs where you'd need to fabricate qualifications.

## What It Won't Do

- Add fake skills or experiences
- Invent achievements you don't have
- Create a generic "one-size-fits-all" resume
- Make you sound like a robot
- Modify your original resume

## What It Will Do

- Create a copy of your resume for each job application
- Rephrase your bullet points to match job keywords (where truthful)
- Reorder content to highlight most relevant experience
- Adjust your summary/objective to target the role
- Use industry terminology from the job posting
- Keep your authentic voice and personality

## License

MIT
