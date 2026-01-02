# 🚀 Grok-Outlook Companion v2.0

**AI-Powered Email Assistant for Microsoft Outlook**

A Windows desktop application that integrates with Microsoft Outlook to provide intelligent email processing using Grok AI (xAI), with support for local AI models via Ollama.

![Version](https://img.shields.io/badge/version-2.0.0-orange)
![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)
![License](https://img.shields.io/badge/license-MIT-green)

---

## 🆕 What's New in v2.0

### 🧠 Smart Shot Mode
One-click workflow that automatically:
1. Fetches the active email from Outlook
2. Extracts and analyzes all attachments (PDF, Word, Excel, images)
3. Runs AI with your selected prompts + attachment context
4. Opens the reply directly in Outlook

### 📬 Recent Inbox Browser
Full-featured inbox browser with:
- **Smart Filters** - Unread only, has attachments, high priority
- **Domain Chips** - Filter by sender domain (e.g., "@company.com")
- **Multi-Select** - Select multiple emails for batch operations
- **Search** - Full-text search across subject, sender, and body preview
- **Preview Pane** - Quick view without leaving the inbox

### 👤 Contact Card & AI Research
- View Outlook contact details with one click
- See email history stats (sent/received count)
- **Deep Research** - AI-powered web research about any contact
- Save research directly to Outlook contact notes

### ✏️ Quick Notes
Add one-time instructions to any AI action without creating a permanent prompt. Perfect for:
- "Keep it brief"
- "Mention I'm traveling next week"
- "Use a formal tone"

---

## ✨ Features

### Core Functionality
- **📧 Outlook Integration** - Direct access to Outlook via local COM API (no cloud/Graph API required)
- **🤖 AI Processing** - Powered by Grok-4 (xAI) with support for Ollama local models
- **📝 Smart Drafting** - AI-generated email replies pushed directly to Outlook
- **📊 Email Analysis** - Summarize, extract insights, and analyze email threads
- **📁 File Analysis** - Analyze attachments and uploaded files (PDF, Word, images, CSV, Excel)
- **🔍 OCR Support** - Extract text from scanned documents and images

### Productivity Tools
- **⚡ One-Shot Mode** - Fetch email, process with AI, and open reply in one click
- **🧠 Smart Shot Mode** - One-Shot + automatic attachment analysis
- **📋 Prompt Library** - Save, organize, and reuse custom AI prompts
- **🏢 Context Profiles** - Personal and company context for tailored responses
- **✏️ Quick Notes** - One-time instructions for immediate customization
- **👤 Contact Cards** - View contact info and AI research
- **📬 Inbox Browser** - Filter, search, and manage recent emails
- **⌨️ Global Hotkeys** - Keyboard shortcuts that work system-wide

### Security & Privacy
- **🔒 Local Processing** - Outlook data accessed locally, never stored externally
- **🔐 Secure Storage** - API keys stored in Windows Credential Manager
- **🛡️ Sandboxed** - Full Chromium sandbox and Content Security Policy
- **📊 Opt-in Telemetry** - Crash reporting only when you enable it

---

## 📋 System Requirements

| Requirement | Details |
|-------------|---------|
| **Operating System** | Windows 10 or Windows 11 |
| **Outlook** | Microsoft Outlook Desktop (Classic) - must be installed and running |
| **API Key** | Grok API key from [xAI Console](https://console.x.ai) |
| **RAM** | 4GB minimum, 8GB recommended |
| **Disk Space** | ~500MB for installation |

> ⚠️ **Important**: The "New Outlook" app is NOT supported. You must use Classic Outlook.

---

## 🚀 Quick Start

### For End Users (Packaged App)

1. **Run the installer** or extract the portable version
2. **Launch** Grok-Outlook Companion
3. **Open Classic Outlook** with an email selected
4. **Configure Settings**:
   - Click the ⚙️ gear icon
   - Go to "AI Provider" tab
   - Enter your Grok API key
   - Click "Save Settings"
5. **Start using**: Click "Fetch Active Email" to pull in the selected email

### For Developers

```bash
# Clone the repository
git clone https://github.com/EarlTheDuke/grok-outlook-companion.git
cd grok-outlook-companion

# Install dependencies
npm install

# Rebuild native modules
npm run rebuild

# Start in development mode
npm run dev
```

---

## ⌨️ Keyboard Shortcuts

| Shortcut | Action |
|----------|--------|
| `Ctrl+Shift+O` | Show/Hide window |
| `Ctrl+Shift+G` | Fetch active email |
| `Ctrl+Shift+S` | Quick summarize |
| `Ctrl+Shift+D` | Quick draft reply |

---

## 📖 How to Use

### Basic Workflow

1. **Select an email** in Outlook
2. **Click "Fetch Active Email"** in the app
3. **Select prompts** from the Active Prompts panel
4. **Add Quick Notes** (optional) for one-time instructions
5. **Click "Run AI"** to process
6. **Click "Open in Outlook"** to create a reply with the AI response

### ⚡ One-Shot Mode

Enable the ⚡ **One Shot Mode** toggle for a streamlined workflow:
1. Select your prompts
2. Click "⚡ ONE SHOT"
3. The app automatically fetches the email, processes with AI, and opens the reply in Outlook

### 🧠 Smart Shot Mode (NEW in v2.0)

For emails with attachments, use 🧠 **Smart Shot Mode**:
1. Enable Smart Shot (automatically disables One Shot)
2. Select your prompts
3. Click "🧠 SMART SHOT"
4. The app:
   - Fetches the active email
   - Extracts all attachments
   - Analyzes each attachment (PDF, Word, Excel, images via OCR)
   - Includes attachment summaries in the AI context
   - Opens the final reply in Outlook

### 📬 Recent Inbox Browser (NEW in v2.0)

Go to the **Recent Inbox** tab to:
1. **Load Emails** - Choose date range (24h to 30 days)
2. **Filter** - Toggle unread, attachments, or high priority
3. **Search** - Find emails by subject, sender, or content
4. **Domain Filter** - Click domain chips to filter by company
5. **Multi-Select** - Check multiple emails for batch AI processing
6. **Preview** - Click any email to see details in the preview pane

### 👤 Contact Cards & Research (NEW in v2.0)

When viewing an email, click the **👤 person icon** to:
- View contact info from Outlook
- See email history (sent/received count)
- Run **Deep Research** - AI searches the web for information about the person
- Save research to Outlook contact notes

### Custom Prompts

1. Go to the **Prompts** tab
2. Click **"+ New Prompt"**
3. Enter name, description, and template
4. Use placeholders like `{email_body}`, `{subject}`, `{sender}`
5. Save and use in your AI actions

### Context Profiles

In **Settings**:
- **About You** - Your name, role, communication style preferences
- **Company** - Detailed company information for context-aware responses

---

## 🔧 Configuration

### AI Provider Settings

| Provider | Setup |
|----------|-------|
| **Grok (xAI)** | Enter API key from [console.x.ai](https://console.x.ai) |
| **Ollama** | Install Ollama locally, runs on `localhost:11434` |
| **Custom** | Configure any OpenAI-compatible API endpoint |

### Data Storage Locations

| Data | Location |
|------|----------|
| Settings | `%AppData%/grok-outlook-companion/ai-settings.json` |
| Prompts | `%AppData%/grok-outlook-companion/prompts.json` |
| API Keys | Windows Credential Manager (encrypted) |
| Logs | `%AppData%/grok-outlook-companion/logs/` |

---

## 🛡️ Security Features

- ✅ **Context Isolation** - Renderer process isolated from Node.js
- ✅ **Sandbox Enabled** - Full Chromium process sandboxing
- ✅ **CSP Headers** - Content Security Policy prevents XSS
- ✅ **No Remote Code** - All code runs locally
- ✅ **Secure Credentials** - API keys in Windows Credential Manager
- ✅ **Rate Limiting** - Prevents API abuse
- ✅ **Navigation Blocked** - Cannot navigate to external URLs

### Network Access

The app only connects to:
- `api.x.ai` - Grok AI API (HTTPS)
- `localhost:11434` - Ollama (if using local AI)

---

## 🐛 Troubleshooting

### "Outlook Disconnected" Error

1. Make sure **Classic Outlook** (not New Outlook) is open
2. Have an email selected/open in Outlook
3. Try clicking "Fetch Active Email" again

### "Classic Outlook Not Found"

1. Open Outlook desktop app
2. If you see "New Outlook" toggle in top-right, turn it **OFF**
3. Restart Outlook and try again

### Blank/Black Screen

1. Press `Ctrl+R` to refresh
2. If persists, close and reopen the app
3. Check logs at `%AppData%/grok-outlook-companion/logs/`

### AI Not Responding

1. Check your API key in Settings
2. Verify internet connection
3. Check if you've exceeded API rate limits

### Attachments Not Analyzed

1. Ensure the email has attachments (📎 icon visible)
2. Use **Smart Shot Mode** instead of One Shot
3. Supported formats: PDF, Word (.docx), Excel (.xlsx), CSV, images (OCR)

---

## 📁 Project Structure

```
grok-outlook-companion/
├── main.js                 # Electron main process (COM, AI, IPC)
├── preload.js              # Secure bridge to renderer
├── package.json            # Dependencies and scripts
├── public/
│   └── index.html          # HTML template
├── src/
│   ├── index.jsx           # React entry + Error Boundary
│   ├── index.css           # Global styles
│   ├── App.jsx             # Main React component
│   └── components/
│       ├── Settings.jsx        # Settings dialog
│       ├── PromptsTab.jsx      # Prompt library manager
│       ├── PromptSelector.jsx  # Active prompts panel
│       ├── FilesTab.jsx        # File analysis
│       ├── InboxTab.jsx        # Recent inbox browser
│       └── ContactCard.jsx     # Contact profile & research
└── assets/
    └── icon.png            # App icon
```

---

## 📦 Building for Distribution

```bash
# Create production build (installer + portable)
npm run build

# Create portable only
npm run build:portable

# Output will be in dist/ folder
```

---

## 🗺️ Roadmap

### ✅ Completed
- [x] Outlook COM Integration
- [x] Grok AI Processing
- [x] Prompt Library System
- [x] One-Shot Mode
- [x] File Analysis & OCR
- [x] Spell Check
- [x] Crash Reporting
- [x] Recent Inbox Browser (v2.0)
- [x] Smart Shot Mode (v2.0)
- [x] Contact Cards & AI Research (v2.0)
- [x] Quick Notes (v2.0)

### 🚧 Coming Soon
- [ ] Sent Items Browser
- [ ] Multi-email batch AI processing
- [ ] Auto-updater
- [ ] Team/enterprise features

---

## 📄 License

MIT License - See [LICENSE](LICENSE) file for details.

---

## 🙏 Acknowledgments

- [Electron](https://www.electronjs.org/) - Desktop framework
- [React](https://reactjs.org/) - UI library
- [Material-UI](https://mui.com/) - Component library
- [xAI](https://x.ai/) - Grok AI API
- [electron-edge-js](https://github.com/nicedoc/electron-edge-js) - .NET interop
- [Tesseract.js](https://tesseract.projectnaptha.com/) - OCR engine

---

## 📝 Version History

### v2.0.0 (Current)
- 🧠 Smart Shot Mode with automatic attachment analysis
- 📬 Full Recent Inbox browser with filters and search
- 👤 Contact Cards with AI-powered deep research
- ✏️ Quick Notes for one-time instructions
- Improved UI and toggles

### v1.0.0
- Initial release
- One-Shot Mode
- Prompt Library
- File Analysis
- Basic Outlook integration

---

**Built with ❤️ for productivity**
