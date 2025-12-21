# Grok-Outlook Companion

A Windows Electron desktop app that integrates with Microsoft Outlook via COM API to provide AI-powered email assistance using Grok, Ollama, or custom open model APIs.

## Features

- **Outlook Integration**: Access Inbox and Sent Items locally via COM (no Microsoft Graph required)
- **AI Toggle**: Switch between Grok API, local Ollama, and custom open APIs
- **Email Tools**: Summarize emails, draft replies, extract insights
- **System Tray**: Quick access and notifications
- **Hotkeys**: Keyboard shortcuts for seamless workflow

## Tech Stack

- **Electron** - Desktop application framework
- **React** - UI library
- **Material-UI** - Component library
- **electron-edge-js** - COM interop for Outlook access
- **keytar** - Secure credential storage

## Prerequisites

- Windows 10/11
- Node.js 18+ and npm
- Microsoft Outlook desktop application (installed and configured)
- Visual Studio Build Tools (for native modules)

## Installation

1. Clone the repository:
   ```bash
   git clone <your-repo-url>
   cd grok-outlook-prototype
   ```

2. Install dependencies:
   ```bash
   npm install
   ```

3. Rebuild native modules:
   ```bash
   npm run rebuild
   ```

## Development

Start the app in development mode:
```bash
npm run dev
```

This will start the React development server and launch Electron.

## Building

Create a production build:
```bash
npm run build
```

The packaged application will be in the `dist/` folder.

## Project Structure

```
├── main.js           # Electron main process
├── preload.js        # Preload script for IPC
├── public/
│   └── index.html    # HTML template
├── src/
│   ├── index.jsx     # React entry point
│   ├── index.css     # Global styles
│   └── App.jsx       # Main React component
├── assets/           # Icons and images
└── package.json      # Dependencies and scripts
```

## Development Steps

- [x] Step 1: Basic Electron + React setup
- [ ] Step 2: COM integration for Outlook
- [ ] Step 3: Full React UI with Settings
- [ ] Step 4: AI processing integration
- [ ] Step 5: Hotkeys and notifications
- [ ] Step 6: Security and packaging

## License

MIT

