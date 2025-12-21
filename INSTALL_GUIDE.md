# Grok-Outlook Companion - Installation Guide

## Quick Install (For IT Teams)

### System Requirements
- Windows 10 or Windows 11 (64-bit)
- Microsoft Outlook Desktop (Classic version, NOT "New Outlook")
- Internet connection (for Grok AI API)
- 500MB disk space

### Installation Options

#### Option 1: Installer (Recommended)
1. Run `Grok-Outlook Companion-Setup-1.0.0.exe`
2. Follow the installation wizard
3. Choose installation directory (default: `C:\Program Files\Grok-Outlook Companion`)
4. Creates Start Menu and Desktop shortcuts

#### Option 2: Portable Version
1. Copy `Grok-Outlook Companion-Portable-1.0.0.exe` to any folder
2. Double-click to run (no installation required)
3. Good for testing or restricted environments

---

## First-Time Setup

1. **Open Classic Outlook** and select an email
2. **Launch Grok-Outlook Companion**
3. **Configure AI Provider:**
   - Click the ⚙️ Settings icon (top right)
   - Go to "AI Provider" tab
   - Enter your Grok API key (get one at https://console.x.ai)
   - Click "Save Settings"
4. **Test:** Click "Fetch Active Email" to verify connection

---

## Network Requirements

| Endpoint | Port | Purpose |
|----------|------|---------|
| `api.x.ai` | 443 (HTTPS) | Grok AI API |
| `localhost` | 11434 | Ollama local AI (optional) |

**No other external connections required.**

---

## Security Notes

✅ **Local Processing** - Outlook data accessed via local COM API, never sent to external servers (except AI requests)

✅ **Encrypted Credentials** - API keys stored in Windows Credential Manager

✅ **Sandboxed** - Application runs in Chromium sandbox

✅ **No Admin Rights** - Standard user installation supported

---

## Data Storage

| Data | Location |
|------|----------|
| Settings | `%AppData%\grok-outlook-companion\` |
| Logs | `%AppData%\grok-outlook-companion\logs\` |
| API Keys | Windows Credential Manager |

---

## Troubleshooting

### "Outlook Disconnected"
- Ensure Classic Outlook is open (not New Outlook)
- Have an email selected
- Restart the application

### App won't start
- Check Windows Event Viewer for errors
- Try running as Administrator once
- Verify .NET Framework 4.7.2+ is installed

### API errors
- Verify API key is correct
- Check internet connection
- Confirm `api.x.ai` is not blocked by firewall

---

## Uninstallation

**Installer version:** Use Windows "Add/Remove Programs"

**Portable version:** Just delete the .exe file

**User data cleanup:** Delete `%AppData%\grok-outlook-companion\` folder

---

## Support

For issues, contact the application administrator or check the GitHub repository.

**Version:** 1.0.0  
**Build Date:** December 2024

