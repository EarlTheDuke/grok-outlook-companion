const { app, BrowserWindow, Tray, Menu, ipcMain, nativeImage, globalShortcut, Notification, clipboard, dialog } = require('electron');
const path = require('path');
const fs = require('fs');
const axios = require('axios');

// File parsing libraries - load each separately to identify issues
let pdfParse = null;
let mammoth = null;
let csvParse = null;
let Tesseract = null;

// Polyfill browser APIs needed by pdf-parse in Electron
if (typeof globalThis.DOMMatrix === 'undefined') {
  globalThis.DOMMatrix = class DOMMatrix {
    constructor() { this.a = 1; this.b = 0; this.c = 0; this.d = 1; this.e = 0; this.f = 0; }
  };
}
if (typeof globalThis.Path2D === 'undefined') {
  globalThis.Path2D = class Path2D { constructor() {} };
}
if (typeof globalThis.ImageData === 'undefined') {
  globalThis.ImageData = class ImageData { constructor(w, h) { this.width = w; this.height = h; this.data = new Uint8ClampedArray(w * h * 4); } };
}

try {
  pdfParse = require('pdf-parse');
  console.log('✓ pdf-parse loaded');
} catch (e) {
  console.warn('✗ pdf-parse not available:', e.message);
}

try {
  mammoth = require('mammoth');
  console.log('✓ mammoth loaded');
} catch (e) {
  console.warn('✗ mammoth not available:', e.message);
}

try {
  csvParse = require('csv-parse/sync');
  console.log('✓ csv-parse loaded');
} catch (e) {
  console.warn('✗ csv-parse not available:', e.message);
}

try {
  Tesseract = require('tesseract.js');
  console.log('✓ tesseract.js loaded');
} catch (e) {
  console.warn('✗ tesseract.js not available:', e.message);
}

let xlsx = null;
try {
  xlsx = require('xlsx');
  console.log('✓ xlsx loaded');
} catch (e) {
  console.warn('✗ xlsx not available:', e.message);
}

// ============================================================================
// LOGGING SYSTEM
// ============================================================================

const LOG_FILE = path.join(app.getPath('userData'), 'app.log');
const MAX_LOG_SIZE = 5 * 1024 * 1024; // 5MB max log size

function rotateLogIfNeeded() {
  try {
    if (fs.existsSync(LOG_FILE)) {
      const stats = fs.statSync(LOG_FILE);
      if (stats.size > MAX_LOG_SIZE) {
        const backupFile = LOG_FILE.replace('.log', '.old.log');
        if (fs.existsSync(backupFile)) {
          fs.unlinkSync(backupFile);
        }
        fs.renameSync(LOG_FILE, backupFile);
      }
    }
  } catch (e) {
    // Ignore rotation errors
  }
}

function log(level, message, data = null) {
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] [${level.toUpperCase()}] ${message}${data ? ' | ' + JSON.stringify(data) : ''}`;
  
  // Console output
  if (level === 'error') {
    console.error(logMessage);
  } else if (level === 'warn') {
    console.warn(logMessage);
  } else {
    console.log(logMessage);
  }
  
  // File output
  try {
    fs.appendFileSync(LOG_FILE, logMessage + '\n');
  } catch (e) {
    // Ignore file write errors
  }
}

// SECURITY: Sanitize sensitive data before logging
function sanitizeForLog(data) {
  if (!data) return data;
  
  const sensitiveKeys = ['apiKey', 'password', 'token', 'secret', 'body', 'bodyHtml', 'HTMLBody'];
  
  if (typeof data === 'object') {
    const sanitized = { ...data };
    for (const key of Object.keys(sanitized)) {
      if (sensitiveKeys.some(sk => key.toLowerCase().includes(sk.toLowerCase()))) {
        sanitized[key] = sanitized[key] ? '[REDACTED]' : null;
      }
    }
    return sanitized;
  }
  
  return data;
}

const logger = {
  info: (msg, data) => log('info', msg, sanitizeForLog(data)),
  warn: (msg, data) => log('warn', msg, sanitizeForLog(data)),
  error: (msg, data) => log('error', msg, sanitizeForLog(data)),
  debug: (msg, data) => log('debug', msg, sanitizeForLog(data)),
};

// Rotate log on startup
rotateLogIfNeeded();

// ============================================================================
// SECURITY UTILITIES
// ============================================================================

// SECURITY: Input validation helpers
function validateString(value, maxLength = 10000) {
  if (typeof value !== 'string') return false;
  if (value.length > maxLength) return false;
  return true;
}

function sanitizeString(value, maxLength = 10000) {
  if (typeof value !== 'string') return '';
  return value.slice(0, maxLength);
}

// SECURITY: Simple rate limiter for AI calls
const rateLimiter = {
  calls: new Map(),
  maxCalls: 20,        // Max calls per window
  windowMs: 60000,     // 1 minute window
  
  check(key) {
    const now = Date.now();
    const windowStart = now - this.windowMs;
    
    // Get call timestamps for this key
    let timestamps = this.calls.get(key) || [];
    
    // Filter out old timestamps
    timestamps = timestamps.filter(t => t > windowStart);
    
    // Check if over limit
    if (timestamps.length >= this.maxCalls) {
      logger.warn('Rate limit exceeded', { key, calls: timestamps.length });
      return false;
    }
    
    // Add current timestamp
    timestamps.push(now);
    this.calls.set(key, timestamps);
    return true;
  }
};

// Keytar for secure credential storage
let keytar;
try {
  keytar = require('keytar');
} catch (e) {
  console.warn('Keytar not available, using fallback storage');
  keytar = null;
}

// Fix for black screen - disable hardware acceleration
app.disableHardwareAcceleration();

// ============================================================================
// AI SETTINGS AND PROCESSING
// ============================================================================

const SERVICE_NAME = 'GrokOutlookCompanion';
let aiSettings = {
  provider: 'grok',
  model: 'grok-4',
  endpoint: 'https://api.x.ai/v1/chat/completions',
  hasApiKey: false,
  globalContext: {
    enabled: true,
    name: '',
    role: '',
    company: '',
    industry: '',
    communicationStyle: '', // empty = none
    detailLevel: '', // empty = none
    customNotes: '',
  },
  companyContext: {
    enabled: false,
    content: '',
  },
};

// Load settings from file
function loadSettings() {
  try {
    const settingsPath = path.join(app.getPath('userData'), 'ai-settings.json');
    if (fs.existsSync(settingsPath)) {
      const data = JSON.parse(fs.readFileSync(settingsPath, 'utf8'));
      aiSettings = { ...aiSettings, ...data };
      console.log('Loaded AI settings:', aiSettings.provider, aiSettings.model);
    }
  } catch (error) {
    console.error('Error loading settings:', error);
  }
}

// Save settings to file
function saveSettings(settings) {
  try {
    const settingsPath = path.join(app.getPath('userData'), 'ai-settings.json');
    // Don't save API key to file - use keytar
    const { apiKey, ...settingsToSave } = settings;
    fs.writeFileSync(settingsPath, JSON.stringify(settingsToSave, null, 2));
    aiSettings = { ...aiSettings, ...settingsToSave };
    console.log('Saved AI settings');
  } catch (error) {
    console.error('Error saving settings:', error);
    throw error;
  }
}

// Get API key from secure storage
async function getApiKey(provider) {
  if (keytar) {
    try {
      return await keytar.getPassword(SERVICE_NAME, provider);
    } catch (error) {
      console.error('Error getting API key:', error);
      return null;
    }
  }
  return null;
}

// Save API key to secure storage
async function saveApiKey(provider, key) {
  if (keytar) {
    try {
      await keytar.setPassword(SERVICE_NAME, provider, key);
      return true;
    } catch (error) {
      console.error('Error saving API key:', error);
      return false;
    }
  }
  return false;
}

// Process with AI - Grok/OpenAI compatible API
async function processWithGrokAPI(apiKey, model, endpoint, messages) {
  try {
    const response = await axios.post(
      endpoint,
      {
        model: model,
        messages: messages,
        temperature: 0.7,
        max_tokens: 2000,
      },
      {
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${apiKey}`,
        },
        timeout: 120000, // 2 minutes
      }
    );

    if (response.data && response.data.choices && response.data.choices[0]) {
      return {
        success: true,
        content: response.data.choices[0].message.content,
        usage: response.data.usage,
      };
    }
    
    return {
      success: false,
      error: 'Invalid response from API',
    };
  } catch (error) {
    console.error('Grok API error:', error.response?.data || error.message);
    return {
      success: false,
      error: error.response?.data?.error?.message || error.message,
    };
  }
}

// Process with Ollama (local)
async function processWithOllama(model, endpoint, messages) {
  try {
    // Convert messages to Ollama format
    const response = await axios.post(
      endpoint,
      {
        model: model,
        messages: messages,
        stream: false,
      },
      {
        headers: {
          'Content-Type': 'application/json',
        },
        timeout: 120000, // Longer timeout for local models
      }
    );

    if (response.data && response.data.message) {
      return {
        success: true,
        content: response.data.message.content,
      };
    }
    
    return {
      success: false,
      error: 'Invalid response from Ollama',
    };
  } catch (error) {
    console.error('Ollama error:', error.message);
    return {
      success: false,
      error: error.code === 'ECONNREFUSED' 
        ? 'Cannot connect to Ollama. Make sure Ollama is running (ollama serve).'
        : error.message,
    };
  }
}

// Build system prompt with global context and company context
function buildSystemPrompt() {
  const ctx = aiSettings.globalContext || {};
  const company = aiSettings.companyContext || {};
  
  const contextParts = [];
  
  // Add company context first (if enabled and has content)
  if (company.enabled && company.content && company.content.trim()) {
    contextParts.push(company.content.trim());
    logger.info('Including company context', { chars: company.content.length });
  }
  
  // Check if global context is enabled
  if (ctx.enabled !== false) {
    // Add user context if available
    if (ctx.name || ctx.role || ctx.company) {
      let userInfo = 'You are helping';
      if (ctx.name) userInfo += ` ${ctx.name}`;
      if (ctx.role) userInfo += `, ${ctx.role}`;
      if (ctx.company) userInfo += ` at ${ctx.company}`;
      if (ctx.industry) userInfo += ` (${ctx.industry} industry)`;
      userInfo += '.';
      contextParts.push(userInfo);
    }
    
    // Add communication preferences (only if set)
    const hasStyle = ctx.communicationStyle && ctx.communicationStyle.trim();
    const hasDetail = ctx.detailLevel && ctx.detailLevel.trim();
    if (hasStyle || hasDetail) {
      let style = 'Communication preferences:';
      if (hasStyle) style += ` ${ctx.communicationStyle} tone`;
      if (hasDetail) style += `${hasStyle ? ',' : ''} ${ctx.detailLevel} responses`;
      style += '.';
      contextParts.push(style);
    }
    
    // Add custom notes
    if (ctx.customNotes && ctx.customNotes.trim()) {
      contextParts.push(`Additional notes:\n${ctx.customNotes}`);
    }
  }
  
  // Build final system prompt
  let systemPrompt;
  if (contextParts.length > 0) {
    systemPrompt = contextParts.join('\n\n');
    // Only add the helper text if we don't have company context (which has its own instructions)
    if (!company.enabled || !company.content) {
      systemPrompt += '\n\nHelp the user understand and respond to their emails efficiently.';
    }
  } else {
    systemPrompt = 'You are a helpful email assistant.';
  }
  
  logger.info('Built system prompt', { 
    hasUserContext: ctx.enabled !== false, 
    hasCompanyContext: company.enabled && !!company.content,
    totalLength: systemPrompt.length 
  });
  
  return systemPrompt;
}

// Clean AI output - remove markdown formatting that could mess up emails
function cleanAiOutput(text) {
  if (!text) return text;
  
  let cleaned = text;
  
  // Remove angle bracket URL formatting: <http://url> or <https://url> → url
  cleaned = cleaned.replace(/<(https?:\/\/[^>]+)>/g, '$1');
  
  // Remove markdown links: [text](url) → text
  cleaned = cleaned.replace(/\[([^\]]+)\]\([^)]+\)/g, '$1');
  
  // Remove markdown bold: **text** → text
  cleaned = cleaned.replace(/\*\*([^*]+)\*\*/g, '$1');
  
  // Remove markdown italic: *text* → text (but not bullet points)
  cleaned = cleaned.replace(/(?<!\n)\*([^*\n]+)\*(?!\*)/g, '$1');
  
  // Remove markdown code blocks: `text` → text
  cleaned = cleaned.replace(/`([^`]+)`/g, '$1');
  
  // Clean up any double spaces
  cleaned = cleaned.replace(/  +/g, ' ');
  
  // Clean up excessive newlines (more than 2 in a row)
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
  
  // Trim and remove ALL trailing newlines/whitespace
  cleaned = cleaned.trim();
  cleaned = cleaned.replace(/[\n\r\s]+$/g, '');
  
  return cleaned;
}

// Main AI processing function
async function processWithAI(prompt, emailData) {
  const provider = aiSettings.provider;
  const model = aiSettings.model;
  const endpoint = aiSettings.endpoint;

  // Build the messages array with global context
  const systemPrompt = buildSystemPrompt();
  
  let userContent = prompt;
  if (emailData) {
    userContent = `${prompt}\n\n--- EMAIL ---\nSubject: ${emailData.subject}\nFrom: ${emailData.senderName} <${emailData.senderEmail}>\nTo: ${emailData.to}\nDate: ${emailData.receivedTime}\n\n${emailData.body}\n--- END EMAIL ---`;
  }

  const messages = [
    { role: 'system', content: systemPrompt },
    { role: 'user', content: userContent },
  ];

  if (provider === 'ollama') {
    return await processWithOllama(model, endpoint, messages);
  } else {
    // Grok, OpenAI, Open WebUI, or custom - all use the same OpenAI-compatible API format
    const apiKey = await getApiKey(provider);
    if (!apiKey) {
      return {
        success: false,
        error: provider === 'openwebui' 
          ? 'API key not configured. Get your key from Open WebUI → Profile → Settings → Account → API Keys'
          : 'API key not configured. Please set your API key in Settings.',
      };
    }
    return await processWithGrokAPI(apiKey, model, endpoint, messages);
  }
}

// Additional flags to fix rendering issues
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');

// Keep references to prevent garbage collection
let mainWindow = null;
let tray = null;

// Check if running in development
const isDev = !app.isPackaged;

// ============================================================================
// OUTLOOK COM INTEGRATION via electron-edge-js
// ============================================================================

let edge;
let outlookFunctions = null;

// Find the Outlook interop assembly path
function findOutlookInteropPath() {
  const possiblePaths = [
    // Office 365 / Office 2019/2021
    'C:\\Program Files\\Microsoft Office\\root\\Office16\\ADDINS\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\ADDINS\\Microsoft.Office.Interop.Outlook.dll',
    // GAC paths
    'C:\\Windows\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Outlook\\15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Windows\\Microsoft.NET\\assembly\\GAC_MSIL\\Microsoft.Office.Interop.Outlook\\v4.0_15.0.0.0__71e9bce111e9429c\\Microsoft.Office.Interop.Outlook.dll',
    // Office 2016
    'C:\\Program Files\\Microsoft Office\\Office16\\ADDINS\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Program Files (x86)\\Microsoft Office\\Office16\\ADDINS\\Microsoft.Office.Interop.Outlook.dll',
    // Visual Studio Primary Interop Assemblies
    'C:\\Program Files\\Microsoft Visual Studio\\Shared\\Visual Studio Tools for Office\\PIA\\Office15\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Program Files\\Microsoft Visual Studio\\Shared\\Visual Studio Tools for Office\\PIA\\Office16\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Program Files (x86)\\Microsoft Visual Studio\\Shared\\Visual Studio Tools for Office\\PIA\\Office15\\Microsoft.Office.Interop.Outlook.dll',
    'C:\\Program Files (x86)\\Microsoft Visual Studio\\Shared\\Visual Studio Tools for Office\\PIA\\Office16\\Microsoft.Office.Interop.Outlook.dll',
  ];

  for (const p of possiblePaths) {
    if (fs.existsSync(p)) {
      console.log('Found Outlook interop at:', p);
      return p;
    }
  }
  
  return null;
}

// Initialize edge-js and Outlook COM functions using late-binding (no DLL required)
async function initializeOutlookCOM() {
  try {
    edge = require('electron-edge-js');
    console.log('electron-edge-js loaded successfully');
    console.log('edge.func type:', typeof edge.func);
    
    // Test if edge-js is working with a simple function
    const testFunc = edge.func(`
      async (input) => {
        return "Edge.js is working!";
      }
    `);
    
    console.log('testFunc type:', typeof testFunc);
    
    // Call with callback style
    const testResult = await new Promise((resolve, reject) => {
      testFunc({}, (error, result) => {
        if (error) {
          console.error('Edge.js test error:', error);
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    console.log('Edge.js test result:', testResult);
    
    // Use late-binding COM approach (doesn't require interop DLL)
    outlookFunctions = {
      getActiveEmail: createGetActiveEmailFunction(),
      getRecentInbox: createGetRecentEmailsFunction('Inbox'),
      getRecentSent: createGetRecentEmailsFunction('SentMail'),
      checkOutlookRunning: createCheckOutlookFunction(),
    };
    
    console.log('Outlook COM functions initialized (late-binding mode)');
    return true;
  } catch (error) {
    console.error('Failed to initialize electron-edge-js:', error.message);
    console.error('Full error:', error);
    console.error('Make sure to run: npm run rebuild');
    return false;
  }
}

// Check if Outlook is running using late-binding
function createCheckOutlookFunction() {
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        try
        {
          // Try to get a running Outlook instance directly
          try
          {
            object outlookApp = Marshal.GetActiveObject("Outlook.Application");
            if (outlookApp != null)
            {
              Marshal.ReleaseComObject(outlookApp);
              return new {
                success = true,
                running = true,
                message = "Outlook is running"
              };
            }
          }
          catch (COMException ex)
          {
            // Check if it's just that Outlook isn't running vs not installed
            if (ex.HResult == -2147221021) // MK_E_UNAVAILABLE - not running
            {
              return new {
                success = true,
                running = false,
                message = "Outlook is not running. Please open the Classic Outlook desktop app."
              };
            }
            // Try another method - see if Outlook.Application is registered
            Type outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType != null)
            {
              return new {
                success = true,
                running = false,
                message = "Outlook is installed but not running. Please open the Classic Outlook desktop app."
              };
            }
            return new {
              success = false,
              running = false,
              message = "Classic Outlook is not installed. The 'New Outlook' app does not support COM automation. Please install or switch to Classic Outlook."
            };
          }

          return new {
            success = true,
            running = false,
            message = "Outlook is not running"
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            running = false,
            message = "Error: " + ex.Message + " (You may need Classic Outlook, not the New Outlook app)"
          };
        }
      }
    }
  `);
}

// Get the currently active/selected email from Outlook using late-binding
function createGetActiveEmailFunction() {
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object explorer = null;
        object selection = null;
        object mailItem = null;

        try
        {
          // Get running Outlook instance
          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              message = "Outlook is not running. Please start Outlook first.",
              data = (object)null
            };
          }

          // Get ActiveExplorer
          Type outlookType = outlookApp.GetType();
          explorer = outlookType.InvokeMember("ActiveExplorer", 
            BindingFlags.GetProperty, null, outlookApp, null);
          
          if (explorer == null)
          {
            return new {
              success = false,
              message = "No active Outlook window found.",
              data = (object)null
            };
          }

          // Get Selection
          Type explorerType = explorer.GetType();
          selection = explorerType.InvokeMember("Selection", 
            BindingFlags.GetProperty, null, explorer, null);

          if (selection == null)
          {
            return new {
              success = false,
              message = "No selection available.",
              data = (object)null
            };
          }

          // Get Count
          Type selType = selection.GetType();
          int count = (int)selType.InvokeMember("Count", 
            BindingFlags.GetProperty, null, selection, null);

          if (count == 0)
          {
            return new {
              success = false,
              message = "No email selected in Outlook. Please select an email.",
              data = (object)null
            };
          }

          // Get first item (index 1 in COM)
          mailItem = selType.InvokeMember("Item", 
            BindingFlags.InvokeMethod, null, selection, new object[] { 1 });

          if (mailItem == null)
          {
            return new {
              success = false,
              message = "Could not get selected item.",
              data = (object)null
            };
          }

          // Check if it's a MailItem (Class = 43)
          Type itemType = mailItem.GetType();
          int itemClass = 0;
          try
          {
            itemClass = (int)itemType.InvokeMember("Class", 
              BindingFlags.GetProperty, null, mailItem, null);
          }
          catch { }

          if (itemClass != 43) // olMail = 43
          {
            return new {
              success = false,
              message = "Selected item is not an email (might be a meeting, contact, etc.)",
              data = (object)null
            };
          }

          // Get email properties
          string subject = SafeGetProperty(itemType, mailItem, "Subject", "");
          string senderName = SafeGetProperty(itemType, mailItem, "SenderName", "");
          string senderEmail = SafeGetProperty(itemType, mailItem, "SenderEmailAddress", "");
          
          // Try to get SMTP address if senderEmail is in Exchange format (starts with /O=)
          if (senderEmail.StartsWith("/O=") || senderEmail.StartsWith("/o="))
          {
            try
            {
              // Try to get Sender object and resolve SMTP address
              object sender = itemType.InvokeMember("Sender", BindingFlags.GetProperty, null, mailItem, null);
              if (sender != null)
              {
                Type senderType = sender.GetType();
                try
                {
                  // Try GetExchangeUser
                  object exchUser = senderType.InvokeMember("GetExchangeUser", BindingFlags.InvokeMethod, null, sender, null);
                  if (exchUser != null)
                  {
                    Type exchType = exchUser.GetType();
                    string smtpAddr = (string)exchType.InvokeMember("PrimarySmtpAddress", BindingFlags.GetProperty, null, exchUser, null);
                    if (!string.IsNullOrEmpty(smtpAddr))
                    {
                      senderEmail = smtpAddr;
                    }
                    Marshal.ReleaseComObject(exchUser);
                  }
                }
                catch { }
                Marshal.ReleaseComObject(sender);
              }
            }
            catch { }
          }
          
          DateTime receivedTime = SafeGetDateProperty(itemType, mailItem, "ReceivedTime");
          string body = SafeGetProperty(itemType, mailItem, "Body", "");
          string bodyHtml = SafeGetProperty(itemType, mailItem, "HTMLBody", "");
          string to = SafeGetProperty(itemType, mailItem, "To", "");
          string cc = SafeGetProperty(itemType, mailItem, "CC", "");
          string entryId = SafeGetProperty(itemType, mailItem, "EntryID", "");
          string conversationId = SafeGetProperty(itemType, mailItem, "ConversationID", "");
          bool unRead = SafeGetBoolProperty(itemType, mailItem, "UnRead", false);
          string importance = SafeGetProperty(itemType, mailItem, "Importance", "1");

          // Get attachment count
          int attachmentCount = 0;
          try
          {
            object attachments = itemType.InvokeMember("Attachments", 
              BindingFlags.GetProperty, null, mailItem, null);
            if (attachments != null)
            {
              attachmentCount = (int)attachments.GetType().InvokeMember("Count", 
                BindingFlags.GetProperty, null, attachments, null);
              Marshal.ReleaseComObject(attachments);
            }
          }
          catch { }

          var emailData = new {
            subject = subject,
            senderName = senderName,
            senderEmail = senderEmail,
            receivedTime = receivedTime.ToString("o"),
            body = body,
            bodyHtml = bodyHtml,
            to = to,
            cc = cc,
            hasAttachments = attachmentCount > 0,
            attachmentCount = attachmentCount,
            importance = importance,
            isRead = !unRead,
            entryId = entryId,
            conversationId = conversationId
          };

          return new {
            success = true,
            message = "Email retrieved successfully",
            data = emailData
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            message = "Error: " + ex.Message,
            data = (object)null
          };
        }
        finally
        {
          if (mailItem != null) Marshal.ReleaseComObject(mailItem);
          if (selection != null) Marshal.ReleaseComObject(selection);
          if (explorer != null) Marshal.ReleaseComObject(explorer);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }

      private static string SafeGetProperty(Type type, object obj, string propName, string defaultVal)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          if (val == null) return defaultVal;
          return val.ToString();
        }
        catch { return defaultVal; }
      }

      private static DateTime SafeGetDateProperty(Type type, object obj, string propName)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          return val != null ? (DateTime)val : DateTime.MinValue;
        }
        catch { return DateTime.MinValue; }
      }

      private static bool SafeGetBoolProperty(Type type, object obj, string propName, bool defaultVal)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          return val != null ? (bool)val : defaultVal;
        }
        catch { return defaultVal; }
      }
    }
  `);
}

// Create function to get recent emails from specified folder using late-binding
function createGetRecentEmailsFunction(folderType) {
  return edge.func({
    source: `
      using System;
      using System.Collections.Generic;
      using System.Threading.Tasks;
      using System.Runtime.InteropServices;
      using System.Reflection;

      public class Startup
      {
        public async Task<object> Invoke(dynamic input)
        {
          object outlookApp = null;
          object ns = null;
          object folder = null;
          object items = null;
          object filteredItems = null;

          try
          {
            int maxItems = (int)(input.maxItems ?? 50);
            int daysBack = (int)(input.daysBack ?? 365);
            string folderTypeStr = (string)(input.folderType ?? "Inbox");
            
            // Get running Outlook instance
            try
            {
              outlookApp = Marshal.GetActiveObject("Outlook.Application");
            }
            catch (COMException)
            {
              return new {
                success = false,
                message = "Outlook is not running. Please start Outlook first.",
                data = new List<object>(),
                count = 0
              };
            }

            // Get MAPI namespace
            Type outlookType = outlookApp.GetType();
            ns = outlookType.InvokeMember("GetNamespace", 
              BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

            // Get folder (olFolderInbox = 6, olFolderSentMail = 5)
            int folderConst = folderTypeStr == "SentMail" ? 5 : 6;
            Type nsType = ns.GetType();
            folder = nsType.InvokeMember("GetDefaultFolder", 
              BindingFlags.InvokeMethod, null, ns, new object[] { folderConst });

            if (folder == null)
            {
              return new {
                success = false,
                message = "Could not access " + folderTypeStr + " folder.",
                data = new List<object>(),
                count = 0
              };
            }

            // Get items
            Type folderType = folder.GetType();
            items = folderType.InvokeMember("Items", 
              BindingFlags.GetProperty, null, folder, null);

            // Sort by date (descending)
            Type itemsType = items.GetType();
            string dateField = folderTypeStr == "SentMail" ? "SentOn" : "ReceivedTime";
            itemsType.InvokeMember("Sort", 
              BindingFlags.InvokeMethod, null, items, new object[] { "[" + dateField + "]", true });

            // Apply date filter
            DateTime filterDate = DateTime.Now.AddDays(-daysBack);
            string filter = "[" + dateField + "] >= '" + filterDate.ToString("MM/dd/yyyy HH:mm") + "'";
            
            filteredItems = itemsType.InvokeMember("Restrict", 
              BindingFlags.InvokeMethod, null, items, new object[] { filter });

            var emails = new List<object>();
            int count = 0;
            Type filteredType = filteredItems.GetType();
            
            // Get total count
            int totalCount = (int)filteredType.InvokeMember("Count", 
              BindingFlags.GetProperty, null, filteredItems, null);

            // Iterate through items
            for (int i = 1; i <= Math.Min(totalCount, maxItems); i++)
            {
              object item = null;
              try
              {
                item = filteredType.InvokeMember("Item", 
                  BindingFlags.InvokeMethod, null, filteredItems, new object[] { i });

                if (item == null) continue;

                Type itemType = item.GetType();
                
                // Check if it's a MailItem (Class = 43)
                int itemClass = 0;
                try
                {
                  itemClass = (int)itemType.InvokeMember("Class", 
                    BindingFlags.GetProperty, null, item, null);
                }
                catch { continue; }

                if (itemClass != 43) continue;

                // Get properties
                string subject = SafeGetProperty(itemType, item, "Subject", "(No Subject)");
                string senderName = SafeGetProperty(itemType, item, "SenderName", "");
                string senderEmail = SafeGetProperty(itemType, item, "SenderEmailAddress", "");
                DateTime receivedTime = SafeGetDateProperty(itemType, item, "ReceivedTime");
                DateTime sentOn = SafeGetDateProperty(itemType, item, "SentOn");
                string body = SafeGetProperty(itemType, item, "Body", "");
                string to = SafeGetProperty(itemType, item, "To", "");
                string cc = SafeGetProperty(itemType, item, "CC", "");
                string entryId = SafeGetProperty(itemType, item, "EntryID", "");
                string conversationId = SafeGetProperty(itemType, item, "ConversationID", "");
                string categories = SafeGetProperty(itemType, item, "Categories", "");
                bool unRead = SafeGetBoolProperty(itemType, item, "UnRead", false);
                string importanceStr = SafeGetProperty(itemType, item, "Importance", "1");
                
                // Parse importance (0=low, 1=normal, 2=high)
                string importanceLevel = "normal";
                if (importanceStr == "0") importanceLevel = "low";
                else if (importanceStr == "2") importanceLevel = "high";

                // Truncate body for preview (clean up whitespace)
                string bodyPreview = body.Replace("\\r\\n", " ").Replace("\\n", " ").Replace("\\r", " ");
                bodyPreview = System.Text.RegularExpressions.Regex.Replace(bodyPreview, @"\\s+", " ").Trim();
                bodyPreview = bodyPreview.Length > 250 ? bodyPreview.Substring(0, 250) + "..." : bodyPreview;

                // Get attachment count
                int attachmentCount = 0;
                try
                {
                  object attachments = itemType.InvokeMember("Attachments", 
                    BindingFlags.GetProperty, null, item, null);
                  if (attachments != null)
                  {
                    attachmentCount = (int)attachments.GetType().InvokeMember("Count", 
                      BindingFlags.GetProperty, null, attachments, null);
                    Marshal.ReleaseComObject(attachments);
                  }
                }
                catch { }
                
                // Extract sender domain for filtering
                string senderDomain = "";
                if (!string.IsNullOrEmpty(senderEmail) && senderEmail.Contains("@"))
                {
                  senderDomain = senderEmail.Substring(senderEmail.LastIndexOf("@") + 1).ToLower();
                }

                emails.Add(new {
                  subject = subject,
                  senderName = senderName,
                  senderEmail = senderEmail,
                  senderDomain = senderDomain,
                  receivedTime = receivedTime.ToString("o"),
                  sentOn = sentOn.ToString("o"),
                  bodyPreview = bodyPreview,
                  to = to,
                  cc = cc,
                  hasAttachments = attachmentCount > 0,
                  attachmentCount = attachmentCount,
                  importance = importanceLevel,
                  isRead = !unRead,
                  entryId = entryId,
                  conversationId = conversationId,
                  categories = categories
                });
                count++;
              }
              catch { }
              finally
              {
                if (item != null) Marshal.ReleaseComObject(item);
              }
            }

            return new {
              success = true,
              message = "Retrieved " + count + " emails from " + folderTypeStr,
              data = emails,
              count = count
            };
          }
          catch (Exception ex)
          {
            return new {
              success = false,
              message = "Error: " + ex.Message,
              data = new List<object>(),
              count = 0
            };
          }
          finally
          {
            if (filteredItems != null) Marshal.ReleaseComObject(filteredItems);
            if (items != null) Marshal.ReleaseComObject(items);
            if (folder != null) Marshal.ReleaseComObject(folder);
            if (ns != null) Marshal.ReleaseComObject(ns);
            if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
          }
        }

        private static string SafeGetProperty(Type type, object obj, string propName, string defaultVal)
        {
          try
          {
            object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
            if (val == null) return defaultVal;
            return val.ToString();
          }
          catch { return defaultVal; }
        }

        private static DateTime SafeGetDateProperty(Type type, object obj, string propName)
        {
          try
          {
            object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
            return val != null ? (DateTime)val : DateTime.MinValue;
          }
          catch { return DateTime.MinValue; }
        }

        private static bool SafeGetBoolProperty(Type type, object obj, string propName, bool defaultVal)
        {
          try
          {
            object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
            return val != null ? (bool)val : defaultVal;
          }
          catch { return defaultVal; }
        }
      }
    `,
    references: ['System.dll']
  });
}

// Create a reply email in Outlook with pre-filled content
function createReplyInOutlook() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object originalItem = null;
        object replyItem = null;

        try
        {
          string entryId = (string)input.entryId;
          string replyBody = (string)input.replyBody;
          bool replyAll = input.replyAll != null ? (bool)input.replyAll : false;
          
          if (string.IsNullOrEmpty(entryId))
          {
            return new {
              success = false,
              message = "Email ID is required"
            };
          }

          // Get running Outlook instance
          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              message = "Outlook is not running. Please start Outlook first."
            };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          // Get the original email
          Type nsType = ns.GetType();
          originalItem = nsType.InvokeMember("GetItemFromID", 
            BindingFlags.InvokeMethod, null, ns, new object[] { entryId });

          if (originalItem == null)
          {
            return new {
              success = false,
              message = "Could not find the original email"
            };
          }

          Type itemType = originalItem.GetType();
          
          // Create reply (Reply = no args, ReplyAll = no args)
          string replyMethod = replyAll ? "ReplyAll" : "Reply";
          replyItem = itemType.InvokeMember(replyMethod, 
            BindingFlags.InvokeMethod, null, originalItem, null);

          if (replyItem == null)
          {
            return new {
              success = false,
              message = "Could not create reply"
            };
          }

          Type replyType = replyItem.GetType();
          
          // Get the existing HTML body (preserves all original formatting)
          string existingHtmlBody = "";
          try
          {
            existingHtmlBody = (string)replyType.InvokeMember("HTMLBody", 
              BindingFlags.GetProperty, null, replyItem, null);
          }
          catch { }

          // Convert AI reply text to HTML (preserve line breaks, escape HTML chars)
          string htmlReply = System.Net.WebUtility.HtmlEncode(replyBody)
            .Replace("\\r\\n", "<br>")
            .Replace("\\n", "<br>")
            .Replace("\\r", "");
          
          // Remove any trailing <br> tags to prevent extra spacing before signature
          while (htmlReply.EndsWith("<br>"))
          {
            htmlReply = htmlReply.Substring(0, htmlReply.Length - 4);
          }
          
          // Wrap AI reply in a span (inline, no block spacing at all)
          string aiHtmlBlock = "<span style=\\"font-family: Calibri, Arial, sans-serif; font-size: 11pt;\\">" + 
            htmlReply + "</span>";
          
          // Find where to insert (after <body> tag if present)
          string newHtmlBody;
          int bodyTagIndex = existingHtmlBody.IndexOf("<body", StringComparison.OrdinalIgnoreCase);
          
          if (bodyTagIndex >= 0)
          {
            // Find the end of the <body> tag
            int bodyTagEnd = existingHtmlBody.IndexOf(">", bodyTagIndex);
            if (bodyTagEnd >= 0)
            {
              // Insert AI reply right after <body...>
              newHtmlBody = existingHtmlBody.Substring(0, bodyTagEnd + 1) + 
                aiHtmlBlock + 
                existingHtmlBody.Substring(bodyTagEnd + 1);
            }
            else
            {
              // Fallback: prepend
              newHtmlBody = aiHtmlBlock + existingHtmlBody;
            }
          }
          else
          {
            // No body tag found, just prepend
            newHtmlBody = aiHtmlBlock + existingHtmlBody;
          }
          
          // Set the HTML body (preserves original message formatting)
          replyType.InvokeMember("HTMLBody", 
            BindingFlags.SetProperty, null, replyItem, new object[] { newHtmlBody });

          // Display the reply (opens it in Outlook for user to review/send)
          // Using false = non-modal, so it doesn't block waiting for user to close
          replyType.InvokeMember("Display", 
            BindingFlags.InvokeMethod, null, replyItem, new object[] { false });

          return new {
            success = true,
            message = "Reply opened in Outlook! Review and click Send when ready."
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            message = "Error: " + ex.Message
          };
        }
        finally
        {
          // Don't release replyItem - user needs it open!
          if (originalItem != null) Marshal.ReleaseComObject(originalItem);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Create a new email in Outlook (for forwarding AI content, etc.)
function createNewEmailInOutlook() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object mailItem = null;

        try
        {
          string to = input.to != null ? (string)input.to : "";
          string subject = input.subject != null ? (string)input.subject : "";
          string body = input.body != null ? (string)input.body : "";
          
          // Get running Outlook instance
          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              message = "Outlook is not running. Please start Outlook first."
            };
          }

          Type outlookType = outlookApp.GetType();
          
          // Create new mail item (olMailItem = 0)
          mailItem = outlookType.InvokeMember("CreateItem", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { 0 });

          if (mailItem == null)
          {
            return new {
              success = false,
              message = "Could not create new email"
            };
          }

          Type mailType = mailItem.GetType();
          
          // Set properties
          if (!string.IsNullOrEmpty(to))
          {
            mailType.InvokeMember("To", 
              BindingFlags.SetProperty, null, mailItem, new object[] { to });
          }
          
          if (!string.IsNullOrEmpty(subject))
          {
            mailType.InvokeMember("Subject", 
              BindingFlags.SetProperty, null, mailItem, new object[] { subject });
          }
          
          mailType.InvokeMember("Body", 
            BindingFlags.SetProperty, null, mailItem, new object[] { body });

          // Display the email (non-modal so it doesn't block)
          mailType.InvokeMember("Display", 
            BindingFlags.InvokeMethod, null, mailItem, new object[] { false });

          return new {
            success = true,
            message = "New email opened in Outlook!"
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            message = "Error: " + ex.Message
          };
        }
        finally
        {
          // Don't release mailItem - user needs it open!
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Get attachments from an email and save them to temp folder
function createGetAttachmentsFunction() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Collections.Generic;
    using System.IO;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object mailItem = null;

        try
        {
          string entryId = (string)input.entryId;
          string tempFolder = (string)input.tempFolder;
          
          if (string.IsNullOrEmpty(entryId))
          {
            return new {
              success = false,
              message = "Email ID is required",
              attachments = new List<object>()
            };
          }

          // Create temp folder if needed
          if (!Directory.Exists(tempFolder))
          {
            Directory.CreateDirectory(tempFolder);
          }

          // Get running Outlook instance
          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              message = "Outlook is not running.",
              attachments = new List<object>()
            };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          Type nsType = ns.GetType();
          mailItem = nsType.InvokeMember("GetItemFromID", 
            BindingFlags.InvokeMethod, null, ns, new object[] { entryId });

          if (mailItem == null)
          {
            return new {
              success = false,
              message = "Email not found",
              attachments = new List<object>()
            };
          }

          Type itemType = mailItem.GetType();
          
          // Get attachments collection
          object attachments = itemType.InvokeMember("Attachments", 
            BindingFlags.GetProperty, null, mailItem, null);
          
          if (attachments == null)
          {
            return new {
              success = true,
              message = "No attachments",
              attachments = new List<object>()
            };
          }

          Type attachType = attachments.GetType();
          int count = (int)attachType.InvokeMember("Count", 
            BindingFlags.GetProperty, null, attachments, null);

          var attachmentList = new List<object>();

          for (int i = 1; i <= count; i++)
          {
            object attachment = null;
            try
            {
              attachment = attachType.InvokeMember("Item", 
                BindingFlags.InvokeMethod, null, attachments, new object[] { i });

              if (attachment == null) continue;

              Type attType = attachment.GetType();
              
              string fileName = (string)attType.InvokeMember("FileName", 
                BindingFlags.GetProperty, null, attachment, null);
              
              long size = 0;
              try {
                size = (long)(int)attType.InvokeMember("Size", 
                  BindingFlags.GetProperty, null, attachment, null);
              } catch { }
              
              int attType2 = 1;
              try {
                attType2 = (int)attType.InvokeMember("Type", 
                  BindingFlags.GetProperty, null, attachment, null);
              } catch { }

              // Only save file attachments (Type 1 = olByValue)
              if (attType2 == 1 && !string.IsNullOrEmpty(fileName))
              {
                // Create unique filename to avoid conflicts
                string safeFileName = Path.GetFileNameWithoutExtension(fileName) + 
                  "_" + Guid.NewGuid().ToString().Substring(0, 8) + 
                  Path.GetExtension(fileName);
                string filePath = Path.Combine(tempFolder, safeFileName);

                // Save attachment
                attType.InvokeMember("SaveAsFile", 
                  BindingFlags.InvokeMethod, null, attachment, new object[] { filePath });

                attachmentList.Add(new {
                  fileName = fileName,
                  savedPath = filePath,
                  size = size,
                  extension = Path.GetExtension(fileName).ToLower()
                });
              }
            }
            catch (Exception ex)
            {
              // Skip problematic attachments
              System.Diagnostics.Debug.WriteLine("Error with attachment: " + ex.Message);
            }
            finally
            {
              if (attachment != null) Marshal.ReleaseComObject(attachment);
            }
          }

          Marshal.ReleaseComObject(attachments);

          return new {
            success = true,
            message = "Attachments extracted: " + attachmentList.Count,
            attachments = attachmentList
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            message = "Error: " + ex.Message,
            attachments = new List<object>()
          };
        }
        finally
        {
          if (mailItem != null) Marshal.ReleaseComObject(mailItem);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Lookup contact by email address
function createContactLookupFunction() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Collections.Generic;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object contactsFolder = null;

        try
        {
          string emailAddress = (string)input.email;
          
          if (string.IsNullOrEmpty(emailAddress))
          {
            return new {
              success = false,
              found = false,
              message = "Email address is required",
              contact = (object)null
            };
          }

          // Get running Outlook instance
          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              found = false,
              message = "Outlook is not running.",
              contact = (object)null
            };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          Type nsType = ns.GetType();
          
          // Get Contacts folder (olFolderContacts = 10)
          contactsFolder = nsType.InvokeMember("GetDefaultFolder", 
            BindingFlags.InvokeMethod, null, ns, new object[] { 10 });

          if (contactsFolder == null)
          {
            return new {
              success = true,
              found = false,
              message = "No contacts folder",
              contact = (object)null
            };
          }

          Type folderType = contactsFolder.GetType();
          object items = folderType.InvokeMember("Items", 
            BindingFlags.GetProperty, null, contactsFolder, null);

          Type itemsType = items.GetType();
          
          // Search for contact by email
          string filter = "[Email1Address] = '" + emailAddress.Replace("'", "''") + "' OR " +
                         "[Email2Address] = '" + emailAddress.Replace("'", "''") + "' OR " +
                         "[Email3Address] = '" + emailAddress.Replace("'", "''") + "'";
          
          object contact = null;
          try
          {
            contact = itemsType.InvokeMember("Find", 
              BindingFlags.InvokeMethod, null, items, new object[] { filter });
          }
          catch { }

          if (contact == null)
          {
            // Try Global Address List
            try
            {
              object addressLists = nsType.InvokeMember("AddressLists", 
                BindingFlags.GetProperty, null, ns, null);
              Type alType = addressLists.GetType();
              int alCount = (int)alType.InvokeMember("Count", 
                BindingFlags.GetProperty, null, addressLists, null);
              
              for (int i = 1; i <= alCount && contact == null; i++)
              {
                object addressList = alType.InvokeMember("Item", 
                  BindingFlags.GetProperty, null, addressLists, new object[] { i });
                Type listType = addressList.GetType();
                
                object entries = listType.InvokeMember("AddressEntries", 
                  BindingFlags.GetProperty, null, addressList, null);
                Type entriesType = entries.GetType();
                int entryCount = (int)entriesType.InvokeMember("Count", 
                  BindingFlags.GetProperty, null, entries, null);
                
                for (int j = 1; j <= entryCount && contact == null; j++)
                {
                  try
                  {
                    object entry = entriesType.InvokeMember("Item", 
                      BindingFlags.GetProperty, null, entries, new object[] { j });
                    Type entryType = entry.GetType();
                    
                    string address = "";
                    try
                    {
                      address = (string)entryType.InvokeMember("Address", 
                        BindingFlags.GetProperty, null, entry, null);
                    }
                    catch { }
                    
                    if (address.ToLower() == emailAddress.ToLower())
                    {
                      // Found in GAL - get what info we can
                      string name = "";
                      string company = "";
                      string title = "";
                      string phone = "";
                      string dept = "";
                      
                      try { name = (string)entryType.InvokeMember("Name", BindingFlags.GetProperty, null, entry, null); } catch { }
                      
                      // Try to get exchange user details
                      try
                      {
                        object exchUser = entryType.InvokeMember("GetExchangeUser", BindingFlags.InvokeMethod, null, entry, null);
                        if (exchUser != null)
                        {
                          Type exchType = exchUser.GetType();
                          try { company = (string)exchType.InvokeMember("CompanyName", BindingFlags.GetProperty, null, exchUser, null); } catch { }
                          try { title = (string)exchType.InvokeMember("JobTitle", BindingFlags.GetProperty, null, exchUser, null); } catch { }
                          try { phone = (string)exchType.InvokeMember("BusinessTelephoneNumber", BindingFlags.GetProperty, null, exchUser, null); } catch { }
                          try { dept = (string)exchType.InvokeMember("Department", BindingFlags.GetProperty, null, exchUser, null); } catch { }
                          Marshal.ReleaseComObject(exchUser);
                        }
                      }
                      catch { }
                      
                      Marshal.ReleaseComObject(entry);
                      Marshal.ReleaseComObject(entries);
                      Marshal.ReleaseComObject(addressList);
                      Marshal.ReleaseComObject(addressLists);
                      
                      return new {
                        success = true,
                        found = true,
                        source = "GAL",
                        message = "Found in Global Address List",
                        contact = new {
                          fullName = name ?? "",
                          email = emailAddress,
                          jobTitle = title ?? "",
                          company = company ?? "",
                          department = dept ?? "",
                          phoneMobile = "",
                          phoneWork = phone ?? "",
                          phoneHome = "",
                          address = "",
                          notes = "",
                          categories = ""
                        }
                      };
                    }
                    
                    Marshal.ReleaseComObject(entry);
                  }
                  catch { }
                }
                
                Marshal.ReleaseComObject(entries);
                Marshal.ReleaseComObject(addressList);
              }
              
              Marshal.ReleaseComObject(addressLists);
            }
            catch { }
            
            Marshal.ReleaseComObject(items);
            
            return new {
              success = true,
              found = false,
              message = "Contact not found",
              contact = (object)null
            };
          }

          // Found contact - extract info
          Type contactType = contact.GetType();
          
          string fullName = "";
          string email1 = "";
          string jobTitle = "";
          string companyName = "";
          string department = "";
          string mobileTel = "";
          string businessTel = "";
          string homeTel = "";
          string businessAddr = "";
          string notes = "";
          string categories = "";
          
          try { fullName = (string)contactType.InvokeMember("FullName", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { email1 = (string)contactType.InvokeMember("Email1Address", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { jobTitle = (string)contactType.InvokeMember("JobTitle", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { companyName = (string)contactType.InvokeMember("CompanyName", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { department = (string)contactType.InvokeMember("Department", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { mobileTel = (string)contactType.InvokeMember("MobileTelephoneNumber", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { businessTel = (string)contactType.InvokeMember("BusinessTelephoneNumber", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { homeTel = (string)contactType.InvokeMember("HomeTelephoneNumber", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { businessAddr = (string)contactType.InvokeMember("BusinessAddress", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { notes = (string)contactType.InvokeMember("Body", BindingFlags.GetProperty, null, contact, null); } catch { }
          try { categories = (string)contactType.InvokeMember("Categories", BindingFlags.GetProperty, null, contact, null); } catch { }
          
          Marshal.ReleaseComObject(contact);
          Marshal.ReleaseComObject(items);
          
          return new {
            success = true,
            found = true,
            source = "Contacts",
            message = "Contact found",
            contact = new {
              fullName = fullName ?? "",
              email = email1 ?? emailAddress,
              jobTitle = jobTitle ?? "",
              company = companyName ?? "",
              department = department ?? "",
              phoneMobile = mobileTel ?? "",
              phoneWork = businessTel ?? "",
              phoneHome = homeTel ?? "",
              address = businessAddr ?? "",
              notes = notes ?? "",
              categories = categories ?? ""
            }
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            found = false,
            message = "Error: " + ex.Message,
            contact = (object)null
          };
        }
        finally
        {
          if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Get email history count with a contact
function createEmailHistoryFunction() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object inbox = null;
        object sent = null;

        try
        {
          string emailAddress = (string)input.email;
          
          if (string.IsNullOrEmpty(emailAddress))
          {
            return new {
              success = false,
              totalCount = 0,
              inboxCount = 0,
              sentCount = 0,
              lastEmailDate = ""
            };
          }

          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              totalCount = 0,
              inboxCount = 0,
              sentCount = 0,
              lastEmailDate = ""
            };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          Type nsType = ns.GetType();
          
          int inboxCount = 0;
          int sentCount = 0;
          DateTime lastDate = DateTime.MinValue;
          
          // Search Inbox
          inbox = nsType.InvokeMember("GetDefaultFolder", 
            BindingFlags.InvokeMethod, null, ns, new object[] { 6 }); // olFolderInbox
          
          Type folderType = inbox.GetType();
          object inboxItems = folderType.InvokeMember("Items", 
            BindingFlags.GetProperty, null, inbox, null);
          
          Type itemsType = inboxItems.GetType();
          string filter = "[SenderEmailAddress] = '" + emailAddress.Replace("'", "''") + "'";
          
          object restricted = itemsType.InvokeMember("Restrict", 
            BindingFlags.InvokeMethod, null, inboxItems, new object[] { filter });
          
          Type restrictedType = restricted.GetType();
          inboxCount = (int)restrictedType.InvokeMember("Count", 
            BindingFlags.GetProperty, null, restricted, null);
          
          // Get last email date from inbox
          if (inboxCount > 0)
          {
            restrictedType.InvokeMember("Sort", 
              BindingFlags.InvokeMethod, null, restricted, new object[] { "[ReceivedTime]", true });
            
            object lastItem = restrictedType.InvokeMember("GetFirst", 
              BindingFlags.InvokeMethod, null, restricted, null);
            if (lastItem != null)
            {
              Type lastType = lastItem.GetType();
              try
              {
                lastDate = (DateTime)lastType.InvokeMember("ReceivedTime", 
                  BindingFlags.GetProperty, null, lastItem, null);
              }
              catch { }
              Marshal.ReleaseComObject(lastItem);
            }
          }
          
          Marshal.ReleaseComObject(restricted);
          Marshal.ReleaseComObject(inboxItems);
          
          // Search Sent Items
          sent = nsType.InvokeMember("GetDefaultFolder", 
            BindingFlags.InvokeMethod, null, ns, new object[] { 5 }); // olFolderSentMail
          
          Type sentFolderType = sent.GetType();
          object sentItems = sentFolderType.InvokeMember("Items", 
            BindingFlags.GetProperty, null, sent, null);
          
          Type sentItemsType = sentItems.GetType();
          
          // For sent items, we need to check recipients
          int sentTotal = (int)sentItemsType.InvokeMember("Count", 
            BindingFlags.GetProperty, null, sentItems, null);
          
          // Simple approach: count last 200 sent items that match
          int checkLimit = Math.Min(200, sentTotal);
          sentItemsType.InvokeMember("Sort", 
            BindingFlags.InvokeMethod, null, sentItems, new object[] { "[SentOn]", true });
          
          for (int i = 1; i <= checkLimit; i++)
          {
            try
            {
              object item = sentItemsType.InvokeMember("Item", 
                BindingFlags.GetProperty, null, sentItems, new object[] { i });
              
              Type itemType = item.GetType();
              object recipients = itemType.InvokeMember("Recipients", 
                BindingFlags.GetProperty, null, item, null);
              
              Type recipType = recipients.GetType();
              int recipCount = (int)recipType.InvokeMember("Count", 
                BindingFlags.GetProperty, null, recipients, null);
              
              bool found = false;
              for (int j = 1; j <= recipCount && !found; j++)
              {
                object recip = recipType.InvokeMember("Item", 
                  BindingFlags.GetProperty, null, recipients, new object[] { j });
                Type rType = recip.GetType();
                
                string addr = "";
                try
                {
                  addr = (string)rType.InvokeMember("Address", 
                    BindingFlags.GetProperty, null, recip, null);
                }
                catch { }
                
                if (addr.ToLower() == emailAddress.ToLower())
                {
                  found = true;
                  sentCount++;
                  
                  DateTime sentDate = DateTime.MinValue;
                  try
                  {
                    sentDate = (DateTime)itemType.InvokeMember("SentOn", 
                      BindingFlags.GetProperty, null, item, null);
                  }
                  catch { }
                  
                  if (sentDate > lastDate)
                  {
                    lastDate = sentDate;
                  }
                }
                
                Marshal.ReleaseComObject(recip);
              }
              
              Marshal.ReleaseComObject(recipients);
              Marshal.ReleaseComObject(item);
            }
            catch { }
          }
          
          Marshal.ReleaseComObject(sentItems);
          
          return new {
            success = true,
            totalCount = inboxCount + sentCount,
            inboxCount = inboxCount,
            sentCount = sentCount,
            lastEmailDate = lastDate != DateTime.MinValue ? lastDate.ToString("yyyy-MM-dd") : ""
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            totalCount = 0,
            inboxCount = 0,
            sentCount = 0,
            lastEmailDate = "",
            error = ex.Message
          };
        }
        finally
        {
          if (sent != null) Marshal.ReleaseComObject(sent);
          if (inbox != null) Marshal.ReleaseComObject(inbox);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Save notes to a contact
function createSaveContactNotesFunction() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object contactsFolder = null;

        try
        {
          string emailAddress = (string)input.email;
          string notes = (string)input.notes;
          bool append = (bool)input.append;
          
          if (string.IsNullOrEmpty(emailAddress))
          {
            return new { success = false, message = "Email address is required" };
          }

          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new { success = false, message = "Outlook is not running." };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          Type nsType = ns.GetType();
          contactsFolder = nsType.InvokeMember("GetDefaultFolder", 
            BindingFlags.InvokeMethod, null, ns, new object[] { 10 });

          Type folderType = contactsFolder.GetType();
          object items = folderType.InvokeMember("Items", 
            BindingFlags.GetProperty, null, contactsFolder, null);

          Type itemsType = items.GetType();
          string filter = "[Email1Address] = '" + emailAddress.Replace("'", "''") + "' OR " +
                         "[Email2Address] = '" + emailAddress.Replace("'", "''") + "' OR " +
                         "[Email3Address] = '" + emailAddress.Replace("'", "''") + "'";
          
          object contact = itemsType.InvokeMember("Find", 
            BindingFlags.InvokeMethod, null, items, new object[] { filter });

          if (contact == null)
          {
            Marshal.ReleaseComObject(items);
            return new { success = false, message = "Contact not found" };
          }

          Type contactType = contact.GetType();
          
          string existingNotes = "";
          try
          {
            existingNotes = (string)contactType.InvokeMember("Body", 
              BindingFlags.GetProperty, null, contact, null) ?? "";
          }
          catch { }
          
          string newNotes = append ? existingNotes + "\\n\\n--- AI Research " + DateTime.Now.ToString("yyyy-MM-dd HH:mm") + " ---\\n" + notes : notes;
          
          contactType.InvokeMember("Body", 
            BindingFlags.SetProperty, null, contact, new object[] { newNotes });
          
          contactType.InvokeMember("Save", 
            BindingFlags.InvokeMethod, null, contact, null);
          
          Marshal.ReleaseComObject(contact);
          Marshal.ReleaseComObject(items);
          
          return new { success = true, message = "Notes saved to contact" };
        }
        catch (Exception ex)
        {
          return new { success = false, message = "Error: " + ex.Message };
        }
        finally
        {
          if (contactsFolder != null) Marshal.ReleaseComObject(contactsFolder);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }
    }
  `);
}

// Get email by EntryID using late-binding
function createGetEmailByIdFunction() {
  if (!edge) return null;
  
  return edge.func(`
    using System;
    using System.Threading.Tasks;
    using System.Runtime.InteropServices;
    using System.Reflection;

    public class Startup
    {
      public async Task<object> Invoke(dynamic input)
      {
        object outlookApp = null;
        object ns = null;
        object mailItem = null;

        try
        {
          string entryId = (string)input.entryId;
          
          if (string.IsNullOrEmpty(entryId))
          {
            return new {
              success = false,
              message = "Entry ID is required",
              data = (object)null
            };
          }

          try
          {
            outlookApp = Marshal.GetActiveObject("Outlook.Application");
          }
          catch (COMException)
          {
            return new {
              success = false,
              message = "Outlook is not running.",
              data = (object)null
            };
          }

          Type outlookType = outlookApp.GetType();
          ns = outlookType.InvokeMember("GetNamespace", 
            BindingFlags.InvokeMethod, null, outlookApp, new object[] { "MAPI" });

          Type nsType = ns.GetType();
          mailItem = nsType.InvokeMember("GetItemFromID", 
            BindingFlags.InvokeMethod, null, ns, new object[] { entryId });

          if (mailItem == null)
          {
            return new {
              success = false,
              message = "Item not found",
              data = (object)null
            };
          }

          Type itemType = mailItem.GetType();
          
          // Check if it's a MailItem
          int itemClass = 0;
          try
          {
            itemClass = (int)itemType.InvokeMember("Class", 
              BindingFlags.GetProperty, null, mailItem, null);
          }
          catch { }

          if (itemClass != 43)
          {
            return new {
              success = false,
              message = "Item is not an email",
              data = (object)null
            };
          }

          string subject = SafeGetProperty(itemType, mailItem, "Subject", "");
          string senderName = SafeGetProperty(itemType, mailItem, "SenderName", "");
          string senderEmail = SafeGetProperty(itemType, mailItem, "SenderEmailAddress", "");
          DateTime receivedTime = SafeGetDateProperty(itemType, mailItem, "ReceivedTime");
          string body = SafeGetProperty(itemType, mailItem, "Body", "");
          string bodyHtml = SafeGetProperty(itemType, mailItem, "HTMLBody", "");
          string to = SafeGetProperty(itemType, mailItem, "To", "");
          string cc = SafeGetProperty(itemType, mailItem, "CC", "");
          string conversationId = SafeGetProperty(itemType, mailItem, "ConversationID", "");
          bool unRead = SafeGetBoolProperty(itemType, mailItem, "UnRead", false);
          string importance = SafeGetProperty(itemType, mailItem, "Importance", "1");

          int attachmentCount = 0;
          try
          {
            object attachments = itemType.InvokeMember("Attachments", 
              BindingFlags.GetProperty, null, mailItem, null);
            if (attachments != null)
            {
              attachmentCount = (int)attachments.GetType().InvokeMember("Count", 
                BindingFlags.GetProperty, null, attachments, null);
              Marshal.ReleaseComObject(attachments);
            }
          }
          catch { }

          var emailData = new {
            subject = subject,
            senderName = senderName,
            senderEmail = senderEmail,
            receivedTime = receivedTime.ToString("o"),
            body = body,
            bodyHtml = bodyHtml,
            to = to,
            cc = cc,
            hasAttachments = attachmentCount > 0,
            attachmentCount = attachmentCount,
            importance = importance,
            isRead = !unRead,
            entryId = entryId,
            conversationId = conversationId
          };

          return new {
            success = true,
            message = "Email retrieved",
            data = emailData
          };
        }
        catch (Exception ex)
        {
          return new {
            success = false,
            message = "Error: " + ex.Message,
            data = (object)null
          };
        }
        finally
        {
          if (mailItem != null) Marshal.ReleaseComObject(mailItem);
          if (ns != null) Marshal.ReleaseComObject(ns);
          if (outlookApp != null) Marshal.ReleaseComObject(outlookApp);
        }
      }

      private static string SafeGetProperty(Type type, object obj, string propName, string defaultVal)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          if (val == null) return defaultVal;
          return val.ToString();
        }
        catch { return defaultVal; }
      }

      private static DateTime SafeGetDateProperty(Type type, object obj, string propName)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          return val != null ? (DateTime)val : DateTime.MinValue;
        }
        catch { return DateTime.MinValue; }
      }

      private static bool SafeGetBoolProperty(Type type, object obj, string propName, bool defaultVal)
      {
        try
        {
          object val = type.InvokeMember(propName, BindingFlags.GetProperty, null, obj, null);
          return val != null ? (bool)val : defaultVal;
        }
        catch { return defaultVal; }
      }
    }
  `);
}

// ============================================================================
// WINDOW AND TRAY SETUP
// ============================================================================

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    icon: path.join(__dirname, 'assets', 'icon.png'),
    webPreferences: {
      nodeIntegration: false,           // SECURITY: Prevent renderer from accessing Node
      contextIsolation: true,           // SECURITY: Isolate preload script from renderer
      preload: path.join(__dirname, 'preload.js'),
      webSecurity: true,                // SECURITY: Enforce same-origin policy
      allowRunningInsecureContent: false, // SECURITY: Block mixed content
      experimentalFeatures: false,      // SECURITY: Disable experimental web features
      enableBlinkFeatures: '',          // SECURITY: Don't enable extra Blink features
      sandbox: true,                    // SECURITY: Enable Chromium sandbox
      spellcheck: true,                 // Enable built-in spell checker
    },
    frame: true,
    show: true,
    backgroundColor: '#0a0a0f'
  });
  
  // SECURITY: Set Content Security Policy
  mainWindow.webContents.session.webRequest.onHeadersReceived((details, callback) => {
    callback({
      responseHeaders: {
        ...details.responseHeaders,
        'Content-Security-Policy': [
          isDev 
            ? "default-src 'self' http://localhost:3000; script-src 'self' 'unsafe-inline' 'unsafe-eval' http://localhost:3000; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; connect-src 'self' https://api.x.ai http://localhost:* ws://localhost:*; img-src 'self' data: https:;"
            : "default-src 'self'; script-src 'self'; style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; font-src 'self' https://fonts.gstatic.com; connect-src 'self' https://api.x.ai http://localhost:11434; img-src 'self' data:;"
        ]
      }
    });
  });
  
  // SECURITY: Prevent navigation to external URLs
  mainWindow.webContents.on('will-navigate', (event, url) => {
    const parsedUrl = new URL(url);
    if (isDev && parsedUrl.hostname === 'localhost') {
      return; // Allow localhost in dev
    }
    if (!url.startsWith('file://')) {
      logger.warn('Blocked navigation to external URL', { url });
      event.preventDefault();
    }
  });
  
  // SECURITY: Prevent new windows from opening
  mainWindow.webContents.setWindowOpenHandler(({ url }) => {
    logger.warn('Blocked attempt to open new window', { url });
    return { action: 'deny' };
  });

  // Load the app
  const startUrl = isDev 
    ? 'http://localhost:3000' 
    : `file://${path.join(__dirname, 'build', 'index.html')}`;
  
  console.log('Loading URL:', startUrl);
  
  mainWindow.loadURL(startUrl).then(() => {
    console.log('Page loaded successfully');
  }).catch((err) => {
    console.error('Failed to load page:', err);
    // Retry after a short delay
    setTimeout(() => {
      console.log('Retrying to load...');
      mainWindow.loadURL(startUrl);
    }, 2000);
  });

  // Handle load failures
  mainWindow.webContents.on('did-fail-load', (event, errorCode, errorDescription) => {
    console.error('Page failed to load:', errorCode, errorDescription);
    if (isDev && errorCode === -102) {
      // Connection refused - React not ready yet, retry
      setTimeout(() => {
        console.log('Retrying connection to React dev server...');
        mainWindow.loadURL(startUrl);
      }, 3000);
    }
  });

  // Log when page finishes loading
  mainWindow.webContents.on('did-finish-load', () => {
    console.log('Page finished loading');
    // DevTools can be opened with F12 if needed
  });

  // Spell check context menu with suggestions
  mainWindow.webContents.on('context-menu', (event, params) => {
    // Only show custom menu if there are spell check suggestions or it's editable
    if (params.isEditable) {
      const menuItems = [];

      // Add spell check suggestions
      if (params.misspelledWord && params.dictionarySuggestions.length > 0) {
        // Add each suggestion as a menu item
        params.dictionarySuggestions.slice(0, 5).forEach((suggestion) => {
          menuItems.push({
            label: suggestion,
            click: () => mainWindow.webContents.replaceMisspelling(suggestion),
          });
        });
        menuItems.push({ type: 'separator' });
        
        // Add to dictionary option
        menuItems.push({
          label: `Add "${params.misspelledWord}" to dictionary`,
          click: () => mainWindow.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord),
        });
        menuItems.push({ type: 'separator' });
      } else if (params.misspelledWord) {
        // Misspelled but no suggestions
        menuItems.push({
          label: 'No suggestions available',
          enabled: false,
        });
        menuItems.push({
          label: `Add "${params.misspelledWord}" to dictionary`,
          click: () => mainWindow.webContents.session.addWordToSpellCheckerDictionary(params.misspelledWord),
        });
        menuItems.push({ type: 'separator' });
      }

      // Standard edit menu items
      menuItems.push(
        { label: 'Cut', role: 'cut', accelerator: 'CmdOrCtrl+X' },
        { label: 'Copy', role: 'copy', accelerator: 'CmdOrCtrl+C' },
        { label: 'Paste', role: 'paste', accelerator: 'CmdOrCtrl+V' },
        { type: 'separator' },
        { label: 'Select All', role: 'selectAll', accelerator: 'CmdOrCtrl+A' }
      );

      const contextMenu = Menu.buildFromTemplate(menuItems);
      contextMenu.popup();
    }
  });

  mainWindow.on('close', (event) => {
    if (!app.isQuitting) {
      event.preventDefault();
      mainWindow.hide();
    }
    return false;
  });

  mainWindow.on('closed', () => {
    mainWindow = null;
  });
}

function createTray() {
  const iconPath = path.join(__dirname, 'assets', 'tray-icon.png');
  
  let trayIcon;
  try {
    trayIcon = nativeImage.createFromPath(iconPath);
    if (trayIcon.isEmpty()) {
      trayIcon = nativeImage.createEmpty();
    }
  } catch (e) {
    trayIcon = nativeImage.createEmpty();
  }

  tray = new Tray(trayIcon.isEmpty() ? 
    nativeImage.createFromDataURL('data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAAdgAAAHYBTnsmCAAAABl0RVh0U29mdHdhcmUAd3d3Lmlua3NjYXBlLm9yZ5vuPBoAAAEgSURBVDiNpZMxTsNAEEWfN7YxkRANDRVdjtAjcAo6OgpuwBE4Ai1H4AQcgYaSBilCIiAR29gz/ytsJ2aSsNKuZndm/s7sLGy5sPHDPnb7B56qCqWADzxR54S2cMfFNoVxQXUEXFO9sE3hHLCOuRqDr5hXA5/fMGrAGaM2cPkRYxI4BQ6BC8bjCM+0gRN6vcF0CjwAB2ybaCR8A04kDiRd2WL9RgNs0BuYGhiPgVOJA0nXwAq4+4TRqAEn9PtDplNgBhwCm2wkaYQuQGMgaYQOoNX6C8YacEqvN5hMgDlwBGyykaQRukCj4UfSCB1Ao/E3jFXglG43m0yAOXAEbNIY3UBcIGmEDqDR+CvGKnBKt5tNJsAcOAI2aYxuIC6QNEI/7vQ3nDWVA2MAAAAASUVORK5CYII=') 
    : trayIcon
  );

  const contextMenu = Menu.buildFromTemplate([
    {
      label: 'Open Grok-Outlook Companion',
      click: () => {
        if (mainWindow) {
          mainWindow.show();
          mainWindow.focus();
        }
      }
    },
    {
      label: 'Fetch Active Email',
      click: () => {
        if (mainWindow) {
          mainWindow.webContents.send('hotkey-active-email');
        }
      }
    },
    { type: 'separator' },
    {
      label: 'Settings',
      click: () => {
        if (mainWindow) {
          mainWindow.show();
          mainWindow.webContents.send('navigate-settings');
        }
      }
    },
    { type: 'separator' },
    {
      label: 'Quit',
      click: () => {
        app.isQuitting = true;
        app.quit();
      }
    }
  ]);

  tray.setToolTip('Grok-Outlook Companion');
  tray.setContextMenu(contextMenu);

  tray.on('double-click', () => {
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();
    }
  });
}

function showNotification(title, body, options = {}) {
  logger.info('Showing notification', { title, body });
  
  // Try Windows native notification first
  if (Notification.isSupported()) {
    const notification = new Notification({
      title: title,
      body: body,
      icon: path.join(__dirname, 'assets', 'icon.png'),
      silent: options.silent || false,
    });
    
    notification.on('click', () => {
      if (mainWindow) {
        mainWindow.show();
        mainWindow.focus();
      }
    });
    
    notification.show();
  }
  
  // Also show tray balloon as backup
  if (tray) {
    tray.displayBalloon({
      title: title,
      content: body,
      iconType: options.iconType || 'info'
    });
  }
}

// ============================================================================
// GLOBAL HOTKEYS
// ============================================================================

function registerGlobalHotkeys() {
  logger.info('Registering global hotkeys...');
  
  // Ctrl+Shift+G - Fetch active email
  const fetchHotkey = globalShortcut.register('Ctrl+Shift+G', () => {
    logger.info('Hotkey triggered: Ctrl+Shift+G (Fetch Active Email)');
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();
      mainWindow.webContents.send('hotkey-fetch-active');
    }
    showNotification('Grok-Outlook', 'Fetching active email...', { silent: true });
  });
  
  if (!fetchHotkey) {
    logger.warn('Failed to register Ctrl+Shift+G hotkey');
  }

  // Ctrl+Shift+S - Quick Summarize (fetches and summarizes in one go)
  const summarizeHotkey = globalShortcut.register('Ctrl+Shift+S', () => {
    logger.info('Hotkey triggered: Ctrl+Shift+S (Quick Summarize)');
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();
      mainWindow.webContents.send('hotkey-quick-summarize');
    }
    showNotification('Grok-Outlook', 'Fetching and summarizing email...', { silent: true });
  });
  
  if (!summarizeHotkey) {
    logger.warn('Failed to register Ctrl+Shift+S hotkey');
  }

  // Ctrl+Shift+D - Quick Draft Reply
  const draftHotkey = globalShortcut.register('Ctrl+Shift+D', () => {
    logger.info('Hotkey triggered: Ctrl+Shift+D (Quick Draft)');
    if (mainWindow) {
      mainWindow.show();
      mainWindow.focus();
      mainWindow.webContents.send('hotkey-quick-draft');
    }
    showNotification('Grok-Outlook', 'Fetching and drafting reply...', { silent: true });
  });
  
  if (!draftHotkey) {
    logger.warn('Failed to register Ctrl+Shift+D hotkey');
  }

  // Ctrl+Shift+O - Show/Hide window
  const toggleHotkey = globalShortcut.register('Ctrl+Shift+O', () => {
    logger.info('Hotkey triggered: Ctrl+Shift+O (Toggle Window)');
    if (mainWindow) {
      if (mainWindow.isVisible()) {
        mainWindow.hide();
      } else {
        mainWindow.show();
        mainWindow.focus();
      }
    }
  });
  
  if (!toggleHotkey) {
    logger.warn('Failed to register Ctrl+Shift+O hotkey');
  }

  logger.info('Global hotkeys registered successfully');
}

function unregisterGlobalHotkeys() {
  globalShortcut.unregisterAll();
  logger.info('Global hotkeys unregistered');
}

// ============================================================================
// IPC HANDLERS
// ============================================================================

// ============================================================================
// AI SETTINGS IPC HANDLERS
// ============================================================================

// Get AI settings
ipcMain.handle('get-ai-settings', async () => {
  // Check if API key exists
  const hasApiKey = await getApiKey(aiSettings.provider) !== null;
  return {
    ...aiSettings,
    hasApiKey,
    globalContext: aiSettings.globalContext || {
      enabled: true,
      name: '',
      role: '',
      company: '',
      industry: '',
      communicationStyle: '',
      detailLevel: '',
      customNotes: '',
    },
    companyContext: aiSettings.companyContext || {
      enabled: false,
      content: '',
    },
  };
});

// Save AI settings
ipcMain.handle('save-ai-settings', async (event, settings) => {
  try {
    // SECURITY: Input validation
    if (settings.provider && !['grok', 'ollama', 'openwebui', 'openai', 'custom'].includes(settings.provider)) {
      return { success: false, error: 'Invalid AI provider.' };
    }
    
    if (settings.model && !validateString(settings.model, 100)) {
      return { success: false, error: 'Invalid model name.' };
    }
    
    if (settings.endpoint && !validateString(settings.endpoint, 500)) {
      return { success: false, error: 'Invalid endpoint.' };
    }
    
    // SECURITY: Validate endpoint is a valid URL
    if (settings.endpoint) {
      try {
        const url = new URL(settings.endpoint);
        // Allow HTTP for localhost and private network IPs (for Ollama/Open WebUI)
        const isLocalOrPrivate = 
            url.hostname === 'localhost' || 
            url.hostname === '127.0.0.1' ||
            url.hostname.startsWith('10.') ||           // 10.x.x.x private range
            url.hostname.startsWith('192.168.') ||      // 192.168.x.x private range
            /^172\.(1[6-9]|2[0-9]|3[0-1])\./.test(url.hostname);  // 172.16-31.x.x private range
        
        if (!url.protocol.startsWith('https') && !isLocalOrPrivate) {
          return { success: false, error: 'API endpoint must use HTTPS (except for local/network servers).' };
        }
      } catch {
        return { success: false, error: 'Invalid API endpoint URL.' };
      }
    }
    
    if (settings.apiKey && !validateString(settings.apiKey, 500)) {
      return { success: false, error: 'Invalid API key format.' };
    }
    
    // Save API key separately to secure storage
    if (settings.apiKey) {
      await saveApiKey(settings.provider, settings.apiKey);
    }
    
    // Validate and sanitize global context
    let globalContext = aiSettings.globalContext;
    if (settings.globalContext) {
      globalContext = {
        enabled: settings.globalContext.enabled !== false, // default true
        name: sanitizeString(settings.globalContext.name || '', 100),
        role: sanitizeString(settings.globalContext.role || '', 100),
        company: sanitizeString(settings.globalContext.company || '', 100),
        industry: sanitizeString(settings.globalContext.industry || '', 100),
        communicationStyle: ['professional', 'friendly', 'casual', ''].includes(settings.globalContext.communicationStyle) 
          ? settings.globalContext.communicationStyle : '',
        detailLevel: ['concise', 'detailed', ''].includes(settings.globalContext.detailLevel)
          ? settings.globalContext.detailLevel : '',
        customNotes: sanitizeString(settings.globalContext.customNotes || '', 2000),
      };
    }
    
    // Validate and sanitize company context
    let companyContext = aiSettings.companyContext;
    if (settings.companyContext) {
      companyContext = {
        enabled: !!settings.companyContext.enabled,
        content: sanitizeString(settings.companyContext.content || '', 15000),
      };
    }
    
    // Save other settings to file
    saveSettings({
      provider: settings.provider,
      model: sanitizeString(settings.model, 100),
      endpoint: sanitizeString(settings.endpoint, 500),
      hasApiKey: !!settings.apiKey || aiSettings.hasApiKey,
      globalContext,
      companyContext,
    });
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Process email with AI
ipcMain.handle('process-with-ai', async (event, { prompt, emailData }) => {
  // SECURITY: Rate limiting
  if (!rateLimiter.check('ai-process')) {
    return {
      success: false,
      error: 'Too many AI requests. Please wait a moment and try again.',
    };
  }
  
  // SECURITY: Input validation
  if (!validateString(prompt, 50000)) {
    return {
      success: false,
      error: 'Invalid prompt provided.',
    };
  }
  
  logger.info('Processing with AI', { provider: aiSettings.provider, model: aiSettings.model });
  const result = await processWithAI(prompt, emailData);
  logger.info('AI result', { success: result.success });
  return result;
});

// Test AI connection
ipcMain.handle('test-ai-connection', async () => {
  try {
    const result = await processWithAI('Say "Hello, I am connected!" in exactly those words.', null);
    return result;
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// ============================================================================
// NOTIFICATION AND OUTLOOK IPC HANDLERS
// ============================================================================

// Notification handler
ipcMain.handle('show-notification', async (event, { title, body }) => {
  showNotification(title, body);
  return true;
});

// Check Outlook status
ipcMain.handle('check-outlook-status', async () => {
  console.log('Checking Outlook status...');
  if (!outlookFunctions) {
    console.log('COM not initialized');
    return {
      success: false,
      running: false,
      message: 'COM integration not initialized. Restart the app.'
    };
  }
  
  try {
    const result = await new Promise((resolve, reject) => {
      outlookFunctions.checkOutlookRunning({}, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    console.log('Outlook status result:', JSON.stringify(result));
    return result;
  } catch (error) {
    console.error('Error checking Outlook status:', error);
    return {
      success: false,
      running: false,
      message: error.message
    };
  }
});

// Get active/selected email from Outlook
ipcMain.handle('get-active-email', async () => {
  if (!outlookFunctions) {
    return {
      success: false,
      message: 'COM integration not initialized. Make sure electron-edge-js is installed and rebuilt.',
      data: null
    };
  }

  try {
    console.log('Fetching active email...');
    const result = await new Promise((resolve, reject) => {
      outlookFunctions.getActiveEmail({}, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    console.log('Active email result:', result && result.success ? 'Success' : (result ? result.message : 'No result'));
    return result;
  } catch (error) {
    console.error('Error fetching active email:', error);
    return {
      success: false,
      message: `Error: ${error.message}`,
      data: null
    };
  }
});

// Get recent emails from Inbox (enhanced for full metadata)
ipcMain.handle('get-recent-inbox', async (event, options = {}) => {
  if (!outlookFunctions) {
    return {
      success: false,
      message: 'COM integration not initialized.',
      data: [],
      count: 0
    };
  }

  try {
    logger.info('Fetching inbox', options);
    const params = {
      maxItems: options.maxItems || 50,
      daysBack: options.daysBack || 365,
      folderType: 'Inbox'
    };
    
    console.log(`Fetching recent inbox (last ${params.daysBack} days, max ${params.maxItems} items)...`);
    const result = await new Promise((resolve, reject) => {
      outlookFunctions.getRecentInbox(params, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    console.log('Inbox result:', result && result.success ? `${result.count} emails` : (result ? result.message : 'No result'));
    return result;
  } catch (error) {
    console.error('Error fetching inbox:', error);
    return {
      success: false,
      message: `Error: ${error.message}`,
      data: [],
      count: 0
    };
  }
});

// Get recent emails from Sent Items
ipcMain.handle('get-recent-sent', async (event, options = {}) => {
  if (!outlookFunctions) {
    return {
      success: false,
      message: 'COM integration not initialized.',
      data: [],
      count: 0
    };
  }

  try {
    const params = {
      maxItems: options.maxItems || 50,
      daysBack: options.daysBack || 365,
      folderType: 'SentMail'
    };
    
    console.log(`Fetching sent items (last ${params.daysBack} days, max ${params.maxItems} items)...`);
    const result = await new Promise((resolve, reject) => {
      outlookFunctions.getRecentSent(params, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    console.log('Sent result:', result && result.success ? `${result.count} emails` : (result ? result.message : 'No result'));
    return result;
  } catch (error) {
    console.error('Error fetching sent items:', error);
    return {
      success: false,
      message: `Error: ${error.message}`,
      data: [],
      count: 0
    };
  }
});

// Get full email by EntryID
ipcMain.handle('get-email-by-id', async (event, { entryId }) => {
  if (!outlookFunctions || !edge) {
    return {
      success: false,
      message: 'COM integration not initialized.',
      data: null
    };
  }

  try {
    const getEmailById = createGetEmailByIdFunction();
    if (!getEmailById) {
      return {
        success: false,
        message: 'Could not create email retrieval function.',
        data: null
      };
    }
    
    const result = await getEmailById({ entryId });
    return result;
  } catch (error) {
    console.error('Error fetching email by ID:', error);
    return {
      success: false,
      message: `Error: ${error.message}`,
      data: null
    };
  }
});

// ============================================================================
// OUTLOOK ACTIONS - CREATE/REPLY EMAILS
// ============================================================================

// Create a reply in Outlook with AI-generated content
ipcMain.handle('create-reply-in-outlook', async (event, { entryId, replyBody, replyAll }) => {
  // SECURITY: Rate limiting
  if (!rateLimiter.check('outlook-reply')) {
    return {
      success: false,
      message: 'Too many requests. Please wait a moment.',
    };
  }
  
  // SECURITY: Input validation
  if (!validateString(entryId, 500) || !validateString(replyBody, 100000)) {
    return {
      success: false,
      message: 'Invalid input provided.',
    };
  }
  
  // Clean AI output - remove markdown formatting before sending to Outlook
  const cleanedReplyBody = cleanAiOutput(replyBody);
  
  logger.info('Creating reply in Outlook', { 
    entryId: entryId?.substring(0, 20) + '...', 
    replyAll,
    originalLength: replyBody.length,
    cleanedLength: cleanedReplyBody.length 
  });
  
  if (!edge) {
    return {
      success: false,
      message: 'COM integration not initialized.'
    };
  }

  try {
    const createReply = createReplyInOutlook();
    if (!createReply) {
      return {
        success: false,
        message: 'Could not create reply function.'
      };
    }
    
    const result = await new Promise((resolve, reject) => {
      createReply({ entryId, replyBody: cleanedReplyBody, replyAll }, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    
    logger.info('Reply created', { success: result.success });
    
    if (result.success) {
      showNotification('Grok-Outlook', 'Reply opened in Outlook! Review and send when ready.');
    }
    
    return result;
  } catch (error) {
    logger.error('Error creating reply', { error: error.message });
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
});

// Create a new email in Outlook
ipcMain.handle('create-new-email', async (event, { to, subject, body }) => {
  logger.info('Creating new email in Outlook', { to, subject: subject?.substring(0, 30) });
  
  if (!edge) {
    return {
      success: false,
      message: 'COM integration not initialized.'
    };
  }

  try {
    const createEmail = createNewEmailInOutlook();
    if (!createEmail) {
      return {
        success: false,
        message: 'Could not create email function.'
      };
    }
    
    const result = await new Promise((resolve, reject) => {
      createEmail({ to, subject, body }, (error, result) => {
        if (error) {
          reject(error);
        } else {
          resolve(result);
        }
      });
    });
    
    logger.info('Email created', { success: result.success });
    return result;
  } catch (error) {
    logger.error('Error creating email', { error: error.message });
    return {
      success: false,
      message: `Error: ${error.message}`
    };
  }
});

// ============================================================================
// PROMPT LIBRARY SYSTEM
// ============================================================================

const PROMPTS_FILE = path.join(app.getPath('userData'), 'prompts.json');
const PRESETS_FILE = path.join(app.getPath('userData'), 'prompt-presets.json');
const CRASH_REPORTS_FILE = path.join(app.getPath('userData'), 'crash-reports.json');

// Default starter prompts
const DEFAULT_PROMPTS = [
  // === CORE ACTIONS (Generic, always at top) ===
  {
    id: 'core-summarize',
    name: 'Summarize',
    description: 'Basic email summary',
    template: 'Please provide a summary of this email.\n\n{email_body}',
    category: 'summarize',
    isFavorite: true,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'core-reply',
    name: 'Draft Reply',
    description: 'Draft a reply to the most recent message in the thread',
    template: 'Draft a reply to the most recent message in this email thread.\n\nInstructions:\n- Respond ONLY to the last/most recent person who wrote\n- Read the full thread for context but reply to the latest message\n- Do NOT include: subject line, email headers, "Re:", or any name or signature/sign-off (my email signature is added automatically)\n- Start directly with the response content\n- Keep it concise and professional\n\nEmail thread:\n{email_body}',
    category: 'reply',
    isFavorite: true,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'core-analyze',
    name: 'Analyze',
    description: 'Analyze the email',
    template: 'Please analyze this email.\n\n{email_body}',
    category: 'insights',
    isFavorite: true,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  // === ENHANCED PROMPTS ===
  {
    id: 'default-1',
    name: 'Bullet Point Summary',
    description: 'Summarize email as bullet points',
    template: 'Summarize this email in 5 concise bullet points, highlighting the key information and any action items:\n\n{email_body}',
    category: 'summarize',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'default-2',
    name: 'Action Items Only',
    description: 'Extract only action items and deadlines',
    template: 'Extract ONLY the action items, tasks, and deadlines from this email. Format as a numbered list with due dates if mentioned:\n\n{email_body}',
    category: 'insights',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'default-3',
    name: 'Professional Reply',
    description: 'Draft a formal, professional response',
    template: 'Draft a professional and formal reply to this email. Be courteous, address all points raised, and suggest next steps if appropriate. Only write the body of the reply - do NOT include a subject line or email headers.\n\nOriginal email from {sender}:\n{email_body}',
    category: 'reply',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'default-4',
    name: 'Friendly Reply',
    description: 'Draft a warm, friendly response',
    template: 'Draft a warm and friendly reply to this email. Keep a positive tone while addressing the main points. Only write the body of the reply - do NOT include a subject line or email headers.\n\nOriginal email from {sender}:\n{email_body}',
    category: 'reply',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'default-5',
    name: 'TL;DR',
    description: 'One-sentence summary',
    template: 'Give me a single sentence TL;DR (too long; didn\'t read) summary of this email:\n\n{email_body}',
    category: 'summarize',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
  {
    id: 'default-6',
    name: 'Meeting Notes',
    description: 'Extract meeting details and action items',
    template: 'Extract meeting information from this email:\n1. Meeting date/time\n2. Attendees\n3. Agenda items\n4. Action items with owners\n5. Follow-up required\n\n{email_body}',
    category: 'insights',
    isFavorite: false,
    usageCount: 0,
    isDefault: true,
    createdAt: new Date().toISOString(),
  },
];

// Load prompts from file
function loadPrompts() {
  try {
    if (fs.existsSync(PROMPTS_FILE)) {
      let data = JSON.parse(fs.readFileSync(PROMPTS_FILE, 'utf8'));
      
      // Ensure core prompts exist (add if missing)
      const corePromptIds = ['core-summarize', 'core-reply', 'core-analyze'];
      const existingIds = data.map(p => p.id);
      let added = false;
      
      for (const corePrompt of DEFAULT_PROMPTS.filter(p => corePromptIds.includes(p.id))) {
        if (!existingIds.includes(corePrompt.id)) {
          // Add at the beginning of the array
          data.unshift(corePrompt);
          added = true;
          logger.info('Added missing core prompt', { id: corePrompt.id });
        }
      }
      
      if (added) {
        savePrompts(data);
      }
      
      logger.info('Loaded prompts', { count: data.length });
      return data;
    }
  } catch (error) {
    logger.error('Error loading prompts', { error: error.message });
  }
  
  // Return default prompts if no file exists
  savePrompts(DEFAULT_PROMPTS);
  return DEFAULT_PROMPTS;
}

// Save prompts to file
function savePrompts(prompts) {
  try {
    fs.writeFileSync(PROMPTS_FILE, JSON.stringify(prompts, null, 2));
    logger.info('Saved prompts', { count: prompts.length });
    return true;
  } catch (error) {
    logger.error('Error saving prompts', { error: error.message });
    return false;
  }
}

// Replace placeholders in prompt template
function replacePlaceholders(template, emailData) {
  if (!emailData) return template;
  
  const replacements = {
    '{email_body}': emailData.body || '',
    '{subject}': emailData.subject || '',
    '{sender}': emailData.senderName || emailData.senderEmail || '',
    '{sender_email}': emailData.senderEmail || '',
    '{sender_name}': emailData.senderName || '',
    '{to}': emailData.to || '',
    '{cc}': emailData.cc || '',
    '{date}': emailData.receivedTime || '',
  };
  
  let result = template;
  for (const [placeholder, value] of Object.entries(replacements)) {
    result = result.split(placeholder).join(value);
  }
  
  return result;
}

// IPC: Get all prompts
ipcMain.handle('get-prompts', async () => {
  try {
    const prompts = loadPrompts();
    return { success: true, prompts };
  } catch (error) {
    return { success: false, error: error.message, prompts: [] };
  }
});

// IPC: Add new prompt
ipcMain.handle('add-prompt', async (event, prompt) => {
  try {
    // Validate
    if (!prompt.name || !prompt.template) {
      return { success: false, error: 'Name and template are required' };
    }
    
    const prompts = loadPrompts();
    
    const newPrompt = {
      id: `prompt-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
      name: sanitizeString(prompt.name, 100),
      description: sanitizeString(prompt.description || '', 500),
      template: sanitizeString(prompt.template, 5000),
      category: ['summarize', 'reply', 'insights', 'custom'].includes(prompt.category) ? prompt.category : 'custom',
      isFavorite: !!prompt.isFavorite,
      usageCount: 0,
      isDefault: false,
      createdAt: new Date().toISOString(),
    };
    
    prompts.push(newPrompt);
    savePrompts(prompts);
    
    logger.info('Added prompt', { id: newPrompt.id, name: newPrompt.name });
    return { success: true, prompt: newPrompt };
  } catch (error) {
    logger.error('Error adding prompt', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Update prompt
ipcMain.handle('update-prompt', async (event, { id, updates }) => {
  try {
    const prompts = loadPrompts();
    const index = prompts.findIndex(p => p.id === id);
    
    if (index === -1) {
      return { success: false, error: 'Prompt not found' };
    }
    
    // Update fields
    if (updates.name) prompts[index].name = sanitizeString(updates.name, 100);
    if (updates.description !== undefined) prompts[index].description = sanitizeString(updates.description, 500);
    if (updates.template) prompts[index].template = sanitizeString(updates.template, 5000);
    if (updates.category) prompts[index].category = ['summarize', 'reply', 'insights', 'custom'].includes(updates.category) ? updates.category : prompts[index].category;
    if (updates.isFavorite !== undefined) prompts[index].isFavorite = !!updates.isFavorite;
    prompts[index].updatedAt = new Date().toISOString();
    
    savePrompts(prompts);
    
    logger.info('Updated prompt', { id });
    return { success: true, prompt: prompts[index] };
  } catch (error) {
    logger.error('Error updating prompt', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Delete prompt
ipcMain.handle('delete-prompt', async (event, { id }) => {
  try {
    const prompts = loadPrompts();
    const index = prompts.findIndex(p => p.id === id);
    
    if (index === -1) {
      return { success: false, error: 'Prompt not found' };
    }
    
    // Don't allow deleting default prompts
    if (prompts[index].isDefault) {
      return { success: false, error: 'Cannot delete default prompts' };
    }
    
    prompts.splice(index, 1);
    savePrompts(prompts);
    
    logger.info('Deleted prompt', { id });
    return { success: true };
  } catch (error) {
    logger.error('Error deleting prompt', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Increment usage count
ipcMain.handle('increment-prompt-usage', async (event, { id }) => {
  try {
    const prompts = loadPrompts();
    const index = prompts.findIndex(p => p.id === id);
    
    if (index !== -1) {
      prompts[index].usageCount = (prompts[index].usageCount || 0) + 1;
      savePrompts(prompts);
    }
    
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Process with custom prompt
ipcMain.handle('process-with-custom-prompt', async (event, { promptId, emailData }) => {
  try {
    const prompts = loadPrompts();
    const prompt = prompts.find(p => p.id === promptId);
    
    if (!prompt) {
      return { success: false, error: 'Prompt not found' };
    }
    
    // Replace placeholders
    const processedTemplate = replacePlaceholders(prompt.template, emailData);
    
    // Increment usage
    const index = prompts.findIndex(p => p.id === promptId);
    if (index !== -1) {
      prompts[index].usageCount = (prompts[index].usageCount || 0) + 1;
      savePrompts(prompts);
    }
    
    // Process with AI
    const result = await processWithAI(processedTemplate, null); // Don't pass emailData again, it's in the template
    
    return result;
  } catch (error) {
    logger.error('Error processing with custom prompt', { error: error.message });
    return { success: false, error: error.message };
  }
});

// ============================================================================
// PROMPT PRESETS
// ============================================================================

// Load presets from file
function loadPresets() {
  try {
    if (fs.existsSync(PRESETS_FILE)) {
      const data = JSON.parse(fs.readFileSync(PRESETS_FILE, 'utf8'));
      logger.info('Loaded presets', { count: data.length });
      return data;
    }
  } catch (error) {
    logger.error('Error loading presets', { error: error.message });
  }
  return [];
}

// Save presets to file
function savePresets(presets) {
  try {
    fs.writeFileSync(PRESETS_FILE, JSON.stringify(presets, null, 2));
    logger.info('Saved presets', { count: presets.length });
    return true;
  } catch (error) {
    logger.error('Error saving presets', { error: error.message });
    return false;
  }
}

// IPC: Get presets
ipcMain.handle('get-presets', async () => {
  try {
    const presets = loadPresets();
    return { success: true, presets };
  } catch (error) {
    return { success: false, error: error.message, presets: [] };
  }
});

// IPC: Save preset
ipcMain.handle('save-preset', async (event, { name, promptIds }) => {
  try {
    if (!name || !promptIds || promptIds.length === 0) {
      return { success: false, error: 'Name and at least one prompt required' };
    }
    
    const presets = loadPresets();
    const newPreset = {
      id: `preset-${Date.now()}`,
      name: sanitizeString(name, 100),
      promptIds,
      createdAt: new Date().toISOString(),
    };
    
    presets.push(newPreset);
    savePresets(presets);
    
    logger.info('Saved preset', { id: newPreset.id, name: newPreset.name });
    return { success: true, preset: newPreset };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Delete preset
ipcMain.handle('delete-preset', async (event, { id }) => {
  try {
    const presets = loadPresets();
    const index = presets.findIndex(p => p.id === id);
    
    if (index === -1) {
      return { success: false, error: 'Preset not found' };
    }
    
    presets.splice(index, 1);
    savePresets(presets);
    
    logger.info('Deleted preset', { id });
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// ============================================================================
// CRASH REPORTING (Opt-in Telemetry)
// ============================================================================

// Load crash reports
function loadCrashReports() {
  try {
    if (fs.existsSync(CRASH_REPORTS_FILE)) {
      return JSON.parse(fs.readFileSync(CRASH_REPORTS_FILE, 'utf8'));
    }
  } catch (error) {
    console.error('Error loading crash reports:', error);
  }
  return [];
}

// Save crash report
function saveCrashReport(report) {
  try {
    // Check if telemetry is enabled
    if (!aiSettings.telemetryEnabled) {
      return false;
    }
    
    const reports = loadCrashReports();
    
    // Add new report
    reports.unshift({
      id: `crash-${Date.now()}`,
      timestamp: new Date().toISOString(),
      version: app.getVersion(),
      platform: process.platform,
      arch: process.arch,
      electronVersion: process.versions.electron,
      nodeVersion: process.versions.node,
      ...report,
    });
    
    // Keep only last 50 reports
    const trimmed = reports.slice(0, 50);
    
    fs.writeFileSync(CRASH_REPORTS_FILE, JSON.stringify(trimmed, null, 2));
    logger.info('Crash report saved', { type: report.type });
    return true;
  } catch (error) {
    console.error('Error saving crash report:', error);
    return false;
  }
}

// Set up global error handlers
function setupCrashReporting() {
  // Uncaught exceptions in main process
  process.on('uncaughtException', (error) => {
    console.error('Uncaught Exception:', error);
    saveCrashReport({
      type: 'uncaughtException',
      error: error.message,
      stack: error.stack,
    });
  });

  // Unhandled promise rejections
  process.on('unhandledRejection', (reason, promise) => {
    console.error('Unhandled Rejection:', reason);
    saveCrashReport({
      type: 'unhandledRejection',
      error: reason?.message || String(reason),
      stack: reason?.stack || 'No stack trace',
    });
  });
}

// IPC: Get crash reports
ipcMain.handle('get-crash-reports', async () => {
  try {
    const reports = loadCrashReports();
    return { success: true, reports };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Clear crash reports
ipcMain.handle('clear-crash-reports', async () => {
  try {
    fs.writeFileSync(CRASH_REPORTS_FILE, JSON.stringify([], null, 2));
    logger.info('Crash reports cleared');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Export crash reports (for manual sharing)
ipcMain.handle('export-crash-reports', async () => {
  try {
    const reports = loadCrashReports();
    if (reports.length === 0) {
      return { success: false, error: 'No crash reports to export' };
    }
    
    const { filePath } = await dialog.showSaveDialog(mainWindow, {
      title: 'Export Crash Reports',
      defaultPath: `grok-outlook-crash-reports-${Date.now()}.json`,
      filters: [{ name: 'JSON Files', extensions: ['json'] }],
    });
    
    if (filePath) {
      // Anonymize data before export
      const anonymized = reports.map(r => ({
        ...r,
        // Remove any potentially identifying info
      }));
      fs.writeFileSync(filePath, JSON.stringify(anonymized, null, 2));
      logger.info('Crash reports exported', { path: filePath });
      return { success: true, path: filePath };
    }
    
    return { success: false, error: 'Export cancelled' };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Report renderer error (called from renderer process)
ipcMain.handle('report-renderer-error', async (event, { error, componentStack, url }) => {
  try {
    saveCrashReport({
      type: 'rendererError',
      error: error,
      componentStack: componentStack,
      url: url,
    });
    return { success: true };
  } catch (err) {
    return { success: false, error: err.message };
  }
});

// IPC: Process with multiple prompts
ipcMain.handle('process-with-multi-prompts', async (event, { promptIds, emailData, quickNotes }) => {
  try {
    const prompts = loadPrompts();
    
    // Build combined prompt from selected prompts
    let combinedPrompt = '';
    const usedPromptNames = [];
    
    if (promptIds && promptIds.length > 0) {
      for (const promptId of promptIds) {
        const prompt = prompts.find(p => p.id === promptId);
        if (prompt) {
          const processedTemplate = replacePlaceholders(prompt.template, emailData);
          combinedPrompt += `### ${prompt.name}:\n${processedTemplate}\n\n`;
          usedPromptNames.push(prompt.name);
          
          // Increment usage
          const index = prompts.findIndex(p => p.id === promptId);
          if (index !== -1) {
            prompts[index].usageCount = (prompts[index].usageCount || 0) + 1;
          }
        }
      }
      savePrompts(prompts);
    }
    
    // If no prompts selected, just send the email for general analysis
    if (!combinedPrompt) {
      combinedPrompt = 'Please analyze and respond to this email:';
    }
    
    // Add quick notes if provided
    if (quickNotes && quickNotes.trim()) {
      combinedPrompt += `\n### Additional Instructions (one-time):\n${quickNotes.trim()}\n`;
      logger.info('Including quick notes', { length: quickNotes.length });
    }
    
    // Process with AI
    const result = await processWithAI(combinedPrompt, emailData);
    
    return {
      ...result,
      usedPrompts: usedPromptNames,
      hadQuickNotes: !!(quickNotes && quickNotes.trim()),
    };
  } catch (error) {
    logger.error('Error processing with multi-prompts', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Smart Shot - Process email with attachment analysis
ipcMain.handle('smart-shot', async (event, { emailData, promptIds, quickNotes }) => {
  try {
    logger.info('Smart Shot initiated', { 
      entryId: emailData?.entryId?.substring(0, 20),
      promptCount: promptIds?.length 
    });
    
    const results = {
      attachmentSummaries: [],
      aiResult: null,
      usedPrompts: [],
      hadQuickNotes: false,
    };
    
    // Step 1: Extract attachments
    logger.info('Smart Shot: Extracting attachments...');
    let attachments = [];
    
    if (emailData?.entryId) {
      const getAttachments = createGetAttachmentsFunction();
      if (getAttachments) {
        // Use the same TEMP_FOLDER constant used elsewhere
        const tempFolder = path.join(require('os').tmpdir(), 'grok-outlook-attachments');
        if (!fs.existsSync(tempFolder)) {
          fs.mkdirSync(tempFolder, { recursive: true });
        }
        
        logger.info('Smart Shot: Using temp folder', { tempFolder });
        
        const attachResult = await new Promise((resolve, reject) => {
          getAttachments({ entryId: emailData.entryId, tempFolder }, (error, result) => {
            if (error) reject(error);
            else resolve(result);
          });
        });
        
        if (attachResult.success && attachResult.attachments) {
          attachments = attachResult.attachments;
          logger.info('Smart Shot: Attachments extracted', { 
            count: attachments.length,
            files: attachments.map(a => a.name)
          });
        } else {
          logger.info('Smart Shot: No attachments found or extraction failed', { 
            success: attachResult.success, 
            message: attachResult.message 
          });
        }
      }
    }
    
    // Step 2: Process each attachment and extract key info
    if (attachments.length > 0) {
      for (let i = 0; i < attachments.length; i++) {
        const attachment = attachments[i];
        
        // C# returns: fileName, savedPath, size, extension
        const attachName = attachment.fileName || attachment.name;
        const attachPath = attachment.savedPath || attachment.tempPath;
        
        // Defensive check for attachment properties
        if (!attachment || !attachName || !attachPath) {
          logger.warn('Smart Shot: Skipping invalid attachment', { attachment });
          continue;
        }
        
        // Get extension early so it's available in catch block
        const ext = path.extname(attachName || '').toLowerCase();
        
        logger.info(`Smart Shot: Processing attachment ${i + 1}/${attachments.length}`, { 
          name: attachName,
          path: attachPath,
          ext: ext
        });
        
        try {
          // Read the file content
          let fileContent = '';
          const filePath = attachPath;
          
          // Extract text based on file type
          if (ext === '.pdf') {
            if (pdfParse) {
              const dataBuffer = fs.readFileSync(filePath);
              const pdfData = await pdfParse(dataBuffer);
              fileContent = pdfData.text || '';
              
              // If PDF has little text, might be scanned - try OCR
              if (fileContent.trim().length < 100 && Tesseract) {
                logger.info('Smart Shot: PDF appears scanned, attempting OCR...');
                const ocrResult = await Tesseract.recognize(filePath, 'eng');
                fileContent = ocrResult.data.text || fileContent;
              }
            }
          } else if (ext === '.docx' || ext === '.doc') {
            if (mammoth) {
              const result = await mammoth.extractRawText({ path: filePath });
              fileContent = result.value || '';
            }
          } else if (ext === '.txt' || ext === '.md') {
            fileContent = fs.readFileSync(filePath, 'utf8');
          } else if (ext === '.csv') {
            fileContent = fs.readFileSync(filePath, 'utf8');
          } else if (ext === '.xlsx' || ext === '.xls') {
            // Excel files
            if (xlsx) {
              const workbook = xlsx.readFile(filePath);
              const sheetNames = workbook.SheetNames;
              let excelContent = '';
              
              // Process each sheet
              for (const sheetName of sheetNames) {
                const sheet = workbook.Sheets[sheetName];
                const csvData = xlsx.utils.sheet_to_csv(sheet);
                excelContent += `\n--- Sheet: ${sheetName} ---\n${csvData}`;
              }
              
              fileContent = excelContent.trim();
              logger.info('Smart Shot: Excel file processed', { 
                sheets: sheetNames.length, 
                chars: fileContent.length 
              });
            }
          } else if (['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp'].includes(ext)) {
            // Image - use OCR
            if (Tesseract) {
              const ocrResult = await Tesseract.recognize(filePath, 'eng');
              fileContent = ocrResult.data.text || '';
            }
          }
          
          // Limit content size (5000 chars max per file)
          if (fileContent.length > 5000) {
            fileContent = fileContent.substring(0, 5000) + '\n... [truncated]';
          }
          
          // Skip if no content extracted
          if (!fileContent.trim()) {
            results.attachmentSummaries.push({
              filename: attachName,
              summary: '[Could not extract text from this file]',
              type: ext,
            });
            continue;
          }
          
          // Extract key info using AI
          const keyInfoPrompt = `Extract the key information from this document. Provide a concise summary with bullet points highlighting the most important facts, figures, dates, names, and action items.

Document name: ${attachName}
Document type: ${ext}

--- DOCUMENT CONTENT ---
${fileContent}
--- END DOCUMENT ---

Provide key information in bullet points (max 10 bullets):`;

          const keyInfoResult = await processWithAI(keyInfoPrompt, null);
          
          results.attachmentSummaries.push({
            filename: attachName,
            summary: keyInfoResult.success ? keyInfoResult.content : '[Error extracting key info]',
            type: ext,
            charCount: fileContent.length,
          });
          
        } catch (attachError) {
          logger.error('Smart Shot: Error processing attachment', { 
            name: attachName, 
            error: attachError.message 
          });
          results.attachmentSummaries.push({
            filename: attachName || 'Unknown',
            summary: `[Error: ${attachError.message}]`,
            type: ext || 'unknown',
          });
        }
      }
    }
    
    // Step 3: Build combined context with attachments
    let attachmentContext = '';
    if (results.attachmentSummaries.length > 0) {
      attachmentContext = '\n\n--- ATTACHMENT SUMMARIES ---\n';
      for (const att of results.attachmentSummaries) {
        attachmentContext += `\n📄 File: ${att.filename}\nKey Information:\n${att.summary}\n`;
      }
      attachmentContext += '\n--- END ATTACHMENTS ---\n';
      attachmentContext += '\nPlease consider both the email content AND the attachment information above when drafting your response.';
    }
    
    // Step 4: Process with prompts (same as One Shot)
    const prompts = loadPrompts();
    let combinedPrompt = '';
    const usedPromptNames = [];
    
    if (promptIds && promptIds.length > 0) {
      for (const promptId of promptIds) {
        const prompt = prompts.find(p => p.id === promptId);
        if (prompt) {
          const processedTemplate = replacePlaceholders(prompt.template, emailData);
          combinedPrompt += `### ${prompt.name}:\n${processedTemplate}\n\n`;
          usedPromptNames.push(prompt.name);
          
          // Increment usage
          const index = prompts.findIndex(p => p.id === promptId);
          if (index !== -1) {
            prompts[index].usageCount = (prompts[index].usageCount || 0) + 1;
          }
        }
      }
      savePrompts(prompts);
    }
    
    if (!combinedPrompt) {
      combinedPrompt = 'Please analyze and respond to this email:';
    }
    
    // Add quick notes if provided
    if (quickNotes && quickNotes.trim()) {
      combinedPrompt += `\n### Additional Instructions (one-time):\n${quickNotes.trim()}\n`;
      results.hadQuickNotes = true;
    }
    
    // Add attachment context
    combinedPrompt += attachmentContext;
    
    // Step 5: Process with AI
    const aiResult = await processWithAI(combinedPrompt, emailData);
    results.aiResult = aiResult;
    results.usedPrompts = usedPromptNames;
    
    // Clean up temp files
    for (const attachment of attachments) {
      try {
        const attachPath = attachment.savedPath || attachment.tempPath;
        if (attachPath && fs.existsSync(attachPath)) {
          fs.unlinkSync(attachPath);
        }
      } catch (e) { /* ignore cleanup errors */ }
    }
    
    logger.info('Smart Shot completed', { 
      attachmentCount: results.attachmentSummaries.length,
      aiSuccess: aiResult?.success 
    });
    
    return {
      success: true,
      ...results,
    };
    
  } catch (error) {
    logger.error('Smart Shot error', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Export prompts
ipcMain.handle('export-prompts', async () => {
  try {
    const prompts = loadPrompts();
    // Filter out default prompts for export
    const customPrompts = prompts.filter(p => !p.isDefault);
    return { success: true, data: JSON.stringify(customPrompts, null, 2) };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// IPC: Import prompts
ipcMain.handle('import-prompts', async (event, { jsonData }) => {
  try {
    const imported = JSON.parse(jsonData);
    
    if (!Array.isArray(imported)) {
      return { success: false, error: 'Invalid format - expected array of prompts' };
    }
    
    const prompts = loadPrompts();
    let addedCount = 0;
    
    for (const p of imported) {
      if (p.name && p.template) {
        prompts.push({
          id: `imported-${Date.now()}-${Math.random().toString(36).substr(2, 9)}`,
          name: sanitizeString(p.name, 100),
          description: sanitizeString(p.description || '', 500),
          template: sanitizeString(p.template, 5000),
          category: ['summarize', 'reply', 'insights', 'custom'].includes(p.category) ? p.category : 'custom',
          isFavorite: false,
          usageCount: 0,
          isDefault: false,
          createdAt: new Date().toISOString(),
        });
        addedCount++;
      }
    }
    
    savePrompts(prompts);
    logger.info('Imported prompts', { count: addedCount });
    return { success: true, count: addedCount };
  } catch (error) {
    logger.error('Error importing prompts', { error: error.message });
    return { success: false, error: error.message };
  }
});

// ============================================================================
// FILE ANALYSIS SYSTEM
// ============================================================================

// Temp folder for attachments
const TEMP_FOLDER = path.join(app.getPath('temp'), 'grok-outlook-attachments');

// Supported file types
const SUPPORTED_TEXT_EXTENSIONS = ['.txt', '.md', '.json', '.xml', '.html', '.htm', '.css', '.js', '.ts', '.py', '.java', '.c', '.cpp', '.h', '.csv', '.log'];
const SUPPORTED_DOC_EXTENSIONS = ['.pdf', '.docx', '.doc'];
const SUPPORTED_IMAGE_EXTENSIONS = ['.png', '.jpg', '.jpeg', '.gif', '.webp', '.bmp'];

// Extract text from various file types
async function extractTextFromFile(filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const fileName = path.basename(filePath);
  
  logger.info('Extracting text from file', { fileName, ext, pdfParseAvailable: !!pdfParse });
  
  try {
    // Plain text files
    if (SUPPORTED_TEXT_EXTENSIONS.includes(ext)) {
      const content = fs.readFileSync(filePath, 'utf8');
      return {
        success: true,
        text: content,
        type: 'text',
        fileName,
      };
    }
    
    // PDF files
    if (ext === '.pdf') {
      if (!pdfParse) {
        logger.error('PDF parse library not loaded');
        return {
          success: false,
          error: 'PDF parsing library not available. Please restart the app.',
          fileName,
        };
      }
      
      logger.info('Parsing PDF file', { fileName });
      const dataBuffer = fs.readFileSync(filePath);
      const data = await pdfParse(dataBuffer);
      logger.info('PDF parsed', { pages: data.numpages, textLength: data.text?.length });
      
      // Check if PDF appears to be scanned (very little text for the number of pages)
      const avgCharsPerPage = data.text.length / (data.numpages || 1);
      const isLikelyScanned = avgCharsPerPage < 100; // Less than 100 chars per page suggests scanned
      
      if (isLikelyScanned && data.text.trim().length < 50) {
        return {
          success: true,
          text: data.text || '',
          type: 'pdf',
          pages: data.numpages,
          fileName,
          isScanned: true,
          warning: 'This PDF appears to be scanned/image-based with little extractable text. For best results: 1) Use a PDF with selectable text, or 2) Export pages as images and upload those for OCR analysis.',
        };
      }
      
      return {
        success: true,
        text: data.text,
        type: 'pdf',
        pages: data.numpages,
        fileName,
        isScanned: false,
      };
    }
    
    // Word documents
    if ((ext === '.docx' || ext === '.doc') && mammoth) {
      const result = await mammoth.extractRawText({ path: filePath });
      return {
        success: true,
        text: result.value,
        type: 'word',
        fileName,
      };
    }
    
    // CSV files
    if (ext === '.csv' && csvParse) {
      const content = fs.readFileSync(filePath, 'utf8');
      const records = csvParse.parse(content, { columns: true, skip_empty_lines: true });
      const textRepresentation = records.map((row, i) => 
        `Row ${i + 1}: ${Object.entries(row).map(([k, v]) => `${k}=${v}`).join(', ')}`
      ).join('\n');
      return {
        success: true,
        text: textRepresentation,
        type: 'csv',
        rowCount: records.length,
        columns: records.length > 0 ? Object.keys(records[0]) : [],
        fileName,
      };
    }
    
    // Images - try OCR first, also provide base64 for vision API
    if (SUPPORTED_IMAGE_EXTENSIONS.includes(ext)) {
      const imageBuffer = fs.readFileSync(filePath);
      const base64 = imageBuffer.toString('base64');
      const mimeType = ext === '.png' ? 'image/png' : 
                       ext === '.gif' ? 'image/gif' :
                       ext === '.webp' ? 'image/webp' : 'image/jpeg';
      
      // Try OCR to extract text from image
      let ocrText = '';
      if (Tesseract) {
        try {
          logger.info('Running OCR on image', { fileName });
          const ocrResult = await Tesseract.recognize(filePath, 'eng', {
            logger: m => {
              if (m.status === 'recognizing text') {
                // Progress logging if needed
              }
            }
          });
          ocrText = ocrResult.data.text.trim();
          logger.info('OCR completed', { fileName, textLength: ocrText.length });
        } catch (ocrError) {
          logger.warn('OCR failed, will use vision API', { error: ocrError.message });
        }
      }
      
      return {
        success: true,
        text: ocrText || '[Image file - OCR found no text, use vision API for analysis]',
        type: 'image',
        hasOcrText: ocrText.length > 0,
        base64,
        mimeType,
        fileName,
      };
    }
    
    return {
      success: false,
      error: `Unsupported file type: ${ext}`,
      fileName,
    };
  } catch (error) {
    logger.error('Error extracting text from file', { filePath, error: error.message });
    return {
      success: false,
      error: error.message,
      fileName,
    };
  }
}

// Process file with AI
async function analyzeFileWithAI(fileContent, analysisType = 'summarize') {
  let prompt;
  
  switch (analysisType) {
    case 'summarize':
      prompt = `Please provide a comprehensive summary of this document. Include:
1. Main topic/purpose
2. Key points and findings
3. Important details or data
4. Conclusions or recommendations (if any)

Document content:
${fileContent}`;
      break;
    case 'extract':
      prompt = `Extract and list all key information from this document:
1. Important dates, numbers, and statistics
2. Names and organizations mentioned
3. Action items or tasks
4. Key decisions or conclusions
5. Any deadlines or time-sensitive information

Document content:
${fileContent}`;
      break;
    case 'questions':
      prompt = `Based on this document, generate 5-10 important questions that someone might have after reading it, along with the answers found in the document.

Document content:
${fileContent}`;
      break;
    default:
      prompt = `Analyze the following document and provide insights:\n\n${fileContent}`;
  }
  
  return await processWithAI(prompt, null);
}

// IPC: Extract attachments from email
ipcMain.handle('get-email-attachments', async (event, { entryId }) => {
  logger.info('Extracting attachments', { entryId: entryId?.substring(0, 20) + '...' });
  
  if (!edge) {
    return {
      success: false,
      message: 'COM integration not initialized.',
      attachments: []
    };
  }

  // Create temp folder
  if (!fs.existsSync(TEMP_FOLDER)) {
    fs.mkdirSync(TEMP_FOLDER, { recursive: true });
  }

  try {
    const getAttachments = createGetAttachmentsFunction();
    if (!getAttachments) {
      return {
        success: false,
        message: 'Could not create attachment function.',
        attachments: []
      };
    }
    
    const result = await new Promise((resolve, reject) => {
      getAttachments({ entryId, tempFolder: TEMP_FOLDER }, (error, result) => {
        if (error) reject(error);
        else resolve(result);
      });
    });
    
    logger.info('Attachments extracted', { count: result.attachments?.length || 0 });
    return result;
  } catch (error) {
    logger.error('Error extracting attachments', { error: error.message });
    return {
      success: false,
      message: `Error: ${error.message}`,
      attachments: []
    };
  }
});

// ============================================================================
// CONTACT LOOKUP & RESEARCH
// ============================================================================

// IPC: Lookup contact by email
ipcMain.handle('lookup-contact', async (event, { email }) => {
  logger.info('Looking up contact', { email });
  
  if (!edge) {
    return {
      success: false,
      found: false,
      message: 'COM integration not initialized.',
      contact: null
    };
  }

  try {
    const lookupFunc = createContactLookupFunction();
    if (!lookupFunc) {
      return {
        success: false,
        found: false,
        message: 'Could not create lookup function.',
        contact: null
      };
    }
    
    const result = await new Promise((resolve, reject) => {
      lookupFunc({ email }, (error, result) => {
        if (error) reject(error);
        else resolve(result);
      });
    });
    
    logger.info('Contact lookup complete', { found: result.found, source: result.source });
    return result;
  } catch (error) {
    logger.error('Error looking up contact', { error: error.message });
    return {
      success: false,
      found: false,
      message: `Error: ${error.message}`,
      contact: null
    };
  }
});

// IPC: Get email history with contact
ipcMain.handle('get-email-history', async (event, { email }) => {
  logger.info('Getting email history', { email });
  
  if (!edge) {
    return {
      success: false,
      totalCount: 0,
      inboxCount: 0,
      sentCount: 0,
      lastEmailDate: ''
    };
  }

  try {
    const historyFunc = createEmailHistoryFunction();
    if (!historyFunc) {
      return {
        success: false,
        totalCount: 0,
        inboxCount: 0,
        sentCount: 0,
        lastEmailDate: ''
      };
    }
    
    const result = await new Promise((resolve, reject) => {
      historyFunc({ email }, (error, result) => {
        if (error) reject(error);
        else resolve(result);
      });
    });
    
    logger.info('Email history retrieved', { totalCount: result.totalCount });
    return result;
  } catch (error) {
    logger.error('Error getting email history', { error: error.message });
    return {
      success: false,
      totalCount: 0,
      inboxCount: 0,
      sentCount: 0,
      lastEmailDate: '',
      error: error.message
    };
  }
});

// IPC: Deep research on contact (uses AI)
ipcMain.handle('research-contact', async (event, { name, email, company, title }) => {
  logger.info('Researching contact', { name, company });
  
  // Rate limiting
  if (!rateLimiter.check('contact-research')) {
    return { success: false, error: 'Too many requests. Please wait.' };
  }
  
  try {
    // Build research prompt
    const researchPrompt = `Research this person and provide a comprehensive profile:

Name: ${name || 'Unknown'}
Email: ${email}
Company: ${company || 'Unknown'}
Title: ${title || 'Unknown'}

Please find and provide:
1. **LinkedIn Summary**: Their professional background, experience, education, and skills based on what you can find.
2. **Company Profile**: Information about ${company || 'their company'} - size, industry, revenue, headquarters, key facts.
3. **Recent News**: Any recent news mentions of this person or their company.
4. **Communication Insights**: Based on their role and industry, suggest how to best communicate with them.
5. **Talking Points**: 3-5 suggested conversation topics based on their background.

Format your response with clear headers for each section. If you cannot find specific information, note what's available and what couldn't be verified.`;

    // Process with AI
    const result = await processWithAI(researchPrompt, { body: '', subject: 'Contact Research' });
    
    if (result.success) {
      return {
        success: true,
        research: result.content,
        generatedAt: new Date().toISOString()
      };
    } else {
      return {
        success: false,
        error: result.error || 'Research failed'
      };
    }
  } catch (error) {
    logger.error('Error researching contact', { error: error.message });
    return {
      success: false,
      error: error.message
    };
  }
});

// IPC: Save research notes to contact
ipcMain.handle('save-contact-notes', async (event, { email, notes, append }) => {
  logger.info('Saving contact notes', { email, append });
  
  if (!edge) {
    return { success: false, message: 'COM integration not initialized.' };
  }

  try {
    const saveFunc = createSaveContactNotesFunction();
    if (!saveFunc) {
      return { success: false, message: 'Could not create save function.' };
    }
    
    const result = await new Promise((resolve, reject) => {
      saveFunc({ email, notes, append: append !== false }, (error, result) => {
        if (error) reject(error);
        else resolve(result);
      });
    });
    
    logger.info('Contact notes saved', { success: result.success });
    return result;
  } catch (error) {
    logger.error('Error saving contact notes', { error: error.message });
    return { success: false, message: error.message };
  }
});

// IPC: Read and extract text from a file
ipcMain.handle('read-file-content', async (event, { filePath }) => {
  logger.info('Reading file content', { filePath: path.basename(filePath) });
  
  if (!validateString(filePath, 1000)) {
    return { success: false, error: 'Invalid file path' };
  }
  
  // Security: Only allow files from temp folder or user-selected files
  const normalizedPath = path.normalize(filePath);
  
  try {
    const result = await extractTextFromFile(normalizedPath);
    return result;
  } catch (error) {
    logger.error('Error reading file', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Analyze file with AI
ipcMain.handle('analyze-file', async (event, { filePath, analysisType, useOcr }) => {
  logger.info('Analyzing file with AI', { filePath: path.basename(filePath), analysisType, useOcr });
  
  // Rate limiting
  if (!rateLimiter.check('file-analysis')) {
    return { success: false, error: 'Too many requests. Please wait.' };
  }
  
  try {
    // First extract text
    const extracted = await extractTextFromFile(filePath);
    if (!extracted.success) {
      return extracted;
    }
    
    // For images with OCR text, we can use the text directly
    if (extracted.type === 'image') {
      // If we have OCR text, use it for text-based analysis
      if (extracted.hasOcrText && extracted.text.length > 20) {
        logger.info('Using OCR text for analysis', { textLength: extracted.text.length });
        const result = await analyzeFileWithAI(extracted.text, analysisType);
        return {
          ...result,
          fileName: extracted.fileName,
          fileType: 'image-ocr',
          ocrUsed: true,
        };
      }
      // Otherwise use vision API
      return await analyzeImageWithAI(extracted.base64, extracted.mimeType, analysisType);
    }
    
    // For scanned PDFs, warn the user
    if (extracted.isScanned && extracted.warning) {
      return {
        success: true,
        content: extracted.warning + '\n\nExtracted text (if any):\n' + (extracted.text || '(none)'),
        fileName: extracted.fileName,
        fileType: 'pdf-scanned',
        isScanned: true,
      };
    }
    
    // Truncate very long documents
    let textToAnalyze = extracted.text;
    if (textToAnalyze.length > 50000) {
      textToAnalyze = textToAnalyze.substring(0, 50000) + '\n\n[Document truncated due to length...]';
    }
    
    // Analyze with AI
    const result = await analyzeFileWithAI(textToAnalyze, analysisType);
    return {
      ...result,
      fileName: extracted.fileName,
      fileType: extracted.type,
      pages: extracted.pages,
    };
  } catch (error) {
    logger.error('Error analyzing file', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Analyze image with vision API
async function analyzeImageWithAI(base64, mimeType, analysisType = 'describe') {
  const provider = aiSettings.provider;
  
  // Grok supports vision
  if (provider === 'grok') {
    const apiKey = await getApiKey(provider);
    if (!apiKey) {
      return { success: false, error: 'API key not configured.' };
    }
    
    let prompt;
    switch (analysisType) {
      case 'summarize':
        prompt = 'Describe this image in detail. What is shown? What are the key elements?';
        break;
      case 'extract':
        prompt = 'Extract all text, numbers, and important information visible in this image.';
        break;
      case 'questions':
        prompt = 'What questions might someone have about this image? Provide answers based on what you can see.';
        break;
      default:
        prompt = 'Analyze this image and describe what you see.';
    }
    
    try {
      const response = await axios.post(
        'https://api.x.ai/v1/chat/completions',
        {
          model: aiSettings.model,
          messages: [
            {
              role: 'user',
              content: [
                { type: 'text', text: prompt },
                { type: 'image_url', image_url: { url: `data:${mimeType};base64,${base64}` } }
              ]
            }
          ],
          max_tokens: 2000,
        },
        {
          headers: {
            'Content-Type': 'application/json',
            'Authorization': `Bearer ${apiKey}`,
          },
          timeout: 120000, // 2 minutes
        }
      );
      
      return {
        success: true,
        content: response.data.choices[0].message.content,
        type: 'image',
      };
    } catch (error) {
      logger.error('Vision API error', { error: error.message });
      return { success: false, error: error.response?.data?.error?.message || error.message };
    }
  }
  
  return { success: false, error: 'Image analysis requires Grok API with vision support.' };
}

// IPC: Open file dialog for user to select files
ipcMain.handle('open-file-dialog', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Documents', extensions: ['pdf', 'docx', 'doc', 'txt', 'md', 'csv', 'json'] },
      { name: 'Images', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp'] },
      { name: 'All Files', extensions: ['*'] }
    ]
  });
  
  if (result.canceled) {
    return { success: false, canceled: true, files: [] };
  }
  
  // Get file info for each selected file
  const files = result.filePaths.map(filePath => {
    const stats = fs.statSync(filePath);
    return {
      path: filePath,
      name: path.basename(filePath),
      size: stats.size,
      extension: path.extname(filePath).toLowerCase(),
    };
  });
  
  return { success: true, files };
});

// IPC: Run OCR on an image file
ipcMain.handle('run-ocr', async (event, { filePath }) => {
  logger.info('Running OCR', { filePath: path.basename(filePath) });
  
  if (!Tesseract) {
    return { success: false, error: 'OCR not available' };
  }
  
  // Rate limiting
  if (!rateLimiter.check('ocr')) {
    return { success: false, error: 'Too many OCR requests. Please wait.' };
  }
  
  try {
    const result = await Tesseract.recognize(filePath, 'eng', {
      logger: m => {
        // Could send progress updates via IPC if needed
      }
    });
    
    return {
      success: true,
      text: result.data.text,
      confidence: result.data.confidence,
    };
  } catch (error) {
    logger.error('OCR error', { error: error.message });
    return { success: false, error: error.message };
  }
});

// IPC: Clean up temp files
ipcMain.handle('cleanup-temp-files', async () => {
  try {
    if (fs.existsSync(TEMP_FOLDER)) {
      const files = fs.readdirSync(TEMP_FOLDER);
      for (const file of files) {
        fs.unlinkSync(path.join(TEMP_FOLDER, file));
      }
    }
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// ============================================================================
// CLIPBOARD AND UTILITY IPC HANDLERS
// ============================================================================

// Copy to clipboard
ipcMain.handle('copy-to-clipboard', async (event, text) => {
  try {
    clipboard.writeText(text);
    logger.info('Copied to clipboard', { length: text.length });
    return { success: true };
  } catch (error) {
    logger.error('Failed to copy to clipboard', { error: error.message });
    return { success: false, error: error.message };
  }
});

// Get log file path
ipcMain.handle('get-log-path', async () => {
  return LOG_FILE;
});

// Read recent logs
ipcMain.handle('get-recent-logs', async (event, lines = 100) => {
  try {
    if (!fs.existsSync(LOG_FILE)) {
      return { success: true, logs: [] };
    }
    const content = fs.readFileSync(LOG_FILE, 'utf8');
    const allLines = content.split('\n').filter(l => l.trim());
    const recentLines = allLines.slice(-lines);
    return { success: true, logs: recentLines };
  } catch (error) {
    return { success: false, error: error.message, logs: [] };
  }
});

// ============================================================================
// APP LIFECYCLE
// ============================================================================

app.whenReady().then(async () => {
  logger.info('App starting...');
  
  // Set up crash reporting (if enabled)
  setupCrashReporting();
  logger.info('Crash reporting initialized');
  
  // Configure spell checker
  const { session } = require('electron');
  session.defaultSession.setSpellCheckerLanguages(['en-US', 'en-GB']);
  session.defaultSession.setSpellCheckerEnabled(true);
  logger.info('Spell checker enabled', { languages: ['en-US', 'en-GB'] });
  
  // Load AI settings
  loadSettings();
  logger.info('AI Settings loaded', { provider: aiSettings.provider, model: aiSettings.model });
  
  // Initialize COM integration first
  const comInitialized = await initializeOutlookCOM();
  logger.info('COM initialization', { success: comInitialized });
  
  createWindow();
  createTray();
  
  // Register global hotkeys
  registerGlobalHotkeys();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
  
  logger.info('App ready');
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    // Don't quit, stay in tray
  }
});

app.on('before-quit', () => {
  app.isQuitting = true;
  unregisterGlobalHotkeys();
  logger.info('App quitting');
});

app.on('will-quit', () => {
  unregisterGlobalHotkeys();
});

// Handle uncaught exceptions
process.on('uncaughtException', (error) => {
  console.error('Uncaught Exception:', error);
});

console.log('Grok-Outlook Companion starting...');
console.log('Development mode:', isDev);
