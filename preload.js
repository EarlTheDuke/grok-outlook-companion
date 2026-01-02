const { contextBridge, ipcRenderer } = require('electron');

// Expose protected methods to renderer process
contextBridge.exposeInMainWorld('electronAPI', {
  // ========== Outlook COM Operations ==========
  
  // Get the currently selected/active email from Outlook
  getActiveEmail: () => ipcRenderer.invoke('get-active-email'),
  
  // Get recent emails from Inbox (options: { maxItems, daysBack })
  getRecentInbox: (options = {}) => ipcRenderer.invoke('get-recent-inbox', options),
  
  // Get recent emails from Sent Items (options: { maxItems, daysBack })
  getRecentSent: (options = {}) => ipcRenderer.invoke('get-recent-sent', options),
  
  // Get full email by EntryID
  getEmailById: (entryId) => ipcRenderer.invoke('get-email-by-id', { entryId }),
  
  // Check if Outlook is running
  checkOutlookStatus: () => ipcRenderer.invoke('check-outlook-status'),
  
  // Create a reply in Outlook with AI content
  createReplyInOutlook: (entryId, replyBody, replyAll = false) => 
    ipcRenderer.invoke('create-reply-in-outlook', { entryId, replyBody, replyAll }),
  
  // Create a new email in Outlook
  createNewEmail: (to, subject, body) => 
    ipcRenderer.invoke('create-new-email', { to, subject, body }),
  
  // ========== File Analysis ==========
  
  // Get attachments from an email
  getEmailAttachments: (entryId) => 
    ipcRenderer.invoke('get-email-attachments', { entryId }),
  
  // ========== Contact Lookup & Research ==========
  lookupContact: (email) => 
    ipcRenderer.invoke('lookup-contact', { email }),
  
  getEmailHistory: (email) => 
    ipcRenderer.invoke('get-email-history', { email }),
  
  researchContact: (name, email, company, title) => 
    ipcRenderer.invoke('research-contact', { name, email, company, title }),
  
  saveContactNotes: (email, notes, append = true) => 
    ipcRenderer.invoke('save-contact-notes', { email, notes, append }),
  
  // Read and extract text from a file
  readFileContent: (filePath) => 
    ipcRenderer.invoke('read-file-content', { filePath }),
  
  // Analyze file with AI
  analyzeFile: (filePath, analysisType, useOcr = false) => 
    ipcRenderer.invoke('analyze-file', { filePath, analysisType, useOcr }),
  
  // Run OCR on an image
  runOcr: (filePath) => 
    ipcRenderer.invoke('run-ocr', { filePath }),
  
  // Open file dialog for user selection
  openFileDialog: () => 
    ipcRenderer.invoke('open-file-dialog'),
  
  // Clean up temporary files
  cleanupTempFiles: () => 
    ipcRenderer.invoke('cleanup-temp-files'),
  
  // ========== Notifications ==========
  showNotification: (title, body) => ipcRenderer.invoke('show-notification', { title, body }),
  
  // ========== AI Settings ==========
  getAISettings: () => ipcRenderer.invoke('get-ai-settings'),
  saveAISettings: (settings) => ipcRenderer.invoke('save-ai-settings', settings),
  
  // ========== AI Processing ==========
  processWithAI: (prompt, emailData) => 
    ipcRenderer.invoke('process-with-ai', { prompt, emailData }),
  testAIConnection: () => ipcRenderer.invoke('test-ai-connection'),
  
  // ========== Prompt Library ==========
  getPrompts: () => ipcRenderer.invoke('get-prompts'),
  addPrompt: (prompt) => ipcRenderer.invoke('add-prompt', prompt),
  updatePrompt: (id, updates) => ipcRenderer.invoke('update-prompt', { id, updates }),
  deletePrompt: (id) => ipcRenderer.invoke('delete-prompt', { id }),
  processWithCustomPrompt: (promptId, emailData) => 
    ipcRenderer.invoke('process-with-custom-prompt', { promptId, emailData }),
  processWithMultiPrompts: (promptIds, emailData, quickNotes) =>
    ipcRenderer.invoke('process-with-multi-prompts', { promptIds, emailData, quickNotes }),
  smartShot: (emailData, promptIds, quickNotes) =>
    ipcRenderer.invoke('smart-shot', { emailData, promptIds, quickNotes }),
  exportPrompts: () => ipcRenderer.invoke('export-prompts'),
  importPrompts: (jsonData) => ipcRenderer.invoke('import-prompts', { jsonData }),
  
  // ========== Prompt Presets ==========
  getPresets: () => ipcRenderer.invoke('get-presets'),
  savePreset: (name, promptIds) => ipcRenderer.invoke('save-preset', { name, promptIds }),
  deletePreset: (id) => ipcRenderer.invoke('delete-preset', { id }),
  
  // ========== Crash Reporting (Telemetry) ==========
  getCrashReports: () => ipcRenderer.invoke('get-crash-reports'),
  clearCrashReports: () => ipcRenderer.invoke('clear-crash-reports'),
  exportCrashReports: () => ipcRenderer.invoke('export-crash-reports'),
  reportRendererError: (error, componentStack, url) => 
    ipcRenderer.invoke('report-renderer-error', { error, componentStack, url }),
  
  // ========== Clipboard ==========
  copyToClipboard: (text) => ipcRenderer.invoke('copy-to-clipboard', text),
  
  // ========== Logging ==========
  getLogPath: () => ipcRenderer.invoke('get-log-path'),
  getRecentLogs: (lines = 100) => ipcRenderer.invoke('get-recent-logs', lines),
  
  // ========== Event Listeners ==========
  
  // Listen for hotkey events from main process
  onHotkeyActiveEmail: (callback) => {
    ipcRenderer.on('hotkey-active-email', callback);
    return () => ipcRenderer.removeListener('hotkey-active-email', callback);
  },
  
  // Hotkey: Fetch active email (Ctrl+Shift+G)
  onHotkeyFetchActive: (callback) => {
    ipcRenderer.on('hotkey-fetch-active', callback);
    return () => ipcRenderer.removeListener('hotkey-fetch-active', callback);
  },
  
  // Hotkey: Quick summarize (Ctrl+Shift+S)
  onHotkeyQuickSummarize: (callback) => {
    ipcRenderer.on('hotkey-quick-summarize', callback);
    return () => ipcRenderer.removeListener('hotkey-quick-summarize', callback);
  },
  
  // Hotkey: Quick draft (Ctrl+Shift+D)
  onHotkeyQuickDraft: (callback) => {
    ipcRenderer.on('hotkey-quick-draft', callback);
    return () => ipcRenderer.removeListener('hotkey-quick-draft', callback);
  },
  
  onNavigateSettings: (callback) => {
    ipcRenderer.on('navigate-settings', callback);
    return () => ipcRenderer.removeListener('navigate-settings', callback);
  }
});

console.log('Preload script loaded - electronAPI exposed');

