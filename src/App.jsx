import React, { useState, useEffect, useCallback } from 'react';
import {
  Box,
  Container,
  Paper,
  Typography,
  Button,
  Tabs,
  Tab,
  Chip,
  Stack,
  IconButton,
  Tooltip,
  Alert,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  CircularProgress,
  Collapse,
  Divider,
  Slider,
  Menu,
  MenuItem,
  ListItemIcon,
  ListItemText,
  Dialog,
  DialogTitle,
  DialogContent,
} from '@mui/material';
import {
  Email as EmailIcon,
  Inbox as InboxIcon,
  Send as SendIcon,
  Settings as SettingsIcon,
  AutoAwesome as AIIcon,
  Refresh as RefreshIcon,
  Circle as CircleIcon,
  ExpandMore as ExpandMoreIcon,
  ExpandLess as ExpandLessIcon,
  AttachFile as AttachmentIcon,
  Star as StarIcon,
  ContentCopy as CopyIcon,
  Summarize as SummarizeIcon,
  Reply as ReplyIcon,
  Lightbulb as InsightIcon,
  Keyboard as KeyboardIcon,
  Close as CloseIcon,
  OpenInNew as OpenInOutlookIcon,
  FolderOpen as FilesIcon,
  ClearAll as ClearAllIcon,
  TextSnippet as PromptsIcon,
  Person as PersonIcon,
  Search as SearchIcon,
} from '@mui/icons-material';
import Settings from './components/Settings';
import FilesTab from './components/FilesTab';
import PromptsTab from './components/PromptsTab';
import PromptSelector from './components/PromptSelector';
import ContactCard from './components/ContactCard';
import InboxTab from './components/InboxTab';

// Tab Panel component
function TabPanel({ children, value, index, ...other }) {
  return (
    <div
      role="tabpanel"
      hidden={value !== index}
      id={`tabpanel-${index}`}
      aria-labelledby={`tab-${index}`}
      {...other}
      style={{ height: '100%', overflow: 'auto' }}
    >
      {value === index && (
        <Box sx={{ pt: 3, height: '100%' }}>
          {children}
        </Box>
      )}
    </div>
  );
}

// Email list component
function EmailTable({ emails, onSelectEmail, selectedId, loading }) {
  if (loading) {
    return (
      <Box sx={{ display: 'flex', justifyContent: 'center', py: 4 }}>
        <CircularProgress sx={{ color: 'primary.main' }} />
      </Box>
    );
  }

  if (!emails || emails.length === 0) {
    return (
      <Paper
        sx={{
          p: 4,
          background: '#1a1a1f',
          border: '1px dashed #3f3f46',
          textAlign: 'center',
          borderRadius: 2,
        }}
      >
        <EmailIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
        <Typography variant="h6" sx={{ color: 'text.secondary', mb: 1 }}>
          No Emails Found
        </Typography>
        <Typography variant="body2" sx={{ color: '#71717a' }}>
          Click the fetch button to load emails from Outlook.
        </Typography>
      </Paper>
    );
  }

  return (
    <TableContainer 
      component={Paper} 
      sx={{ 
        background: '#1a1a1f',
        maxHeight: 400,
        '& .MuiTableCell-root': {
          borderColor: '#27272a',
        }
      }}
    >
      <Table stickyHeader size="small">
        <TableHead>
          <TableRow>
            <TableCell sx={{ background: '#0f0f12', fontWeight: 600 }}>Subject</TableCell>
            <TableCell sx={{ background: '#0f0f12', fontWeight: 600 }}>From</TableCell>
            <TableCell sx={{ background: '#0f0f12', fontWeight: 600 }}>Date</TableCell>
            <TableCell sx={{ background: '#0f0f12', fontWeight: 600, width: 50 }}></TableCell>
          </TableRow>
        </TableHead>
        <TableBody>
          {emails.map((email, index) => (
            <TableRow
              key={email.entryId || index}
              onClick={() => onSelectEmail(email)}
              sx={{
                cursor: 'pointer',
                background: selectedId === email.entryId ? 'rgba(249, 115, 22, 0.1)' : 'transparent',
                '&:hover': {
                  background: selectedId === email.entryId 
                    ? 'rgba(249, 115, 22, 0.15)' 
                    : 'rgba(255, 255, 255, 0.03)',
                },
              }}
            >
              <TableCell>
                <Stack direction="row" alignItems="center" spacing={1}>
                  {!email.isRead && (
                    <CircleIcon sx={{ fontSize: 8, color: 'primary.main' }} />
                  )}
                  <Typography 
                    variant="body2" 
                    sx={{ 
                      fontWeight: email.isRead ? 400 : 600,
                      maxWidth: 300,
                      overflow: 'hidden',
                      textOverflow: 'ellipsis',
                      whiteSpace: 'nowrap',
                    }}
                  >
                    {email.subject || '(No Subject)'}
                  </Typography>
                </Stack>
              </TableCell>
              <TableCell>
                <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                  {email.senderName || email.senderEmail || 'Unknown'}
                </Typography>
              </TableCell>
              <TableCell>
                <Typography variant="body2" sx={{ color: 'text.secondary', fontSize: '0.8rem' }}>
                  {formatDate(email.receivedTime || email.sentOn)}
                </Typography>
              </TableCell>
              <TableCell>
                <Stack direction="row" spacing={0.5}>
                  {email.hasAttachments && (
                    <AttachmentIcon sx={{ fontSize: 16, color: '#71717a' }} />
                  )}
                  {email.importance === 'High' && (
                    <StarIcon sx={{ fontSize: 16, color: '#f59e0b' }} />
                  )}
                </Stack>
              </TableCell>
            </TableRow>
          ))}
        </TableBody>
      </Table>
    </TableContainer>
  );
}

// Format date helper
function formatDate(dateString) {
  if (!dateString) return '';
  try {
    const date = new Date(dateString);
    const now = new Date();
    const diffDays = Math.floor((now - date) / (1000 * 60 * 60 * 24));
    
    if (diffDays === 0) {
      return date.toLocaleTimeString('en-US', { hour: '2-digit', minute: '2-digit' });
    } else if (diffDays < 7) {
      return date.toLocaleDateString('en-US', { weekday: 'short', hour: '2-digit', minute: '2-digit' });
    } else {
      return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    }
  } catch {
    return dateString;
  }
}

// Email detail viewer component
function EmailViewer({ email, onClose, onContactClick }) {
  const [expanded, setExpanded] = useState(true);

  const copyToClipboard = (text) => {
    navigator.clipboard.writeText(text);
  };

  if (!email) return null;

  return (
    <Paper
      sx={{
        p: 3,
        background: 'linear-gradient(135deg, #1a1a1f 0%, #151518 100%)',
        border: '1px solid #27272a',
        borderRadius: 2,
        mt: 2,
      }}
    >
      <Stack direction="row" justifyContent="space-between" alignItems="flex-start" mb={2}>
        <Box>
          <Typography variant="h6" sx={{ fontWeight: 600, mb: 0.5 }}>
            {email.subject || '(No Subject)'}
          </Typography>
          <Stack direction="row" spacing={2} alignItems="center">
            <Typography variant="body2" sx={{ color: 'text.secondary' }}>
              From: <span style={{ color: '#f97316' }}>{email.senderName || email.senderEmail}</span>
            </Typography>
            <Typography variant="body2" sx={{ color: 'text.secondary' }}>
              {formatDate(email.receivedTime)}
            </Typography>
          </Stack>
          {email.to && (
            <Typography variant="body2" sx={{ color: 'text.secondary', mt: 0.5 }}>
              To: {email.to}
            </Typography>
          )}
        </Box>
        <Stack direction="row" spacing={1}>
          <Tooltip title="View contact profile & research">
            <IconButton 
              size="small" 
              onClick={onContactClick}
              sx={{ color: 'text.secondary', '&:hover': { color: '#3b82f6' } }}
            >
              <PersonIcon fontSize="small" />
            </IconButton>
          </Tooltip>
          <Tooltip title="Copy body to clipboard">
            <IconButton 
              size="small" 
              onClick={() => copyToClipboard(email.body)}
              sx={{ color: 'text.secondary', '&:hover': { color: 'primary.main' } }}
            >
              <CopyIcon fontSize="small" />
            </IconButton>
          </Tooltip>
          <IconButton 
            size="small" 
            onClick={() => setExpanded(!expanded)}
            sx={{ color: 'text.secondary' }}
          >
            {expanded ? <ExpandLessIcon /> : <ExpandMoreIcon />}
          </IconButton>
        </Stack>
      </Stack>
      
      <Collapse in={expanded}>
        <Divider sx={{ mb: 2, borderColor: '#27272a' }} />
        <Box
          sx={{
            p: 2,
            background: '#0f0f12',
            borderRadius: 1,
            maxHeight: 300,
            overflow: 'auto',
            fontFamily: '"JetBrains Mono", monospace',
            fontSize: '0.85rem',
            lineHeight: 1.6,
            whiteSpace: 'pre-wrap',
            wordBreak: 'break-word',
          }}
        >
          {email.body || '(No content)'}
        </Box>
      </Collapse>

      {email.hasAttachments && (
        <Stack direction="row" spacing={1} mt={2} alignItems="center">
          <AttachmentIcon sx={{ fontSize: 16, color: '#71717a' }} />
          <Typography variant="body2" sx={{ color: 'text.secondary' }}>
            {email.attachmentCount || 'Has'} attachment(s)
          </Typography>
        </Stack>
      )}
    </Paper>
  );
}

function App() {
  const [tabValue, setTabValue] = useState(0);
  const [status, setStatus] = useState({
    outlook: 'pending',
    ai: 'offline',
  });
  const [message, setMessage] = useState(null);
  const [loading, setLoading] = useState({
    active: false,
    inbox: false,
    sent: false,
    ai: false,
  });
  
  // One Shot Mode
  const [oneShotMode, setOneShotMode] = useState(false);
  const [oneShotStatus, setOneShotStatus] = useState(null); // 'fetching' | 'processing' | 'sending' | null
  
  // Smart Shot Mode (includes attachment analysis)
  const [smartShotMode, setSmartShotMode] = useState(false);
  const [smartShotStatus, setSmartShotStatus] = useState(null); // 'fetching' | 'extracting' | 'analyzing' | 'drafting' | 'sending' | null
  const [attachmentSummaries, setAttachmentSummaries] = useState([]);
  
  // Quick Notes (one-time disposable prompt)
  const [quickNotes, setQuickNotes] = useState('');
  const [keepQuickNotes, setKeepQuickNotes] = useState(false);
  
  // Contact Card state
  const [showContactCard, setShowContactCard] = useState(false);
  
  // Email data states
  const [activeEmail, setActiveEmail] = useState(null);
  const [inboxEmails, setInboxEmails] = useState([]);
  const [sentEmails, setSentEmails] = useState([]);
  const [selectedEmail, setSelectedEmail] = useState(null);
  
  // Settings
  const [daysBack, setDaysBack] = useState(365);
  const [maxItems, setMaxItems] = useState(50);
  
  // AI States
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [aiSettings, setAiSettings] = useState(null);
  const [aiResult, setAiResult] = useState(null);
  const [aiMenuAnchor, setAiMenuAnchor] = useState(null);
  
  // Prompt System States
  const [allPrompts, setAllPrompts] = useState([]);
  const [selectedPromptIds, setSelectedPromptIds] = useState([]);
  const [presets, setPresets] = useState([]);
  const [favoritePrompts, setFavoritePrompts] = useState([]);
  
  // UI States
  const [shortcutsOpen, setShortcutsOpen] = useState(false);

  // Check if electronAPI is available (safely)
  const isElectron = typeof window !== 'undefined' && window.electronAPI !== undefined;

  // Load AI settings and favorite prompts on mount
  useEffect(() => {
    const loadAISettings = async () => {
      if (!isElectron) return;
      
      try {
        const settings = await window.electronAPI.getAISettings();
        setAiSettings(settings);
        setStatus(prev => ({
          ...prev,
          ai: settings.hasApiKey || settings.provider === 'ollama' ? 'online' : 'offline',
        }));
      } catch (error) {
        console.error('Error loading AI settings:', error);
      }
    };
    
    const loadPromptsAndPresets = async () => {
      if (!isElectron) return;
      
      try {
        // Load all prompts
        const promptsResult = await window.electronAPI.getPrompts();
        if (promptsResult.success) {
          setAllPrompts(promptsResult.prompts);
          setFavoritePrompts(promptsResult.prompts.filter(p => p.isFavorite));
        }
        
        // Load presets
        const presetsResult = await window.electronAPI.getPresets();
        if (presetsResult.success) {
          setPresets(presetsResult.presets);
        }
      } catch (error) {
        console.error('Error loading prompts/presets:', error);
      }
    };
    
    loadAISettings();
    loadPromptsAndPresets();
  }, [isElectron]);

  // Check Outlook status on mount
  useEffect(() => {
    const checkStatus = async () => {
      if (!isElectron) {
        setStatus(prev => ({ ...prev, outlook: 'offline' }));
        return;
      }
      
      try {
        const result = await window.electronAPI.checkOutlookStatus();
        setStatus(prev => ({
          ...prev,
          outlook: result && result.running ? 'online' : 'offline',
        }));
      } catch (error) {
        console.error('Error checking Outlook status:', error);
        setStatus(prev => ({ ...prev, outlook: 'offline' }));
      }
    };
    
    // Delay initial check to ensure everything is loaded
    const timeout = setTimeout(checkStatus, 1000);
    // Check status every 30 seconds
    const interval = setInterval(checkStatus, 30000);
    return () => {
      clearTimeout(timeout);
      clearInterval(interval);
    };
  }, [isElectron]);

  // Helper function to fetch and process with AI
  const fetchAndProcess = useCallback(async (action) => {
    if (!isElectron) return;
    
    setLoading(prev => ({ ...prev, active: true }));
    try {
      const result = await window.electronAPI.getActiveEmail();
      if (result.success) {
        setActiveEmail(result.data);
        setTabValue(0); // Switch to active email tab
        
        // Now process with AI
        if (action && (aiSettings?.hasApiKey || aiSettings?.provider === 'ollama')) {
          setLoading(prev => ({ ...prev, active: false, ai: true }));
          
          let prompt;
          switch (action) {
            case 'summarize':
              prompt = 'Please provide a concise summary of this email, highlighting the key points and any action items.';
              break;
            case 'reply':
              prompt = 'Please draft a professional reply to this email. Keep it concise and address the main points.';
              break;
            default:
              prompt = 'Please summarize this email.';
          }
          
          const aiResult = await window.electronAPI.processWithAI(prompt, result.data);
          if (aiResult.success) {
            setAiResult({
              type: action,
              content: aiResult.content,
              email: result.data.subject,
            });
            await window.electronAPI.showNotification(
              'Grok-Outlook', 
              `${action === 'summarize' ? 'Summary' : 'Draft'} ready! Click to view.`
            );
          }
          setLoading(prev => ({ ...prev, ai: false }));
        }
      }
    } catch (error) {
      console.error('Error in fetchAndProcess:', error);
    } finally {
      setLoading(prev => ({ ...prev, active: false }));
    }
  }, [isElectron, aiSettings]);

  // Listen for hotkey events
  useEffect(() => {
    if (!isElectron) return undefined;
    
    const cleanups = [];
    
    try {
      // Ctrl+Shift+G - Fetch active email
      cleanups.push(
        window.electronAPI.onHotkeyFetchActive(() => {
          console.log('Hotkey received: Fetch Active');
          fetchAndProcess(null);
        })
      );
      
      // Ctrl+Shift+S - Quick summarize
      cleanups.push(
        window.electronAPI.onHotkeyQuickSummarize(() => {
          console.log('Hotkey received: Quick Summarize');
          fetchAndProcess('summarize');
        })
      );
      
      // Ctrl+Shift+D - Quick draft
      cleanups.push(
        window.electronAPI.onHotkeyQuickDraft(() => {
          console.log('Hotkey received: Quick Draft');
          fetchAndProcess('reply');
        })
      );
      
      // Navigate to settings
      cleanups.push(
        window.electronAPI.onNavigateSettings(() => {
          setSettingsOpen(true);
        })
      );
      
      return () => {
        cleanups.forEach(cleanup => {
          if (typeof cleanup === 'function') cleanup();
        });
      };
    } catch (error) {
      console.error('Error setting up hotkey listeners:', error);
      return undefined;
    }
  }, [isElectron, fetchAndProcess]);

  const handleTabChange = (event, newValue) => {
    setTabValue(newValue);
    setSelectedEmail(null);
  };

  const showMessage = useCallback((type, text, duration = 5000) => {
    setMessage({ type, text });
    if (duration > 0) {
      setTimeout(() => setMessage(null), duration);
    }
  }, []);

  // Clear all content in the app
  const handleClearAll = useCallback(() => {
    setActiveEmail(null);
    setSelectedEmail(null);
    setAiResult(null);
    setInboxEmails([]);
    setSentEmails([]);
    setQuickNotes('');
    setAttachmentSummaries([]);
    showMessage('info', 'All content cleared');
  }, [showMessage]);

  const handleFetchActiveEmail = async () => {
    if (!isElectron) {
      showMessage('warning', 'Running in browser - Electron API not available');
      return;
    }

    // Clear previous results before fetching new email
    setActiveEmail(null);
    setAiResult(null);
    
    setLoading(prev => ({ ...prev, active: true }));
    
    try {
      const result = await window.electronAPI.getActiveEmail();
      
      if (result.success) {
        setActiveEmail(result.data);
        showMessage('success', `Email loaded: "${result.data.subject}"`);
      } else {
        showMessage('warning', result.message);
        setActiveEmail(null);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
      setActiveEmail(null);
    } finally {
      setLoading(prev => ({ ...prev, active: false }));
    }
  };

  // One Shot: Fetch → AI → Outlook in one click
  const handleOneShot = async () => {
    if (!isElectron) {
      showMessage('warning', 'Running in browser - Electron API not available');
      return;
    }

    if (!aiSettings?.hasApiKey && aiSettings?.provider !== 'ollama') {
      showMessage('warning', 'Please configure your AI settings first (click the gear icon)');
      return;
    }

    // Clear previous results
    setActiveEmail(null);
    setAiResult(null);
    
    try {
      // Step 1: Fetch Active Email
      setOneShotStatus('fetching');
      setLoading(prev => ({ ...prev, active: true }));
      
      const emailResult = await window.electronAPI.getActiveEmail();
      
      if (!emailResult.success) {
        showMessage('warning', emailResult.message || 'No email selected in Outlook');
        return;
      }
      
      const email = emailResult.data;
      setActiveEmail(email);
      showMessage('info', `⚡ One Shot: Got "${email.subject}" - Processing...`);
      
      // Step 2: Run AI with selected prompts
      setOneShotStatus('processing');
      setLoading(prev => ({ ...prev, active: false, ai: true }));
      
      const aiResult = await window.electronAPI.processWithMultiPrompts(selectedPromptIds, email, quickNotes);
      
      if (!aiResult.success) {
        showMessage('error', aiResult.error || 'AI processing failed');
        return;
      }
      
      setAiResult({
        type: 'multi',
        content: aiResult.content,
        email: email.subject,
        usedPrompts: aiResult.usedPrompts || [],
        hadQuickNotes: aiResult.hadQuickNotes,
      });
      
      // Clear quick notes after use (unless "keep" is checked)
      if (!keepQuickNotes) {
        setQuickNotes('');
      }
      
      showMessage('info', '⚡ One Shot: AI complete - Opening in Outlook...');
      
      // Step 3: Send to Outlook
      setOneShotStatus('sending');
      
      if (!email.entryId) {
        showMessage('warning', 'Cannot find original email to reply to');
        return;
      }
      
      const outlookResult = await window.electronAPI.createReplyInOutlook(
        email.entryId,
        aiResult.content,
        true // reply all - includes CC recipients
      );
      
      if (outlookResult.success) {
        showMessage('success', '⚡ One Shot Complete! Reply opened in Outlook.');
      } else {
        showMessage('error', outlookResult.message || 'Failed to open in Outlook');
      }
      
    } catch (error) {
      console.error('One Shot error:', error);
      showMessage('error', `One Shot failed: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, active: false, ai: false }));
      setOneShotStatus(null);
    }
  };

  // Smart Shot: Fetch → Extract Attachments → Analyze → AI → Outlook
  const handleSmartShot = async () => {
    if (!isElectron) {
      showMessage('warning', 'Running in browser - Electron API not available');
      return;
    }

    if (!aiSettings?.hasApiKey && aiSettings?.provider !== 'ollama') {
      showMessage('warning', 'Please configure your AI settings first (click the gear icon)');
      return;
    }

    // Clear previous results
    setActiveEmail(null);
    setAiResult(null);
    setAttachmentSummaries([]);
    
    try {
      // Step 1: Fetch Active Email
      setSmartShotStatus('fetching');
      setLoading(prev => ({ ...prev, active: true }));
      
      const emailResult = await window.electronAPI.getActiveEmail();
      
      if (!emailResult.success) {
        showMessage('warning', emailResult.message || 'No email selected in Outlook');
        return;
      }
      
      const email = emailResult.data;
      setActiveEmail(email);
      showMessage('info', `🧠 Smart Shot: Got "${email.subject}" - Analyzing attachments...`);
      
      // Step 2: Run Smart Shot (handles attachment extraction, analysis, and AI)
      setSmartShotStatus('extracting');
      setLoading(prev => ({ ...prev, active: false, ai: true }));
      
      // Update status as processing progresses
      setSmartShotStatus('analyzing');
      
      const result = await window.electronAPI.smartShot(email, selectedPromptIds, quickNotes);
      
      if (!result.success) {
        showMessage('error', result.error || 'Smart Shot failed');
        return;
      }
      
      // Store attachment summaries
      if (result.attachmentSummaries && result.attachmentSummaries.length > 0) {
        setAttachmentSummaries(result.attachmentSummaries);
        showMessage('info', `🧠 Smart Shot: Analyzed ${result.attachmentSummaries.length} attachment(s)`);
      } else {
        showMessage('info', '🧠 Smart Shot: No attachments found - using email only');
      }
      
      // Store AI result
      if (result.aiResult?.success) {
        setAiResult({
          type: 'smart-shot',
          content: result.aiResult.content,
          email: email.subject,
          usedPrompts: result.usedPrompts || [],
          hadQuickNotes: result.hadQuickNotes,
          attachmentCount: result.attachmentSummaries?.length || 0,
        });
      } else {
        showMessage('error', result.aiResult?.error || 'AI processing failed');
        return;
      }
      
      // Clear quick notes after use (unless "keep" is checked)
      if (!keepQuickNotes) {
        setQuickNotes('');
      }
      
      showMessage('info', '🧠 Smart Shot: AI complete - Opening in Outlook...');
      
      // Step 3: Send to Outlook
      setSmartShotStatus('sending');
      
      if (!email.entryId) {
        showMessage('warning', 'Cannot find original email to reply to');
        return;
      }
      
      const outlookResult = await window.electronAPI.createReplyInOutlook(
        email.entryId,
        result.aiResult.content,
        true // reply all - includes CC recipients
      );
      
      if (outlookResult.success) {
        showMessage('success', '🧠 Smart Shot Complete! Reply opened in Outlook.');
      } else {
        showMessage('error', outlookResult.message || 'Failed to open in Outlook');
      }
      
    } catch (error) {
      console.error('Smart Shot error:', error);
      showMessage('error', `Smart Shot failed: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, active: false, ai: false }));
      setSmartShotStatus(null);
    }
  };

  const handleFetchInbox = async () => {
    if (!isElectron) {
      showMessage('warning', 'Running in browser - Electron API not available');
      return;
    }

    setLoading(prev => ({ ...prev, inbox: true }));
    
    try {
      const result = await window.electronAPI.getRecentInbox({ maxItems, daysBack });
      
      if (result.success) {
        setInboxEmails(result.data);
        showMessage('success', `Loaded ${result.count} emails from Inbox`);
      } else {
        showMessage('warning', result.message);
        setInboxEmails([]);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
      setInboxEmails([]);
    } finally {
      setLoading(prev => ({ ...prev, inbox: false }));
    }
  };

  const handleFetchSent = async () => {
    if (!isElectron) {
      showMessage('warning', 'Running in browser - Electron API not available');
      return;
    }

    setLoading(prev => ({ ...prev, sent: true }));
    
    try {
      const result = await window.electronAPI.getRecentSent({ maxItems, daysBack });
      
      if (result.success) {
        setSentEmails(result.data);
        showMessage('success', `Loaded ${result.count} emails from Sent Items`);
      } else {
        showMessage('warning', result.message);
        setSentEmails([]);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
      setSentEmails([]);
    } finally {
      setLoading(prev => ({ ...prev, sent: false }));
    }
  };

  const handleSelectEmail = async (email) => {
    setSelectedEmail(email);
    
    // If we only have a preview, fetch the full email
    if (email.entryId && !email.body && isElectron) {
      try {
        const result = await window.electronAPI.getEmailById(email.entryId);
        if (result.success) {
          setSelectedEmail(result.data);
        }
      } catch (error) {
        console.error('Error fetching full email:', error);
      }
    }
  };

  // AI Processing Functions
  const handleSaveSettings = async (settings) => {
    if (!isElectron) return;
    
    try {
      const result = await window.electronAPI.saveAISettings(settings);
      if (result.success) {
        const newSettings = await window.electronAPI.getAISettings();
        setAiSettings(newSettings);
        setStatus(prev => ({
          ...prev,
          ai: newSettings.hasApiKey || newSettings.provider === 'ollama' ? 'online' : 'offline',
        }));
      }
    } catch (error) {
      console.error('Error saving settings:', error);
      throw error;
    }
  };

  // Preset handlers
  const handleSavePreset = async (name, promptIds) => {
    if (!isElectron) return;
    
    try {
      const result = await window.electronAPI.savePreset(name, promptIds);
      if (result.success) {
        showMessage('success', `Preset "${name}" saved!`);
        // Refresh presets
        const presetsResult = await window.electronAPI.getPresets();
        if (presetsResult.success) {
          setPresets(presetsResult.presets);
        }
      } else {
        showMessage('error', result.error);
      }
    } catch (error) {
      showMessage('error', `Error saving preset: ${error.message}`);
    }
  };

  const handleLoadPreset = (preset) => {
    setSelectedPromptIds(preset.promptIds);
    showMessage('info', `Loaded preset: ${preset.name}`);
  };

  const handleDeletePreset = async (presetId) => {
    if (!isElectron) return;
    
    try {
      const result = await window.electronAPI.deletePreset(presetId);
      if (result.success) {
        showMessage('success', 'Preset deleted');
        setPresets(prev => prev.filter(p => p.id !== presetId));
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Handle multi-prompt AI action (new main handler)
  const handleMultiPromptAction = async () => {
    setAiMenuAnchor(null);
    
    const email = activeEmail || selectedEmail;
    if (!email) {
      showMessage('warning', 'Please select an email first');
      return;
    }
    
    if (!aiSettings?.hasApiKey && aiSettings?.provider !== 'ollama') {
      showMessage('warning', 'Please configure your AI settings first (click the gear icon)');
      return;
    }

    setLoading(prev => ({ ...prev, ai: true }));
    setAiResult(null);
    
    try {
      const result = await window.electronAPI.processWithMultiPrompts(selectedPromptIds, email, quickNotes);
      
      if (result.success) {
        setAiResult({
          type: 'multi',
          content: result.content,
          email: email.subject,
          usedPrompts: result.usedPrompts || [],
          hadQuickNotes: result.hadQuickNotes,
        });
        
        // Clear quick notes after use (unless "keep" is checked)
        if (!keepQuickNotes) {
          setQuickNotes('');
        }
        
        const promptCount = selectedPromptIds.length || 0;
        showMessage('success', promptCount > 0 
          ? `AI completed with ${promptCount} prompt(s)!` 
          : 'AI analysis completed!');
        
        // Refresh prompts to update usage counts
        const promptsResult = await window.electronAPI.getPrompts();
        if (promptsResult.success) {
          setAllPrompts(promptsResult.prompts);
          setFavoritePrompts(promptsResult.prompts.filter(p => p.isFavorite));
        }
      } else {
        showMessage('error', result.error || 'Failed to process');
      }
    } catch (error) {
      console.error('Error with multi-prompt:', error);
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, ai: false }));
    }
  };

  // Handle custom prompt execution (legacy - for menu favorites)
  const handleCustomPrompt = async (promptId) => {
    // Add to selected and run
    setSelectedPromptIds([promptId]);
    setAiMenuAnchor(null);
    
    const email = activeEmail || selectedEmail;
    if (!email) {
      showMessage('warning', 'Please select an email first');
      return;
    }
    
    if (!aiSettings?.hasApiKey && aiSettings?.provider !== 'ollama') {
      showMessage('warning', 'Please configure your AI settings first (click the gear icon)');
      return;
    }

    setLoading(prev => ({ ...prev, ai: true }));
    setAiResult(null);
    
    try {
      const result = await window.electronAPI.processWithMultiPrompts([promptId], email);
      
      if (result.success) {
        setAiResult({
          type: 'multi',
          content: result.content,
          email: email.subject,
          usedPrompts: result.usedPrompts || [],
        });
        showMessage('success', 'AI completed!');
        
        // Refresh prompts to update usage counts
        const promptsResult = await window.electronAPI.getPrompts();
        if (promptsResult.success) {
          setAllPrompts(promptsResult.prompts);
          setFavoritePrompts(promptsResult.prompts.filter(p => p.isFavorite));
        }
      } else {
        showMessage('error', result.error || 'Failed to process with custom prompt');
      }
    } catch (error) {
      console.error('Error with custom prompt:', error);
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, ai: false }));
    }
  };

  const handleAIAction = async (action) => {
    setAiMenuAnchor(null);
    
    const email = activeEmail || selectedEmail;
    if (!email) {
      showMessage('warning', 'Please select an email first');
      return;
    }

    if (!aiSettings?.hasApiKey && aiSettings?.provider !== 'ollama') {
      showMessage('warning', 'Please configure your AI settings first (click the gear icon)');
      setSettingsOpen(true);
      return;
    }

    setLoading(prev => ({ ...prev, ai: true }));
    setAiResult(null);

    try {
      let prompt;
      switch (action) {
        case 'summarize':
          prompt = 'Please provide a concise summary of this email, highlighting the key points and any action items.';
          break;
        case 'reply':
          prompt = 'Please draft a professional reply to this email. Keep it concise and address the main points.';
          break;
        case 'insights':
          prompt = 'Please analyze this email and provide: 1) Key topics discussed 2) Any deadlines or time-sensitive items 3) Action items required 4) Tone/sentiment of the email';
          break;
        default:
          prompt = 'Please summarize this email.';
      }

      const result = await window.electronAPI.processWithAI(prompt, email);
      
      if (result.success) {
        setAiResult({
          type: action,
          content: result.content,
          email: email.subject,
        });
        showMessage('success', `AI ${action} completed! See result above the email.`);
        
        // Scroll to AI result after a short delay
        setTimeout(() => {
          const aiPanel = document.getElementById('ai-result-panel');
          if (aiPanel) {
            aiPanel.scrollIntoView({ behavior: 'smooth', block: 'start' });
          }
        }, 100);
      } else {
        showMessage('error', result.error || 'AI processing failed');
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, ai: false }));
    }
  };

  const copyAIResult = () => {
    if (aiResult?.content) {
      navigator.clipboard.writeText(aiResult.content);
      showMessage('success', 'Copied to clipboard!');
    }
  };

  // Send AI draft reply to Outlook
  const sendDraftToOutlook = async (replyAll = false) => {
    if (!isElectron || !aiResult?.content) return;
    
    const email = activeEmail || selectedEmail;
    if (!email?.entryId) {
      showMessage('error', 'Cannot find original email to reply to');
      return;
    }

    try {
      showMessage('info', 'Opening reply in Outlook...');
      const result = await window.electronAPI.createReplyInOutlook(
        email.entryId,
        aiResult.content,
        replyAll
      );
      
      if (result.success) {
        showMessage('success', 'Reply opened in Outlook! Review and send when ready.');
      } else {
        showMessage('error', result.message);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  const StatusDot = ({ status: dotStatus }) => (
    <CircleIcon
      sx={{
        fontSize: 10,
        color: dotStatus === 'online' ? '#22c55e' : dotStatus === 'pending' ? '#f59e0b' : '#ef4444',
        animation: dotStatus === 'pending' ? 'pulse 2s infinite' : 'none',
        '@keyframes pulse': {
          '0%, 100%': { opacity: 1 },
          '50%': { opacity: 0.4 },
        },
      }}
    />
  );

  return (
    <Box
      sx={{
        minHeight: '100vh',
        background: 'linear-gradient(180deg, #0a0a0f 0%, #0f0f15 50%, #0a0a0f 100%)',
        position: 'relative',
        overflow: 'hidden',
      }}
    >
      {/* Decorative background elements */}
      <Box
        sx={{
          position: 'absolute',
          top: -200,
          right: -200,
          width: 500,
          height: 500,
          background: 'radial-gradient(circle, rgba(249, 115, 22, 0.08) 0%, transparent 70%)',
          pointerEvents: 'none',
        }}
      />
      <Box
        sx={{
          position: 'absolute',
          bottom: -100,
          left: -100,
          width: 400,
          height: 400,
          background: 'radial-gradient(circle, rgba(6, 182, 212, 0.06) 0%, transparent 70%)',
          pointerEvents: 'none',
        }}
      />

      {/* Header */}
      <Paper
        elevation={0}
        sx={{
          background: 'rgba(24, 24, 27, 0.8)',
          backdropFilter: 'blur(10px)',
          borderBottom: '1px solid #27272a',
          borderRadius: 0,
          position: 'sticky',
          top: 0,
          zIndex: 100,
        }}
      >
        <Container maxWidth="xl">
          <Box
            sx={{
              display: 'flex',
              alignItems: 'center',
              justifyContent: 'space-between',
              py: 2,
            }}
          >
            {/* Logo & Title */}
            <Stack direction="row" alignItems="center" spacing={2}>
              <Box
                sx={{
                  width: 40,
                  height: 40,
                  borderRadius: 2,
                  background: 'linear-gradient(135deg, #f97316 0%, #ea580c 100%)',
                  display: 'flex',
                  alignItems: 'center',
                  justifyContent: 'center',
                  boxShadow: '0 4px 12px rgba(249, 115, 22, 0.3)',
                }}
              >
                <AIIcon sx={{ color: '#000', fontSize: 24 }} />
              </Box>
              <Box>
                <Typography variant="h6" sx={{ fontWeight: 700, letterSpacing: '-0.02em' }}>
                  Grok-Outlook Companion
                </Typography>
                <Typography variant="caption" sx={{ color: 'text.secondary' }}>
                  AI-Powered Email Assistant
                </Typography>
              </Box>
            </Stack>

            {/* Status & Settings */}
            <Stack direction="row" alignItems="center" spacing={3}>
              {/* Status Chips */}
              <Stack direction="row" spacing={1}>
                <Chip
                  icon={<StatusDot status={status.outlook} />}
                  label={`Outlook ${status.outlook === 'online' ? 'Connected' : status.outlook === 'pending' ? 'Checking...' : 'Disconnected'}`}
                  size="small"
                  sx={{
                    background: '#27272a',
                    border: '1px solid #3f3f46',
                  }}
                />
                <Chip
                  icon={<StatusDot status={status.ai} />}
                  label="AI"
                  size="small"
                  sx={{
                    background: '#27272a',
                    border: '1px solid #3f3f46',
                  }}
                />
              </Stack>

              {/* Clear All Button */}
              <Button
                variant="outlined"
                startIcon={<ClearAllIcon />}
                onClick={handleClearAll}
                sx={{
                  borderColor: '#ef4444',
                  color: '#ef4444',
                  '&:hover': { 
                    borderColor: '#dc2626',
                    background: 'rgba(239, 68, 68, 0.1)',
                    color: '#dc2626',
                  },
                }}
              >
                Clear All
              </Button>

              {/* Keyboard Shortcuts Button */}
              <Tooltip title="Keyboard Shortcuts">
                <IconButton
                  onClick={() => setShortcutsOpen(true)}
                  sx={{
                    color: 'text.secondary',
                    '&:hover': { color: 'secondary.main' },
                  }}
                >
                  <KeyboardIcon />
                </IconButton>
              </Tooltip>
              
              {/* Settings Button */}
              <Tooltip title="AI Settings">
                <IconButton
                  onClick={() => setSettingsOpen(true)}
                  sx={{
                    color: 'text.secondary',
                    '&:hover': { color: 'primary.main' },
                  }}
                >
                  <SettingsIcon />
                </IconButton>
              </Tooltip>
            </Stack>
          </Box>
        </Container>
      </Paper>

      {/* Main Content */}
      <Container maxWidth="xl" sx={{ py: 4, height: 'calc(100vh - 88px)' }}>
        {/* Alert Messages */}
        {message && (
          <Alert
            severity={message.type}
            onClose={() => setMessage(null)}
            sx={{ mb: 3, animation: 'fadeIn 0.3s ease-out' }}
          >
            {message.text}
          </Alert>
        )}

        {/* Tabs Section */}
        <Paper
          sx={{
            background: 'rgba(24, 24, 27, 0.6)',
            backdropFilter: 'blur(5px)',
            borderRadius: 3,
            height: 'calc(100% - 60px)',
            display: 'flex',
            flexDirection: 'column',
            overflow: 'hidden',
          }}
        >
          {/* Tab Headers */}
          <Box sx={{ borderBottom: 1, borderColor: 'divider', px: 2 }}>
            <Tabs
              value={tabValue}
              onChange={handleTabChange}
              sx={{
                '& .MuiTab-root': {
                  minHeight: 56,
                },
                '& .Mui-selected': {
                  color: 'primary.main',
                },
                '& .MuiTabs-indicator': {
                  backgroundColor: 'primary.main',
                  height: 3,
                  borderRadius: '3px 3px 0 0',
                },
              }}
            >
              <Tab
                icon={<EmailIcon sx={{ fontSize: 20 }} />}
                iconPosition="start"
                label="Active Email"
              />
              <Tab
                icon={<InboxIcon sx={{ fontSize: 20 }} />}
                iconPosition="start"
                label="Recent Inbox"
              />
              <Tab
                icon={<SendIcon sx={{ fontSize: 20 }} />}
                iconPosition="start"
                label={
                  <Box sx={{ textAlign: 'left' }}>
                    <span>{`Sent Items ${sentEmails.length > 0 ? `(${sentEmails.length})` : ''}`}</span>
                    <Typography variant="caption" sx={{ display: 'block', fontSize: '0.6rem', color: '#f97316', mt: -0.5 }}>
                      🚧 coming soon
                    </Typography>
                  </Box>
                }
              />
              <Tab
                icon={<FilesIcon sx={{ fontSize: 20 }} />}
                iconPosition="start"
                label="Files"
              />
              <Tab
                icon={<PromptsIcon sx={{ fontSize: 20 }} />}
                iconPosition="start"
                label="Prompts"
              />
            </Tabs>
          </Box>

          {/* Tab Panels */}
          <Box sx={{ flex: 1, overflow: 'hidden', px: 3, pb: 3 }}>
            {/* Active Email Tab */}
            <TabPanel value={tabValue} index={0}>
              <Box sx={{ display: 'flex', gap: 3, height: '100%' }}>
                {/* Main Content - Left Side */}
                <Box sx={{ flex: 1, overflow: 'auto' }}>
                  <Stack spacing={3}>
                    {/* Shot Mode Toggles */}
                    <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
                      {/* One Shot Toggle */}
                      <Box
                        sx={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: 1.5,
                          p: 1.5,
                          flex: 1,
                          minWidth: 280,
                          background: oneShotMode 
                            ? 'linear-gradient(135deg, rgba(234, 179, 8, 0.15) 0%, rgba(249, 115, 22, 0.1) 100%)'
                            : 'rgba(39, 39, 42, 0.3)',
                          border: `1px solid ${oneShotMode ? 'rgba(234, 179, 8, 0.4)' : '#27272a'}`,
                          borderRadius: 2,
                          cursor: 'pointer',
                          transition: 'all 0.2s',
                          '&:hover': {
                            borderColor: oneShotMode ? 'rgba(234, 179, 8, 0.6)' : '#3f3f46',
                          },
                        }}
                        onClick={() => {
                          setOneShotMode(!oneShotMode);
                          if (!oneShotMode) setSmartShotMode(false); // Turn off Smart Shot if turning on One Shot
                        }}
                      >
                        <Typography 
                          variant="body2" 
                          sx={{ fontWeight: 600, color: oneShotMode ? '#eab308' : '#71717a' }}
                        >
                          ⚡ One Shot
                        </Typography>
                        <Chip
                          label={oneShotMode ? 'ON' : 'OFF'}
                          size="small"
                          sx={{
                            background: oneShotMode ? '#eab308' : '#3f3f46',
                            color: oneShotMode ? '#000' : '#a1a1aa',
                            fontWeight: 700,
                          }}
                        />
                        <Typography variant="caption" sx={{ color: '#71717a', flex: 1 }}>
                          Email → AI → Outlook
                        </Typography>
                      </Box>

                      {/* Smart Shot Toggle */}
                      <Box
                        sx={{
                          display: 'flex',
                          alignItems: 'center',
                          gap: 1.5,
                          p: 1.5,
                          flex: 1,
                          minWidth: 280,
                          background: smartShotMode 
                            ? 'linear-gradient(135deg, rgba(139, 92, 246, 0.15) 0%, rgba(236, 72, 153, 0.1) 100%)'
                            : 'rgba(39, 39, 42, 0.3)',
                          border: `1px solid ${smartShotMode ? 'rgba(139, 92, 246, 0.4)' : '#27272a'}`,
                          borderRadius: 2,
                          cursor: 'pointer',
                          transition: 'all 0.2s',
                          '&:hover': {
                            borderColor: smartShotMode ? 'rgba(139, 92, 246, 0.6)' : '#3f3f46',
                          },
                        }}
                        onClick={() => {
                          setSmartShotMode(!smartShotMode);
                          if (!smartShotMode) setOneShotMode(false); // Turn off One Shot if turning on Smart Shot
                        }}
                      >
                        <Typography 
                          variant="body2" 
                          sx={{ fontWeight: 600, color: smartShotMode ? '#8b5cf6' : '#71717a' }}
                        >
                          🧠 Smart Shot
                        </Typography>
                        <Chip
                          label={smartShotMode ? 'ON' : 'OFF'}
                          size="small"
                          sx={{
                            background: smartShotMode ? '#8b5cf6' : '#3f3f46',
                            color: smartShotMode ? '#fff' : '#a1a1aa',
                            fontWeight: 700,
                          }}
                        />
                        <Typography variant="caption" sx={{ color: '#71717a', flex: 1 }}>
                          + Attachments
                        </Typography>
                      </Box>
                    </Box>
                    
                    {/* Warning if prompts not selected */}
                    {(oneShotMode || smartShotMode) && selectedPromptIds.length === 0 && (
                      <Chip 
                        label="⚠️ Select prompts first" 
                        size="small" 
                        sx={{ background: 'rgba(239, 68, 68, 0.2)', color: '#ef4444', alignSelf: 'flex-start' }} 
                      />
                    )}

                    <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                      {/* Main Action Button - changes based on mode */}
                      {smartShotMode ? (
                        <Button
                          variant="contained"
                          startIcon={
                            smartShotStatus ? <CircularProgress size={18} color="inherit" /> : <AIIcon />
                          }
                          onClick={handleSmartShot}
                          disabled={loading.active || loading.ai || selectedPromptIds.length === 0}
                          sx={{
                            px: 4,
                            py: 1.5,
                            fontSize: '1rem',
                            fontWeight: 700,
                            background: 'linear-gradient(135deg, #8b5cf6 0%, #ec4899 100%)',
                            '&:hover': {
                              background: 'linear-gradient(135deg, #a78bfa 0%, #f472b6 100%)',
                            },
                            '&.Mui-disabled': {
                              background: '#3f3f46',
                            },
                          }}
                        >
                          {smartShotStatus === 'fetching' ? '🧠 Fetching Email...' :
                           smartShotStatus === 'extracting' ? '🧠 Extracting Attachments...' :
                           smartShotStatus === 'analyzing' ? '🧠 Analyzing Files...' :
                           smartShotStatus === 'drafting' ? '🧠 Drafting Reply...' :
                           smartShotStatus === 'sending' ? '🧠 Opening Outlook...' :
                           `🧠 SMART SHOT (${selectedPromptIds.length} prompts)`}
                        </Button>
                      ) : oneShotMode ? (
                        <Button
                          variant="contained"
                          startIcon={
                            oneShotStatus === 'fetching' ? <CircularProgress size={18} color="inherit" /> :
                            oneShotStatus === 'processing' ? <CircularProgress size={18} color="inherit" /> :
                            oneShotStatus === 'sending' ? <CircularProgress size={18} color="inherit" /> :
                            <AIIcon />
                          }
                          onClick={handleOneShot}
                          disabled={loading.active || loading.ai || selectedPromptIds.length === 0}
                          sx={{
                            px: 4,
                            py: 1.5,
                            fontSize: '1rem',
                            fontWeight: 700,
                            background: 'linear-gradient(135deg, #eab308 0%, #f97316 100%)',
                            '&:hover': {
                              background: 'linear-gradient(135deg, #facc15 0%, #fb923c 100%)',
                            },
                            '&.Mui-disabled': {
                              background: '#3f3f46',
                            },
                          }}
                        >
                          {oneShotStatus === 'fetching' ? '⚡ Fetching Email...' :
                           oneShotStatus === 'processing' ? '⚡ Running AI...' :
                           oneShotStatus === 'sending' ? '⚡ Opening Outlook...' :
                           `⚡ ONE SHOT (${selectedPromptIds.length} prompts)`}
                        </Button>
                      ) : (
                        <>
                          <Button
                            variant="contained"
                            startIcon={loading.active ? <CircularProgress size={18} color="inherit" /> : <RefreshIcon />}
                            onClick={handleFetchActiveEmail}
                            disabled={loading.active}
                            sx={{ px: 3 }}
                          >
                            {loading.active ? 'Fetching...' : 'Fetch Active Email'}
                          </Button>
                          
                          <Divider orientation="vertical" flexItem sx={{ borderColor: '#27272a' }} />
                          
                          {/* Run with Selected Prompts Button */}
                          <Button
                            variant="contained"
                            startIcon={loading.ai ? <CircularProgress size={18} color="inherit" /> : <AIIcon />}
                            disabled={!activeEmail || loading.ai}
                            onClick={handleMultiPromptAction}
                            sx={{
                              background: 'linear-gradient(135deg, #f97316 0%, #ea580c 100%)',
                              '&:hover': {
                                background: 'linear-gradient(135deg, #fb923c 0%, #f97316 100%)',
                              },
                            }}
                          >
                            {loading.ai ? 'Processing...' : `Run AI ${selectedPromptIds.length > 0 ? `(${selectedPromptIds.length} prompts)` : '(no prompts)'}`}
                          </Button>
                        </>
                      )}
                      
                      {/* Quick Actions Dropdown */}
                      <Button
                        variant="outlined"
                        size="small"
                        disabled={!activeEmail || loading.ai}
                        onClick={(e) => setAiMenuAnchor(e.currentTarget)}
                        sx={{
                          borderColor: '#3f3f46',
                          color: 'text.secondary',
                          '&:hover': {
                            borderColor: 'primary.main',
                            color: 'primary.main',
                          },
                        }}
                      >
                        Quick Actions ▾
                      </Button>
                      <Menu
                        anchorEl={aiMenuAnchor}
                        open={Boolean(aiMenuAnchor)}
                        onClose={() => setAiMenuAnchor(null)}
                        PaperProps={{
                          sx: {
                            background: '#18181b',
                            border: '1px solid #27272a',
                            minWidth: 200,
                          }
                        }}
                      >
                        <Typography variant="caption" sx={{ px: 2, py: 0.5, color: '#71717a', display: 'block' }}>
                          Built-in (ignores selection):
                        </Typography>
                        <MenuItem onClick={() => handleAIAction('summarize')}>
                          <ListItemIcon><SummarizeIcon sx={{ color: 'primary.main' }} /></ListItemIcon>
                          <ListItemText>Summarize Email</ListItemText>
                        </MenuItem>
                        <MenuItem onClick={() => handleAIAction('reply')}>
                          <ListItemIcon><ReplyIcon sx={{ color: 'secondary.main' }} /></ListItemIcon>
                          <ListItemText>Draft Reply</ListItemText>
                        </MenuItem>
                        <MenuItem onClick={() => handleAIAction('insights')}>
                          <ListItemIcon><InsightIcon sx={{ color: '#f59e0b' }} /></ListItemIcon>
                          <ListItemText>Extract Insights</ListItemText>
                        </MenuItem>
                        
                        {/* Favorite Custom Prompts */}
                        {favoritePrompts.length > 0 && (
                          <>
                            <Divider sx={{ my: 1, borderColor: '#3f3f46' }} />
                            <Typography 
                              variant="caption" 
                              sx={{ px: 2, py: 0.5, color: '#71717a', display: 'block' }}
                            >
                              ⭐ Quick use (single prompt):
                            </Typography>
                            {favoritePrompts.slice(0, 5).map((prompt) => (
                              <MenuItem 
                                key={prompt.id} 
                                onClick={() => handleCustomPrompt(prompt.id)}
                              >
                                <ListItemIcon>
                                  <PromptsIcon sx={{ color: '#a855f7', fontSize: 20 }} />
                                </ListItemIcon>
                                <ListItemText 
                                  primary={prompt.name}
                                  primaryTypographyProps={{ fontSize: '0.875rem' }}
                                />
                              </MenuItem>
                            ))}
                            {favoritePrompts.length > 5 && (
                              <MenuItem 
                                onClick={() => { setAiMenuAnchor(null); setTabValue(4); }}
                                sx={{ color: '#71717a' }}
                              >
                                <ListItemText 
                                  primary={`+ ${favoritePrompts.length - 5} more...`}
                                  primaryTypographyProps={{ fontSize: '0.8rem' }}
                                />
                              </MenuItem>
                            )}
                          </>
                        )}
                      </Menu>
                    </Box>

                {activeEmail ? (
                  <>
                    {/* AI Result Display - ABOVE EMAIL for visibility */}
                    {aiResult && (
                      <Paper
                        id="ai-result-panel"
                        sx={{
                          p: 3,
                          mb: 2,
                          background: 'linear-gradient(135deg, rgba(249, 115, 22, 0.15) 0%, rgba(6, 182, 212, 0.08) 100%)',
                          border: '2px solid rgba(249, 115, 22, 0.5)',
                          borderRadius: 2,
                          boxShadow: '0 4px 20px rgba(249, 115, 22, 0.2)',
                        }}
                      >
                        <Stack direction="row" justifyContent="space-between" alignItems="center" mb={2}>
                          <Stack direction="row" alignItems="center" spacing={1}>
                            <AIIcon sx={{ color: 'primary.main', fontSize: 28 }} />
                            <Typography variant="h6" sx={{ fontWeight: 700, color: 'primary.main' }}>
                              {aiResult.type === 'summarize' && '📝 AI Summary'}
                              {aiResult.type === 'reply' && '✉️ Draft Reply'}
                              {aiResult.type === 'insights' && '💡 Insights'}
                              {aiResult.type === 'custom' && '✨ Custom Prompt'}
                              {aiResult.type === 'multi' && '🎯 AI Response'}
                            </Typography>
                            <Chip 
                              label={aiSettings?.model || 'AI'} 
                              size="small" 
                              sx={{ 
                                background: 'rgba(249, 115, 22, 0.3)',
                                color: 'primary.main',
                                fontWeight: 600,
                              }} 
                            />
                            {aiResult.usedPrompts && aiResult.usedPrompts.length > 0 && (
                              <Chip
                                label={`${aiResult.usedPrompts.length} prompt${aiResult.usedPrompts.length > 1 ? 's' : ''}`}
                                size="small"
                                sx={{
                                  background: 'rgba(168, 85, 247, 0.3)',
                                  color: '#c084fc',
                                  fontWeight: 600,
                                }}
                              />
                            )}
                            {aiResult.hadQuickNotes && (
                              <Chip
                                label="+ quick notes"
                                size="small"
                                sx={{
                                  background: 'rgba(34, 197, 94, 0.3)',
                                  color: '#4ade80',
                                  fontWeight: 600,
                                }}
                              />
                            )}
                          </Stack>
                          <Stack direction="row" spacing={1}>
                            {/* Always show "Open in Outlook" button for any AI result */}
                            <Tooltip title="Open as Reply in Outlook">
                              <Button
                                variant="contained"
                                size="small"
                                startIcon={<OpenInOutlookIcon />}
                                onClick={() => sendDraftToOutlook(false)}
                                sx={{
                                  background: 'linear-gradient(135deg, #22c55e 0%, #16a34a 100%)',
                                  '&:hover': {
                                    background: 'linear-gradient(135deg, #16a34a 0%, #15803d 100%)',
                                  },
                                  }}
                                >
                                  Open in Outlook
                                </Button>
                              </Tooltip>
                            <Tooltip title="Copy to clipboard">
                              <IconButton onClick={copyAIResult} size="small" sx={{ color: 'primary.main' }}>
                                <CopyIcon />
                              </IconButton>
                            </Tooltip>
                            <Tooltip title="Close">
                              <IconButton onClick={() => setAiResult(null)} size="small" sx={{ color: 'text.secondary' }}>
                                <ExpandLessIcon />
                              </IconButton>
                            </Tooltip>
                          </Stack>
                        </Stack>
                        <Divider sx={{ mb: 2, borderColor: 'rgba(249, 115, 22, 0.3)' }} />
                        <Box
                          sx={{
                            p: 2.5,
                            background: 'rgba(15, 15, 18, 0.8)',
                            borderRadius: 2,
                            maxHeight: 350,
                            overflow: 'auto',
                            fontFamily: '"Plus Jakarta Sans", sans-serif',
                            fontSize: '0.95rem',
                            lineHeight: 1.9,
                            whiteSpace: 'pre-wrap',
                            wordBreak: 'break-word',
                            color: '#e4e4e7',
                          }}
                        >
                          {aiResult.content}
                        </Box>
                        <Box sx={{ mt: 1.5 }}>
                          <Typography variant="caption" sx={{ display: 'block', color: 'text.secondary' }}>
                            Generated for: "{aiResult.email}"
                          </Typography>
                          {aiResult.usedPrompts && aiResult.usedPrompts.length > 0 && (
                            <Typography variant="caption" sx={{ display: 'block', color: '#a855f7', mt: 0.5 }}>
                              📎 Prompts used: {aiResult.usedPrompts.join(', ')}
                            </Typography>
                          )}
                          {(!aiResult.usedPrompts || aiResult.usedPrompts.length === 0) && aiResult.type === 'multi' && (
                            <Typography variant="caption" sx={{ display: 'block', color: '#71717a', mt: 0.5 }}>
                              📎 No prompts selected — raw email analysis
                            </Typography>
                          )}
                          {aiResult.hadQuickNotes && (
                            <Typography variant="caption" sx={{ display: 'block', color: '#22c55e', mt: 0.5 }}>
                              ✏️ Included quick notes
                            </Typography>
                          )}
                          {aiResult.attachmentCount > 0 && (
                            <Typography variant="caption" sx={{ display: 'block', color: '#8b5cf6', mt: 0.5 }}>
                              📎 Analyzed {aiResult.attachmentCount} attachment(s)
                            </Typography>
                          )}
                        </Box>
                      </Paper>
                    )}

                    {/* Attachment Summaries Panel */}
                    {attachmentSummaries.length > 0 && (
                      <Paper
                        sx={{
                          p: 2.5,
                          mb: 2,
                          background: 'linear-gradient(135deg, rgba(139, 92, 246, 0.15) 0%, rgba(236, 72, 153, 0.08) 100%)',
                          border: '2px solid rgba(139, 92, 246, 0.4)',
                          borderRadius: 2,
                        }}
                      >
                        <Stack direction="row" justifyContent="space-between" alignItems="center" mb={2}>
                          <Stack direction="row" alignItems="center" spacing={1}>
                            <Typography variant="h6" sx={{ fontWeight: 700, color: '#8b5cf6' }}>
                              📎 Attachment Analysis ({attachmentSummaries.length} file{attachmentSummaries.length > 1 ? 's' : ''})
                            </Typography>
                          </Stack>
                          <Tooltip title="Clear attachments">
                            <IconButton 
                              size="small" 
                              onClick={() => setAttachmentSummaries([])}
                              sx={{ color: '#71717a', '&:hover': { color: '#8b5cf6' } }}
                            >
                              <ClearAllIcon fontSize="small" />
                            </IconButton>
                          </Tooltip>
                        </Stack>
                        <Stack spacing={2}>
                          {attachmentSummaries.map((att, index) => (
                            <Box
                              key={index}
                              sx={{
                                p: 2,
                                background: 'rgba(15, 15, 18, 0.8)',
                                borderRadius: 1.5,
                                border: '1px solid rgba(139, 92, 246, 0.2)',
                              }}
                            >
                              <Stack direction="row" alignItems="center" spacing={1} mb={1}>
                                <Typography 
                                  variant="subtitle2" 
                                  sx={{ 
                                    fontWeight: 700, 
                                    color: '#c084fc',
                                    fontFamily: 'monospace',
                                  }}
                                >
                                  📄 {att.filename}
                                </Typography>
                                <Chip
                                  label={att.type?.toUpperCase() || 'FILE'}
                                  size="small"
                                  sx={{
                                    background: 'rgba(139, 92, 246, 0.3)',
                                    color: '#c084fc',
                                    fontSize: '0.65rem',
                                    height: 20,
                                  }}
                                />
                                {att.charCount && (
                                  <Typography variant="caption" sx={{ color: '#71717a' }}>
                                    ({Math.round(att.charCount / 1000)}k chars)
                                  </Typography>
                                )}
                              </Stack>
                              <Typography
                                variant="body2"
                                sx={{
                                  color: '#a1a1aa',
                                  whiteSpace: 'pre-wrap',
                                  fontSize: '0.85rem',
                                  lineHeight: 1.7,
                                  maxHeight: 200,
                                  overflow: 'auto',
                                }}
                              >
                                {att.summary}
                              </Typography>
                            </Box>
                          ))}
                        </Stack>
                      </Paper>
                    )}

                    <EmailViewer 
                      email={activeEmail} 
                      onContactClick={() => setShowContactCard(true)}
                    />
                  </>
                ) : (
                  <Paper
                    sx={{
                      p: 4,
                      background: '#1a1a1f',
                      border: '1px dashed #3f3f46',
                      textAlign: 'center',
                      borderRadius: 2,
                    }}
                  >
                    <EmailIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
                    <Typography variant="h6" sx={{ color: 'text.secondary', mb: 1 }}>
                      No Active Email
                    </Typography>
                    <Typography variant="body2" sx={{ color: '#71717a' }}>
                      {status.outlook === 'online' 
                        ? 'Select an email in Outlook and click "Fetch Active Email"'
                        : 'Please open Outlook and select an email'}
                    </Typography>
                  </Paper>
                )}
                  </Stack>
                </Box>

                {/* Prompt Selector - Right Side */}
                <Box sx={{ width: 320, flexShrink: 0 }}>
                  <PromptSelector
                    prompts={allPrompts}
                    selectedPromptIds={selectedPromptIds}
                    setSelectedPromptIds={setSelectedPromptIds}
                    hasPersonalContext={Boolean(aiSettings?.globalContext?.enabled !== false && (aiSettings?.globalContext?.name || aiSettings?.globalContext?.customNotes))}
                    personalContextEnabled={aiSettings?.globalContext?.enabled !== false}
                    companyContextEnabled={aiSettings?.companyContext?.enabled && !!aiSettings?.companyContext?.content}
                    presets={presets}
                    onSavePreset={handleSavePreset}
                    onLoadPreset={handleLoadPreset}
                    onDeletePreset={handleDeletePreset}
                    showMessage={showMessage}
                    quickNotes={quickNotes}
                    setQuickNotes={setQuickNotes}
                    keepQuickNotes={keepQuickNotes}
                    setKeepQuickNotes={setKeepQuickNotes}
                  />
                </Box>
              </Box>
            </TabPanel>

            {/* Recent Inbox Tab - New Enhanced Version */}
            <TabPanel value={tabValue} index={1}>
              <InboxTab
                showMessage={showMessage}
                aiSettings={aiSettings}
                selectedPromptIds={selectedPromptIds}
                allPrompts={allPrompts}
              />
            </TabPanel>

            {/* Sent Items Tab */}
            <TabPanel value={tabValue} index={2}>
              <Stack spacing={3} sx={{ height: '100%', overflow: 'auto' }}>
                <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
                  <Button
                    variant="contained"
                    startIcon={loading.sent ? <CircularProgress size={18} color="inherit" /> : <RefreshIcon />}
                    onClick={handleFetchSent}
                    disabled={loading.sent || status.outlook !== 'online'}
                  >
                    {loading.sent ? 'Fetching...' : 'Fetch Sent Items'}
                  </Button>
                  
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1, ml: 2 }}>
                    <Typography variant="body2" sx={{ color: 'text.secondary', minWidth: 80 }}>
                      Days: {daysBack}
                    </Typography>
                    <Slider
                      value={daysBack}
                      onChange={(e, val) => setDaysBack(val)}
                      min={7}
                      max={365}
                      sx={{ width: 120 }}
                      size="small"
                    />
                  </Box>
                  
                  <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                    <Typography variant="body2" sx={{ color: 'text.secondary', minWidth: 60 }}>
                      Max: {maxItems}
                    </Typography>
                    <Slider
                      value={maxItems}
                      onChange={(e, val) => setMaxItems(val)}
                      min={10}
                      max={200}
                      step={10}
                      sx={{ width: 100 }}
                      size="small"
                    />
                  </Box>
                </Box>

                <EmailTable 
                  emails={sentEmails} 
                  onSelectEmail={handleSelectEmail}
                  selectedId={selectedEmail?.entryId}
                  loading={loading.sent}
                />
                
                {selectedEmail && tabValue === 2 && (
                  <EmailViewer email={selectedEmail} />
                )}
              </Stack>
            </TabPanel>

            {/* Files Tab */}
            <TabPanel value={tabValue} index={3}>
              <FilesTab
                activeEmail={activeEmail}
                selectedEmail={selectedEmail}
                showMessage={showMessage}
                isElectron={isElectron}
              />
            </TabPanel>

            {/* Prompts Tab */}
            <TabPanel value={tabValue} index={4}>
              <PromptsTab
                showMessage={showMessage}
                isElectron={isElectron}
                onPromptsChange={async () => {
                  // Refresh ALL prompts when they change
                  if (isElectron) {
                    try {
                      const result = await window.electronAPI.getPrompts();
                      if (result.success) {
                        setAllPrompts(result.prompts);
                        setFavoritePrompts(result.prompts.filter(p => p.isFavorite));
                      }
                    } catch (e) { /* ignore */ }
                  }
                }}
              />
            </TabPanel>
          </Box>
        </Paper>
      </Container>

      {/* Settings Dialog */}
      <Settings
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        settings={aiSettings}
        onSaveSettings={handleSaveSettings}
      />
      
      {/* Keyboard Shortcuts Dialog */}
      <Dialog 
        open={shortcutsOpen} 
        onClose={() => setShortcutsOpen(false)}
        PaperProps={{
          sx: {
            background: '#18181b',
            border: '1px solid #27272a',
            minWidth: 400,
          }
        }}
      >
        <DialogTitle sx={{ 
          display: 'flex', 
          justifyContent: 'space-between', 
          alignItems: 'center',
          borderBottom: '1px solid #27272a',
        }}>
          <Stack direction="row" alignItems="center" spacing={1}>
            <KeyboardIcon sx={{ color: 'secondary.main' }} />
            <Typography variant="h6">Keyboard Shortcuts</Typography>
          </Stack>
          <IconButton onClick={() => setShortcutsOpen(false)} size="small">
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        <DialogContent sx={{ pt: 3 }}>
          <Stack spacing={2}>
            <Typography variant="subtitle2" sx={{ color: 'primary.main', mb: 1 }}>
              Global Hotkeys (work even when app is minimized)
            </Typography>
            
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Typography variant="body2">Show/Hide Window</Typography>
              <Chip label="Ctrl + Shift + O" size="small" sx={{ fontFamily: 'monospace', background: '#27272a' }} />
            </Box>
            
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Typography variant="body2">Fetch Active Email</Typography>
              <Chip label="Ctrl + Shift + G" size="small" sx={{ fontFamily: 'monospace', background: '#27272a' }} />
            </Box>
            
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Typography variant="body2">Quick Summarize</Typography>
              <Chip label="Ctrl + Shift + S" size="small" sx={{ fontFamily: 'monospace', background: '#27272a' }} />
            </Box>
            
            <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
              <Typography variant="body2">Quick Draft Reply</Typography>
              <Chip label="Ctrl + Shift + D" size="small" sx={{ fontFamily: 'monospace', background: '#27272a' }} />
            </Box>
            
            <Divider sx={{ my: 2, borderColor: '#27272a' }} />
            
            <Typography variant="caption" sx={{ color: 'text.secondary' }}>
              💡 Tip: Use Quick Summarize or Quick Draft from Outlook - they'll fetch the selected email and process it automatically!
            </Typography>
          </Stack>
        </DialogContent>
      </Dialog>

      {/* Contact Card Dialog */}
      <Dialog
        open={showContactCard && !!activeEmail}
        onClose={() => setShowContactCard(false)}
        maxWidth="md"
        fullWidth
        PaperProps={{
          sx: {
            background: 'transparent',
            boxShadow: 'none',
            overflow: 'visible',
          }
        }}
      >
        <ContactCard
          email={activeEmail?.senderEmail}
          senderName={activeEmail?.senderName}
          onClose={() => setShowContactCard(false)}
          showMessage={showMessage}
        />
      </Dialog>
    </Box>
  );
}

export default App;
