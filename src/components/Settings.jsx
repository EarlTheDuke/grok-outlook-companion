import React, { useState, useEffect } from 'react';
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  TextField,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  Box,
  Typography,
  IconButton,
  Alert,
  Divider,
  InputAdornment,
  Chip,
  Stack,
  RadioGroup,
  FormControlLabel,
  Radio,
  FormLabel,
  Tabs,
  Tab,
  Switch,
} from '@mui/material';
import {
  Close as CloseIcon,
  Visibility,
  VisibilityOff,
  Check as CheckIcon,
  AutoAwesome as AIIcon,
  Person as PersonIcon,
  Settings as SettingsIcon,
  Business as BusinessIcon,
  BugReport as BugReportIcon,
  Delete as DeleteIcon,
  Download as DownloadIcon,
  Refresh as RefreshIcon,
} from '@mui/icons-material';

const AI_PROVIDERS = [
  { id: 'grok', name: 'Grok (xAI)', model: 'grok-4', endpoint: 'https://api.x.ai/v1/chat/completions' },
  { id: 'ollama', name: 'Ollama (Local)', model: 'llama3.2', endpoint: 'http://localhost:11434/api/chat' },
  { id: 'openai', name: 'OpenAI', model: 'gpt-4', endpoint: 'https://api.openai.com/v1/chat/completions' },
  { id: 'custom', name: 'Custom API', model: '', endpoint: '' },
];

function Settings({ open, onClose, settings, onSaveSettings }) {
  const [tabValue, setTabValue] = useState(0);
  
  // AI Provider settings
  const [provider, setProvider] = useState(settings?.provider || 'grok');
  const [apiKey, setApiKey] = useState('');
  const [model, setModel] = useState(settings?.model || 'grok-4');
  const [endpoint, setEndpoint] = useState(settings?.endpoint || '');
  const [showApiKey, setShowApiKey] = useState(false);
  const [saveStatus, setSaveStatus] = useState(null);
  const [hasExistingKey, setHasExistingKey] = useState(false);
  
  // Global Context settings
  const [globalContext, setGlobalContext] = useState({
    enabled: true,
    name: '',
    role: '',
    company: '',
    industry: '',
    communicationStyle: '', // empty = none selected
    detailLevel: '', // empty = none selected
    customNotes: '',
  });
  
  // Company Context settings
  const [companyContext, setCompanyContext] = useState({
    enabled: false,
    content: '',
  });
  
  // Telemetry settings
  const [telemetryEnabled, setTelemetryEnabled] = useState(false);
  const [crashReports, setCrashReports] = useState([]);
  const [loadingReports, setLoadingReports] = useState(false);

  // Load existing settings when dialog opens
  useEffect(() => {
    if (open && settings) {
      setProvider(settings.provider || 'grok');
      setModel(settings.model || 'grok-4');
      setEndpoint(settings.endpoint || '');
      setHasExistingKey(settings.hasApiKey || false);
      
      // Load global context
      if (settings.globalContext) {
        setGlobalContext({
          enabled: settings.globalContext.enabled !== false, // default true
          name: settings.globalContext.name || '',
          role: settings.globalContext.role || '',
          company: settings.globalContext.company || '',
          industry: settings.globalContext.industry || '',
          communicationStyle: settings.globalContext.communicationStyle || '', // allow empty
          detailLevel: settings.globalContext.detailLevel || '', // allow empty
          customNotes: settings.globalContext.customNotes || '',
        });
      }
      
      // Load company context
      if (settings.companyContext) {
        setCompanyContext({
          enabled: settings.companyContext.enabled || false,
          content: settings.companyContext.content || '',
        });
      }
      
      // Load telemetry setting
      setTelemetryEnabled(settings.telemetryEnabled || false);
    }
  }, [open, settings]);
  
  // Load crash reports when telemetry tab is selected
  const loadCrashReports = async () => {
    if (window.electronAPI?.getCrashReports) {
      setLoadingReports(true);
      try {
        const result = await window.electronAPI.getCrashReports();
        if (result.success) {
          setCrashReports(result.reports || []);
        }
      } catch (e) {
        console.error('Error loading crash reports:', e);
      } finally {
        setLoadingReports(false);
      }
    }
  };
  
  useEffect(() => {
    if (tabValue === 3 && open) {
      loadCrashReports();
    }
  }, [tabValue, open]);
  
  // Estimate tokens (rough: ~4 chars per token)
  const estimateTokens = (text) => {
    if (!text) return 0;
    return Math.ceil(text.length / 4);
  };
  
  const handleGlobalContextChange = (field, value) => {
    setGlobalContext(prev => ({ ...prev, [field]: value }));
  };

  // Update model and endpoint when provider changes
  useEffect(() => {
    const providerConfig = AI_PROVIDERS.find(p => p.id === provider);
    if (providerConfig && provider !== 'custom') {
      setModel(providerConfig.model);
      setEndpoint(providerConfig.endpoint);
    }
  }, [provider]);

  const handleSave = async () => {
    try {
      setSaveStatus('saving');
      
      await onSaveSettings({
        provider,
        model,
        endpoint,
        apiKey: apiKey || undefined, // Only send if changed
        globalContext,
        companyContext,
        telemetryEnabled,
      });
      
      setSaveStatus('success');
      setTimeout(() => {
        setSaveStatus(null);
        onClose();
      }, 1000);
    } catch (error) {
      setSaveStatus('error');
      console.error('Error saving settings:', error);
    }
  };
  
  // Handle clearing crash reports
  const handleClearReports = async () => {
    if (window.electronAPI?.clearCrashReports) {
      const result = await window.electronAPI.clearCrashReports();
      if (result.success) {
        setCrashReports([]);
      }
    }
  };
  
  // Handle exporting crash reports
  const handleExportReports = async () => {
    if (window.electronAPI?.exportCrashReports) {
      await window.electronAPI.exportCrashReports();
    }
  };

  const selectedProvider = AI_PROVIDERS.find(p => p.id === provider);

  return (
    <Dialog 
      open={open} 
      onClose={onClose}
      maxWidth="sm"
      fullWidth
      PaperProps={{
        sx: {
          background: 'linear-gradient(135deg, #18181b 0%, #0f0f12 100%)',
          border: '1px solid #27272a',
        }
      }}
    >
      <DialogTitle sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between' }}>
        <Stack direction="row" alignItems="center" spacing={1}>
          <AIIcon sx={{ color: 'primary.main' }} />
          <Typography variant="h6">AI Settings</Typography>
        </Stack>
        <IconButton onClick={onClose} sx={{ color: 'text.secondary' }}>
          <CloseIcon />
        </IconButton>
      </DialogTitle>

      <DialogContent sx={{ p: 0 }}>
        {/* Tabs */}
        <Tabs 
          value={tabValue} 
          onChange={(e, v) => setTabValue(v)}
          sx={{ 
            borderBottom: '1px solid #27272a',
            px: 2,
            '& .MuiTab-root': { minHeight: 48 },
          }}
        >
          <Tab icon={<PersonIcon sx={{ fontSize: 18 }} />} iconPosition="start" label="About You" />
          <Tab icon={<BusinessIcon sx={{ fontSize: 18 }} />} iconPosition="start" label="Company" />
          <Tab icon={<SettingsIcon sx={{ fontSize: 18 }} />} iconPosition="start" label="AI Provider" />
          <Tab icon={<BugReportIcon sx={{ fontSize: 18 }} />} iconPosition="start" label="Telemetry" />
        </Tabs>

        {/* About You Tab */}
        {tabValue === 0 && (
          <Box sx={{ p: 3 }}>
            {/* Enable/Disable Toggle */}
            <Box
              sx={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                p: 2,
                mb: 2,
                background: globalContext.enabled ? 'rgba(34, 197, 94, 0.1)' : 'rgba(239, 68, 68, 0.1)',
                border: `1px solid ${globalContext.enabled ? 'rgba(34, 197, 94, 0.3)' : 'rgba(239, 68, 68, 0.3)'}`,
                borderRadius: 1,
              }}
            >
              <Box>
                <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
                  Include Global Context in AI Requests
                </Typography>
                <Typography variant="caption" sx={{ color: 'text.secondary' }}>
                  {globalContext.enabled 
                    ? 'Your info below is sent with every AI request' 
                    : 'Disabled — only prompts are sent (faster, fewer tokens)'}
                </Typography>
              </Box>
              <Switch
                checked={globalContext.enabled}
                onChange={(e) => handleGlobalContextChange('enabled', e.target.checked)}
                color="success"
              />
            </Box>

            <Alert severity="info" sx={{ mb: 3, opacity: globalContext.enabled ? 1 : 0.5 }}>
              This information is included in <strong>ALL</strong> AI interactions to give context about you and your preferences.
            </Alert>

            <Stack spacing={2.5} sx={{ opacity: globalContext.enabled ? 1 : 0.5 }}>
              <Stack direction="row" spacing={2}>
                <TextField
                  fullWidth
                  label="Your Name"
                  value={globalContext.name}
                  onChange={(e) => handleGlobalContextChange('name', e.target.value)}
                  placeholder="John Smith"
                  size="small"
                />
                <TextField
                  fullWidth
                  label="Your Role"
                  value={globalContext.role}
                  onChange={(e) => handleGlobalContextChange('role', e.target.value)}
                  placeholder="Sales Director"
                  size="small"
                />
              </Stack>

              <Stack direction="row" spacing={2}>
                <TextField
                  fullWidth
                  label="Company"
                  value={globalContext.company}
                  onChange={(e) => handleGlobalContextChange('company', e.target.value)}
                  placeholder="Acme Corporation"
                  size="small"
                />
                <TextField
                  fullWidth
                  label="Industry"
                  value={globalContext.industry}
                  onChange={(e) => handleGlobalContextChange('industry', e.target.value)}
                  placeholder="Technology / SaaS"
                  size="small"
                />
              </Stack>

              <Divider sx={{ my: 1 }} />

              <Stack direction="row" spacing={4}>
                <FormControl component="fieldset">
                  <FormLabel sx={{ fontSize: '0.85rem', mb: 1 }}>Communication Style</FormLabel>
                  <RadioGroup
                    value={globalContext.communicationStyle}
                    onChange={(e) => handleGlobalContextChange('communicationStyle', e.target.value)}
                    row
                  >
                    <FormControlLabel value="" control={<Radio size="small" />} label="None" sx={{ '& .MuiFormControlLabel-label': { color: '#71717a' } }} />
                    <FormControlLabel value="professional" control={<Radio size="small" />} label="Professional" />
                    <FormControlLabel value="friendly" control={<Radio size="small" />} label="Friendly" />
                    <FormControlLabel value="casual" control={<Radio size="small" />} label="Casual" />
                  </RadioGroup>
                </FormControl>

                <FormControl component="fieldset">
                  <FormLabel sx={{ fontSize: '0.85rem', mb: 1 }}>Detail Level</FormLabel>
                  <RadioGroup
                    value={globalContext.detailLevel}
                    onChange={(e) => handleGlobalContextChange('detailLevel', e.target.value)}
                    row
                  >
                    <FormControlLabel value="" control={<Radio size="small" />} label="None" sx={{ '& .MuiFormControlLabel-label': { color: '#71717a' } }} />
                    <FormControlLabel value="concise" control={<Radio size="small" />} label="Concise" />
                    <FormControlLabel value="detailed" control={<Radio size="small" />} label="Detailed" />
                  </RadioGroup>
                </FormControl>
              </Stack>

              <TextField
                fullWidth
                label="Additional Context / Notes"
                value={globalContext.customNotes}
                onChange={(e) => handleGlobalContextChange('customNotes', e.target.value)}
                placeholder="- I prefer bullet points over paragraphs&#10;- Our fiscal year ends in March&#10;- Key clients: Microsoft, Google&#10;- Always include action items"
                multiline
                rows={4}
                helperText="Add any extra context the AI should know about you, your preferences, or your work"
                inputProps={{ spellCheck: true }}
              />
            </Stack>
          </Box>
        )}

        {/* Company Context Tab */}
        {tabValue === 1 && (
          <Box sx={{ p: 3 }}>
            {/* Enable/Disable Toggle */}
            <Box
              sx={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                p: 2,
                mb: 3,
                background: companyContext.enabled ? 'rgba(59, 130, 246, 0.1)' : 'rgba(239, 68, 68, 0.1)',
                border: `1px solid ${companyContext.enabled ? 'rgba(59, 130, 246, 0.3)' : 'rgba(239, 68, 68, 0.3)'}`,
                borderRadius: 1,
              }}
            >
              <Box>
                <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
                  Include Company Context in AI Requests
                </Typography>
                <Typography variant="caption" sx={{ color: 'text.secondary' }}>
                  {companyContext.enabled 
                    ? 'Full company context sent with every AI request' 
                    : 'Disabled — faster responses, fewer tokens'}
                </Typography>
              </Box>
              <Switch
                checked={companyContext.enabled}
                onChange={(e) => setCompanyContext(prev => ({ ...prev, enabled: e.target.checked }))}
                color="primary"
              />
            </Box>

            {/* Token Estimate */}
            {companyContext.content && (
              <Alert 
                severity={estimateTokens(companyContext.content) > 2000 ? 'warning' : 'info'} 
                sx={{ mb: 2 }}
              >
                <strong>Estimated tokens:</strong> ~{estimateTokens(companyContext.content).toLocaleString()}
                {estimateTokens(companyContext.content) > 2000 && (
                  <> — Large context may slow responses</>
                )}
              </Alert>
            )}

            <Typography variant="body2" sx={{ color: 'text.secondary', mb: 2 }}>
              Paste your detailed company information, capabilities, clients, and any context 
              you want the AI to consider when generating responses. This is ideal for company-specific 
              knowledge that applies to most communications.
            </Typography>

            <TextField
              fullWidth
              label="Company Context"
              value={companyContext.content}
              onChange={(e) => setCompanyContext(prev => ({ ...prev, content: e.target.value }))}
              placeholder="Paste your company overview, capabilities, client list, values, processes, etc..."
              multiline
              rows={14}
              inputProps={{ maxLength: 15000, spellCheck: true }}
              helperText={`${companyContext.content.length.toLocaleString()} / 15,000 characters`}
              sx={{
                '& .MuiInputBase-input': {
                  fontFamily: 'monospace',
                  fontSize: '0.85rem',
                  lineHeight: 1.5,
                },
                opacity: companyContext.enabled ? 1 : 0.5,
              }}
            />
          </Box>
        )}

        {/* AI Provider Tab */}
        {tabValue === 2 && (
          <Box sx={{ p: 3 }}>
            {/* Provider Selection */}
            <FormControl fullWidth sx={{ mb: 3 }}>
              <InputLabel>AI Provider</InputLabel>
              <Select
                value={provider}
                label="AI Provider"
                onChange={(e) => setProvider(e.target.value)}
              >
                {AI_PROVIDERS.map((p) => (
                  <MenuItem key={p.id} value={p.id}>
                    <Stack direction="row" alignItems="center" spacing={1}>
                      <span>{p.name}</span>
                      {p.id === 'ollama' && (
                        <Chip label="Local" size="small" color="success" sx={{ height: 20 }} />
                      )}
                    </Stack>
                  </MenuItem>
                ))}
              </Select>
            </FormControl>

            {/* API Key (not needed for Ollama) */}
            {provider !== 'ollama' && (
              <TextField
                fullWidth
                label="API Key"
                type={showApiKey ? 'text' : 'password'}
                value={apiKey}
                onChange={(e) => setApiKey(e.target.value)}
                placeholder={hasExistingKey ? '••••••••••••••••' : 'Enter your API key'}
                helperText={
                  provider === 'grok' 
                    ? 'Get your API key from console.x.ai' 
                    : provider === 'openai'
                    ? 'Get your API key from platform.openai.com'
                    : 'Enter your API key'
                }
                sx={{ mb: 3 }}
                InputProps={{
                  endAdornment: (
                    <InputAdornment position="end">
                      <IconButton onClick={() => setShowApiKey(!showApiKey)} edge="end">
                        {showApiKey ? <VisibilityOff /> : <Visibility />}
                      </IconButton>
                    </InputAdornment>
                  ),
                }}
              />
            )}

            {provider === 'ollama' && (
              <Alert severity="info" sx={{ mb: 3 }}>
                Ollama runs locally - no API key needed! Make sure Ollama is running on your computer.
              </Alert>
            )}

            <Divider sx={{ my: 2 }} />

            {/* Model Selection */}
            <TextField
              fullWidth
              label="Model"
              value={model}
              onChange={(e) => setModel(e.target.value)}
              helperText={
                provider === 'grok' 
                  ? 'e.g., grok-4, grok-3' 
                  : provider === 'ollama'
                  ? 'e.g., llama3.2, mistral, codellama'
                  : 'Enter the model name'
              }
              sx={{ mb: 3 }}
            />

            {/* Custom Endpoint (for custom provider) */}
            {provider === 'custom' && (
              <TextField
                fullWidth
                label="API Endpoint"
                value={endpoint}
                onChange={(e) => setEndpoint(e.target.value)}
                placeholder="https://api.example.com/v1/chat/completions"
                helperText="Full URL to the chat completions endpoint"
                sx={{ mb: 3 }}
              />
            )}
          </Box>
        )}

        {/* Telemetry Tab */}
        {tabValue === 3 && (
          <Box sx={{ p: 3 }}>
            <Typography variant="h6" sx={{ mb: 2, color: 'text.primary' }}>
              Crash Reports & Telemetry
            </Typography>
            
            <Alert severity="info" sx={{ mb: 3 }}>
              Help improve Grok-Outlook Companion by sharing anonymous crash reports.
              No email content or personal data is ever collected.
            </Alert>
            
            {/* Telemetry Toggle */}
            <Box
              sx={{
                display: 'flex',
                alignItems: 'center',
                justifyContent: 'space-between',
                p: 2,
                mb: 3,
                background: telemetryEnabled ? 'rgba(34, 197, 94, 0.1)' : 'rgba(113, 113, 122, 0.1)',
                borderRadius: 2,
                border: `1px solid ${telemetryEnabled ? 'rgba(34, 197, 94, 0.3)' : '#3f3f46'}`,
              }}
            >
              <Box>
                <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
                  Enable Crash Reporting
                </Typography>
                <Typography variant="body2" sx={{ color: 'text.secondary' }}>
                  Automatically log errors locally for troubleshooting
                </Typography>
              </Box>
              <Switch
                checked={telemetryEnabled}
                onChange={(e) => setTelemetryEnabled(e.target.checked)}
                color="primary"
              />
            </Box>
            
            <Divider sx={{ my: 2 }} />
            
            {/* Crash Reports Section */}
            <Box sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', mb: 2 }}>
              <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
                Saved Crash Reports ({crashReports.length})
              </Typography>
              <Stack direction="row" spacing={1}>
                <Button
                  size="small"
                  startIcon={<RefreshIcon />}
                  onClick={loadCrashReports}
                  disabled={loadingReports}
                >
                  Refresh
                </Button>
                <Button
                  size="small"
                  startIcon={<DownloadIcon />}
                  onClick={handleExportReports}
                  disabled={crashReports.length === 0}
                >
                  Export
                </Button>
                <Button
                  size="small"
                  color="error"
                  startIcon={<DeleteIcon />}
                  onClick={handleClearReports}
                  disabled={crashReports.length === 0}
                >
                  Clear All
                </Button>
              </Stack>
            </Box>
            
            {/* Crash Reports List */}
            <Box
              sx={{
                maxHeight: 250,
                overflow: 'auto',
                border: '1px solid #27272a',
                borderRadius: 1,
                background: '#0f0f12',
              }}
            >
              {crashReports.length === 0 ? (
                <Box sx={{ p: 3, textAlign: 'center' }}>
                  <Typography variant="body2" sx={{ color: '#71717a' }}>
                    {telemetryEnabled 
                      ? '✅ No crash reports — everything is running smoothly!' 
                      : 'Enable crash reporting to start collecting error logs'}
                  </Typography>
                </Box>
              ) : (
                crashReports.map((report, index) => (
                  <Box
                    key={report.id || index}
                    sx={{
                      p: 1.5,
                      borderBottom: index < crashReports.length - 1 ? '1px solid #27272a' : 'none',
                      '&:hover': { background: 'rgba(255,255,255,0.02)' },
                    }}
                  >
                    <Stack direction="row" justifyContent="space-between" alignItems="center">
                      <Typography variant="caption" sx={{ color: '#f97316', fontWeight: 600 }}>
                        {report.type}
                      </Typography>
                      <Typography variant="caption" sx={{ color: '#71717a' }}>
                        {new Date(report.timestamp).toLocaleString()}
                      </Typography>
                    </Stack>
                    <Typography
                      variant="body2"
                      sx={{
                        color: '#a1a1aa',
                        fontFamily: 'monospace',
                        fontSize: '0.75rem',
                        mt: 0.5,
                        whiteSpace: 'nowrap',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                      }}
                    >
                      {report.error}
                    </Typography>
                  </Box>
                ))
              )}
            </Box>
            
            <Typography variant="caption" sx={{ display: 'block', color: '#52525b', mt: 2 }}>
              💡 Tip: Export crash reports and share them when reporting issues for faster troubleshooting.
            </Typography>
          </Box>
        )}

        {/* Status Messages */}
        {saveStatus === 'success' && (
          <Alert 
            severity="success" 
            icon={<CheckIcon />}
            sx={{ mx: 3, mb: 2 }}
          >
            Settings saved successfully!
          </Alert>
        )}
        {saveStatus === 'error' && (
          <Alert severity="error" sx={{ mx: 3, mb: 2 }}>
            Failed to save settings. Please try again.
          </Alert>
        )}
      </DialogContent>

      <DialogActions sx={{ px: 3, pb: 3 }}>
        <Button onClick={onClose} sx={{ color: 'text.secondary' }}>
          Cancel
        </Button>
        <Button 
          variant="contained" 
          onClick={handleSave}
          disabled={saveStatus === 'saving'}
        >
          {saveStatus === 'saving' ? 'Saving...' : 'Save Settings'}
        </Button>
      </DialogActions>
    </Dialog>
  );
}

export default Settings;

