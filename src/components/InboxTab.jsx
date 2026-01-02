import React, { useState, useMemo, useCallback } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  IconButton,
  Checkbox,
  Chip,
  TextField,
  InputAdornment,
  FormControlLabel,
  Switch,
  Tooltip,
  CircularProgress,
  Divider,
  Stack,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Badge,
  LinearProgress,
} from '@mui/material';
import {
  Search as SearchIcon,
  Refresh as RefreshIcon,
  FilterList as FilterIcon,
  Email as EmailIcon,
  AttachFile as AttachmentIcon,
  PriorityHigh as HighPriorityIcon,
  CheckCircle as ReadIcon,
  RadioButtonUnchecked as UnreadIcon,
  ExpandMore as ExpandIcon,
  ExpandLess as CollapseIcon,
  Clear as ClearIcon,
  SelectAll as SelectAllIcon,
  Psychology as AIIcon,
  Summarize as SummarizeIcon,
  ChecklistRtl as ActionItemsIcon,
  Person as PersonIcon,
  Business as DomainIcon,
} from '@mui/icons-material';

const InboxTab = ({ 
  showMessage, 
  aiSettings,
  selectedPromptIds,
  allPrompts,
}) => {
  // Email data
  const [emails, setEmails] = useState([]);
  const [loading, setLoading] = useState(false);
  const [loadProgress, setLoadProgress] = useState(0);
  
  // Selection
  const [selectedIds, setSelectedIds] = useState(new Set());
  const [previewEmail, setPreviewEmail] = useState(null);
  
  // Filters
  const [searchQuery, setSearchQuery] = useState('');
  const [showUnreadOnly, setShowUnreadOnly] = useState(false);
  const [showWithAttachments, setShowWithAttachments] = useState(false);
  const [showHighPriority, setShowHighPriority] = useState(false);
  const [selectedSenders, setSelectedSenders] = useState([]);
  const [selectedDomains, setSelectedDomains] = useState([]);
  const [dateRange, setDateRange] = useState('24h');
  
  // UI State
  const [showFilters, setShowFilters] = useState(true);

  // Fetch emails from Outlook
  const fetchEmails = useCallback(async () => {
    if (!window.electronAPI) {
      showMessage('warning', 'Electron API not available');
      return;
    }

    setLoading(true);
    setLoadProgress(10);
    setEmails([]);
    setSelectedIds(new Set());
    setPreviewEmail(null);

    try {
      const daysMap = {
        '24h': 1,
        '48h': 2,
        '7d': 7,
        '30d': 30,
      };
      
      const result = await window.electronAPI.getRecentInbox({
        daysBack: daysMap[dateRange] || 1,
        maxItems: 300, // Fetch up to 300
      });

      setLoadProgress(80);

      if (result.success) {
        setEmails(result.data || []);
        showMessage('success', `Loaded ${result.count} emails`);
      } else {
        showMessage('error', result.message || 'Failed to fetch emails');
      }
    } catch (error) {
      console.error('Error fetching inbox:', error);
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(false);
      setLoadProgress(100);
    }
  }, [dateRange, showMessage]);

  // Compute unique senders and domains for filter chips
  const { uniqueSenders, uniqueDomains } = useMemo(() => {
    const senderCounts = {};
    const domainCounts = {};
    
    emails.forEach(email => {
      const key = email.senderEmail || email.senderName;
      if (key) {
        senderCounts[key] = (senderCounts[key] || 0) + 1;
      }
      if (email.senderDomain) {
        domainCounts[email.senderDomain] = (domainCounts[email.senderDomain] || 0) + 1;
      }
    });
    
    const senders = Object.entries(senderCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 10)
      .map(([email, count]) => ({ email, count }));
      
    const domains = Object.entries(domainCounts)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([domain, count]) => ({ domain, count }));
    
    return { uniqueSenders: senders, uniqueDomains: domains };
  }, [emails]);

  // Filter emails based on current filters
  const filteredEmails = useMemo(() => {
    return emails.filter(email => {
      // Search query
      if (searchQuery) {
        const query = searchQuery.toLowerCase();
        const matchesSearch = 
          email.subject?.toLowerCase().includes(query) ||
          email.senderName?.toLowerCase().includes(query) ||
          email.senderEmail?.toLowerCase().includes(query) ||
          email.bodyPreview?.toLowerCase().includes(query);
        if (!matchesSearch) return false;
      }
      
      // Unread only
      if (showUnreadOnly && email.isRead) return false;
      
      // Has attachments
      if (showWithAttachments && !email.hasAttachments) return false;
      
      // High priority
      if (showHighPriority && email.importance !== 'high') return false;
      
      // Selected senders
      if (selectedSenders.length > 0) {
        const senderKey = email.senderEmail || email.senderName;
        if (!selectedSenders.includes(senderKey)) return false;
      }
      
      // Selected domains
      if (selectedDomains.length > 0) {
        if (!selectedDomains.includes(email.senderDomain)) return false;
      }
      
      return true;
    });
  }, [emails, searchQuery, showUnreadOnly, showWithAttachments, showHighPriority, selectedSenders, selectedDomains]);

  // Selection handlers
  const toggleSelection = (entryId) => {
    setSelectedIds(prev => {
      const newSet = new Set(prev);
      if (newSet.has(entryId)) {
        newSet.delete(entryId);
      } else {
        newSet.add(entryId);
      }
      return newSet;
    });
  };

  const selectAll = () => {
    setSelectedIds(new Set(filteredEmails.map(e => e.entryId)));
  };

  const deselectAll = () => {
    setSelectedIds(new Set());
  };

  const toggleSenderFilter = (sender) => {
    setSelectedSenders(prev => 
      prev.includes(sender) 
        ? prev.filter(s => s !== sender)
        : [...prev, sender]
    );
  };

  const toggleDomainFilter = (domain) => {
    setSelectedDomains(prev => 
      prev.includes(domain) 
        ? prev.filter(d => d !== domain)
        : [...prev, domain]
    );
  };

  const clearAllFilters = () => {
    setSearchQuery('');
    setShowUnreadOnly(false);
    setShowWithAttachments(false);
    setShowHighPriority(false);
    setSelectedSenders([]);
    setSelectedDomains([]);
  };

  // Format time for display
  const formatTime = (isoString) => {
    const date = new Date(isoString);
    const now = new Date();
    const isToday = date.toDateString() === now.toDateString();
    
    if (isToday) {
      return date.toLocaleTimeString('en-US', { hour: 'numeric', minute: '2-digit', hour12: true });
    }
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric' });
  };

  // Get selected emails data
  const selectedEmails = filteredEmails.filter(e => selectedIds.has(e.entryId));

  // Count active filters
  const activeFilterCount = [
    showUnreadOnly,
    showWithAttachments,
    showHighPriority,
    selectedSenders.length > 0,
    selectedDomains.length > 0,
    searchQuery.length > 0,
  ].filter(Boolean).length;

  return (
    <Box sx={{ display: 'flex', height: '100%', gap: 2 }}>
      {/* Main Email List */}
      <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', minWidth: 0 }}>
        {/* Controls Bar */}
        <Paper sx={{ p: 2, mb: 2, background: '#1a1a1f', border: '1px solid #27272a' }}>
          <Stack spacing={2}>
            {/* Top Row: Date Range, Search, Refresh */}
            <Stack direction="row" spacing={2} alignItems="center" flexWrap="wrap">
              <FormControl size="small" sx={{ minWidth: 120 }}>
                <InputLabel>Date Range</InputLabel>
                <Select
                  value={dateRange}
                  label="Date Range"
                  onChange={(e) => setDateRange(e.target.value)}
                >
                  <MenuItem value="24h">Last 24 Hours</MenuItem>
                  <MenuItem value="48h">Last 48 Hours</MenuItem>
                  <MenuItem value="7d">Last 7 Days</MenuItem>
                  <MenuItem value="30d">Last 30 Days</MenuItem>
                </Select>
              </FormControl>
              
              <TextField
                size="small"
                placeholder="Search emails..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                sx={{ flex: 1, minWidth: 200 }}
                InputProps={{
                  startAdornment: (
                    <InputAdornment position="start">
                      <SearchIcon sx={{ color: '#71717a' }} />
                    </InputAdornment>
                  ),
                  endAdornment: searchQuery && (
                    <InputAdornment position="end">
                      <IconButton size="small" onClick={() => setSearchQuery('')}>
                        <ClearIcon fontSize="small" />
                      </IconButton>
                    </InputAdornment>
                  ),
                }}
              />
              
              <Button
                variant="contained"
                startIcon={loading ? <CircularProgress size={18} color="inherit" /> : <RefreshIcon />}
                onClick={fetchEmails}
                disabled={loading}
              >
                {loading ? 'Loading...' : 'Load Emails'}
              </Button>
            </Stack>

            {/* Filter Toggles */}
            <Stack direction="row" spacing={2} alignItems="center" flexWrap="wrap">
              <FormControlLabel
                control={
                  <Switch 
                    checked={showUnreadOnly} 
                    onChange={(e) => setShowUnreadOnly(e.target.checked)}
                    size="small"
                  />
                }
                label={<Typography variant="body2">Unread Only</Typography>}
              />
              <FormControlLabel
                control={
                  <Switch 
                    checked={showWithAttachments} 
                    onChange={(e) => setShowWithAttachments(e.target.checked)}
                    size="small"
                  />
                }
                label={<Typography variant="body2">Has Attachments</Typography>}
              />
              <FormControlLabel
                control={
                  <Switch 
                    checked={showHighPriority} 
                    onChange={(e) => setShowHighPriority(e.target.checked)}
                    size="small"
                  />
                }
                label={<Typography variant="body2">High Priority</Typography>}
              />
              
              {activeFilterCount > 0 && (
                <Button 
                  size="small" 
                  startIcon={<ClearIcon />}
                  onClick={clearAllFilters}
                  sx={{ color: '#ef4444' }}
                >
                  Clear Filters ({activeFilterCount})
                </Button>
              )}
            </Stack>

            {/* Domain/Sender Filter Chips */}
            {emails.length > 0 && (
              <Box>
                <Stack direction="row" spacing={1} flexWrap="wrap" useFlexGap>
                  <Typography variant="caption" sx={{ color: '#71717a', mr: 1, alignSelf: 'center' }}>
                    <DomainIcon fontSize="small" sx={{ verticalAlign: 'middle', mr: 0.5 }} />
                    Domains:
                  </Typography>
                  {uniqueDomains.map(({ domain, count }) => (
                    <Chip
                      key={domain}
                      label={`${domain} (${count})`}
                      size="small"
                      onClick={() => toggleDomainFilter(domain)}
                      sx={{
                        background: selectedDomains.includes(domain) ? '#8b5cf6' : '#27272a',
                        color: selectedDomains.includes(domain) ? '#fff' : '#a1a1aa',
                        '&:hover': { background: selectedDomains.includes(domain) ? '#7c3aed' : '#3f3f46' },
                      }}
                    />
                  ))}
                </Stack>
              </Box>
            )}
          </Stack>
        </Paper>

        {/* Loading Progress */}
        {loading && (
          <LinearProgress 
            variant="determinate" 
            value={loadProgress} 
            sx={{ mb: 1, borderRadius: 1 }}
          />
        )}

        {/* Email Count & Selection Controls */}
        {emails.length > 0 && (
          <Stack direction="row" spacing={2} alignItems="center" sx={{ mb: 1, px: 1 }}>
            <Typography variant="body2" sx={{ color: '#71717a' }}>
              Showing {filteredEmails.length} of {emails.length} emails
            </Typography>
            <Divider orientation="vertical" flexItem />
            <Typography variant="body2" sx={{ color: selectedIds.size > 0 ? '#8b5cf6' : '#71717a' }}>
              {selectedIds.size} selected
            </Typography>
            <Button size="small" onClick={selectAll} startIcon={<SelectAllIcon />}>
              Select All
            </Button>
            {selectedIds.size > 0 && (
              <Button size="small" onClick={deselectAll} sx={{ color: '#71717a' }}>
                Deselect
              </Button>
            )}
          </Stack>
        )}

        {/* Action Buttons for Selected */}
        {selectedIds.size > 0 && (
          <Paper sx={{ p: 1.5, mb: 2, background: 'rgba(139, 92, 246, 0.1)', border: '1px solid rgba(139, 92, 246, 0.3)' }}>
            <Stack direction="row" spacing={1} alignItems="center" flexWrap="wrap">
              <Typography variant="body2" sx={{ color: '#8b5cf6', fontWeight: 600 }}>
                Actions for {selectedIds.size} email{selectedIds.size > 1 ? 's' : ''}:
              </Typography>
              <Button size="small" variant="outlined" startIcon={<SummarizeIcon />}>
                Summarize
              </Button>
              <Button size="small" variant="outlined" startIcon={<ActionItemsIcon />}>
                Extract Actions
              </Button>
              <Button size="small" variant="outlined" startIcon={<AIIcon />}>
                Custom Prompt
              </Button>
            </Stack>
          </Paper>
        )}

        {/* Email List */}
        <Paper 
          sx={{ 
            flex: 1, 
            overflow: 'auto', 
            background: '#0f0f12', 
            border: '1px solid #27272a',
          }}
        >
          {emails.length === 0 && !loading ? (
            <Box sx={{ p: 4, textAlign: 'center' }}>
              <EmailIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
              <Typography variant="h6" sx={{ color: '#71717a' }}>
                No Emails Loaded
              </Typography>
              <Typography variant="body2" sx={{ color: '#52525b', mb: 2 }}>
                Click "Load Emails" to fetch your recent inbox
              </Typography>
            </Box>
          ) : filteredEmails.length === 0 && emails.length > 0 ? (
            <Box sx={{ p: 4, textAlign: 'center' }}>
              <FilterIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
              <Typography variant="h6" sx={{ color: '#71717a' }}>
                No Matching Emails
              </Typography>
              <Typography variant="body2" sx={{ color: '#52525b' }}>
                Try adjusting your filters
              </Typography>
            </Box>
          ) : (
            <Box>
              {filteredEmails.map((email, index) => (
                <Box
                  key={email.entryId}
                  onClick={() => setPreviewEmail(email)}
                  sx={{
                    display: 'flex',
                    alignItems: 'flex-start',
                    gap: 1.5,
                    p: 1.5,
                    borderBottom: '1px solid #1f1f23',
                    cursor: 'pointer',
                    background: previewEmail?.entryId === email.entryId 
                      ? 'rgba(139, 92, 246, 0.1)' 
                      : email.isRead ? 'transparent' : 'rgba(59, 130, 246, 0.05)',
                    '&:hover': {
                      background: previewEmail?.entryId === email.entryId 
                        ? 'rgba(139, 92, 246, 0.15)'
                        : 'rgba(255, 255, 255, 0.02)',
                    },
                  }}
                >
                  {/* Checkbox */}
                  <Checkbox
                    checked={selectedIds.has(email.entryId)}
                    onClick={(e) => {
                      e.stopPropagation();
                      toggleSelection(email.entryId);
                    }}
                    size="small"
                    sx={{ p: 0.5, mt: 0.5 }}
                  />
                  
                  {/* Read/Unread indicator */}
                  <Box sx={{ width: 8, height: 8, borderRadius: '50%', mt: 1, flexShrink: 0,
                    background: email.isRead ? 'transparent' : '#3b82f6',
                    border: email.isRead ? '1px solid #3f3f46' : 'none',
                  }} />
                  
                  {/* Email Content */}
                  <Box sx={{ flex: 1, minWidth: 0 }}>
                    <Stack direction="row" alignItems="center" spacing={1} sx={{ mb: 0.5 }}>
                      <Typography 
                        variant="body2" 
                        sx={{ 
                          fontWeight: email.isRead ? 400 : 600,
                          color: email.isRead ? '#a1a1aa' : '#e4e4e7',
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                          whiteSpace: 'nowrap',
                          maxWidth: 200,
                        }}
                      >
                        {email.senderName || email.senderEmail}
                      </Typography>
                      {email.importance === 'high' && (
                        <HighPriorityIcon sx={{ fontSize: 14, color: '#ef4444' }} />
                      )}
                      {email.hasAttachments && (
                        <Badge badgeContent={email.attachmentCount} color="primary" max={9}>
                          <AttachmentIcon sx={{ fontSize: 14, color: '#71717a' }} />
                        </Badge>
                      )}
                      <Typography variant="caption" sx={{ color: '#52525b', ml: 'auto !important', flexShrink: 0 }}>
                        {formatTime(email.receivedTime)}
                      </Typography>
                    </Stack>
                    <Typography 
                      variant="body2" 
                      sx={{ 
                        fontWeight: email.isRead ? 400 : 500,
                        color: email.isRead ? '#a1a1aa' : '#e4e4e7',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap',
                        mb: 0.5,
                      }}
                    >
                      {email.subject}
                    </Typography>
                    <Typography 
                      variant="caption" 
                      sx={{ 
                        color: '#71717a',
                        overflow: 'hidden',
                        textOverflow: 'ellipsis',
                        whiteSpace: 'nowrap',
                        display: 'block',
                      }}
                    >
                      {email.bodyPreview || '(No preview)'}
                    </Typography>
                  </Box>
                </Box>
              ))}
            </Box>
          )}
        </Paper>
      </Box>

      {/* Preview Pane */}
      <Paper 
        sx={{ 
          width: 380, 
          flexShrink: 0, 
          display: 'flex', 
          flexDirection: 'column',
          background: '#1a1a1f', 
          border: '1px solid #27272a',
          overflow: 'hidden',
        }}
      >
        {previewEmail ? (
          <>
            {/* Preview Header */}
            <Box sx={{ p: 2, borderBottom: '1px solid #27272a' }}>
              <Typography variant="subtitle1" sx={{ fontWeight: 600, mb: 1, color: '#e4e4e7' }}>
                {previewEmail.subject}
              </Typography>
              <Stack direction="row" alignItems="center" spacing={1} sx={{ mb: 0.5 }}>
                <PersonIcon sx={{ fontSize: 16, color: '#71717a' }} />
                <Typography variant="body2" sx={{ color: '#a1a1aa' }}>
                  {previewEmail.senderName}
                </Typography>
              </Stack>
              <Typography variant="caption" sx={{ color: '#52525b' }}>
                &lt;{previewEmail.senderEmail}&gt;
              </Typography>
              <Stack direction="row" spacing={2} sx={{ mt: 1 }}>
                <Typography variant="caption" sx={{ color: '#71717a' }}>
                  To: {previewEmail.to || 'Me'}
                </Typography>
                {previewEmail.cc && (
                  <Typography variant="caption" sx={{ color: '#71717a' }}>
                    CC: {previewEmail.cc}
                  </Typography>
                )}
              </Stack>
              <Typography variant="caption" sx={{ color: '#52525b', display: 'block', mt: 0.5 }}>
                {new Date(previewEmail.receivedTime).toLocaleString()}
              </Typography>
              
              {/* Badges */}
              <Stack direction="row" spacing={1} sx={{ mt: 1 }}>
                {previewEmail.importance === 'high' && (
                  <Chip label="High Priority" size="small" sx={{ background: 'rgba(239, 68, 68, 0.2)', color: '#ef4444' }} />
                )}
                {previewEmail.hasAttachments && (
                  <Chip 
                    icon={<AttachmentIcon sx={{ fontSize: 14 }} />}
                    label={`${previewEmail.attachmentCount} file${previewEmail.attachmentCount > 1 ? 's' : ''}`} 
                    size="small" 
                    sx={{ background: 'rgba(59, 130, 246, 0.2)', color: '#3b82f6' }} 
                  />
                )}
                {!previewEmail.isRead && (
                  <Chip label="Unread" size="small" sx={{ background: 'rgba(34, 197, 94, 0.2)', color: '#22c55e' }} />
                )}
              </Stack>
            </Box>
            
            {/* Preview Body */}
            <Box sx={{ flex: 1, overflow: 'auto', p: 2 }}>
              <Typography variant="body2" sx={{ color: '#a1a1aa', whiteSpace: 'pre-wrap', lineHeight: 1.6 }}>
                {previewEmail.bodyPreview}
                {previewEmail.bodyPreview?.endsWith('...') && (
                  <Typography component="span" sx={{ color: '#71717a', fontStyle: 'italic' }}>
                    {'\n\n'}[Preview only - click "Load Full Email" for complete content]
                  </Typography>
                )}
              </Typography>
            </Box>
            
            {/* Preview Actions */}
            <Box sx={{ p: 2, borderTop: '1px solid #27272a' }}>
              <Stack spacing={1}>
                <Button 
                  variant="contained" 
                  fullWidth 
                  startIcon={<AIIcon />}
                  sx={{ background: 'linear-gradient(135deg, #8b5cf6 0%, #ec4899 100%)' }}
                >
                  Analyze This Email
                </Button>
                <Button 
                  variant="outlined" 
                  fullWidth
                >
                  Load Full Email
                </Button>
              </Stack>
            </Box>
          </>
        ) : (
          <Box sx={{ flex: 1, display: 'flex', alignItems: 'center', justifyContent: 'center', p: 3 }}>
            <Box sx={{ textAlign: 'center' }}>
              <EmailIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
              <Typography variant="body1" sx={{ color: '#71717a' }}>
                Select an email to preview
              </Typography>
            </Box>
          </Box>
        )}
      </Paper>
    </Box>
  );
};

export default InboxTab;

