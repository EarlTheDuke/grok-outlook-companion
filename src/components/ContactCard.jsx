import React, { useState, useEffect } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  IconButton,
  Stack,
  Chip,
  Divider,
  CircularProgress,
  Tooltip,
  Collapse,
  Alert,
} from '@mui/material';
import {
  Person as PersonIcon,
  Business as BusinessIcon,
  Phone as PhoneIcon,
  Smartphone as MobileIcon,
  Email as EmailIcon,
  LocationOn as LocationIcon,
  Notes as NotesIcon,
  Search as SearchIcon,
  Close as CloseIcon,
  ContentCopy as CopyIcon,
  Save as SaveIcon,
  Refresh as RefreshIcon,
  ExpandMore as ExpandIcon,
  ExpandLess as CollapseIcon,
  History as HistoryIcon,
  LinkedIn as LinkedInIcon,
  Article as NewsIcon,
  Lightbulb as InsightIcon,
  Chat as TalkingIcon,
} from '@mui/icons-material';

function ContactCard({ email, senderName, onClose, showMessage }) {
  const [loading, setLoading] = useState({ contact: true, history: false, research: false });
  const [contactData, setContactData] = useState(null);
  const [emailHistory, setEmailHistory] = useState(null);
  const [research, setResearch] = useState(null);
  const [showResearch, setShowResearch] = useState(false);
  const [savingNotes, setSavingNotes] = useState(false);

  // Load contact data when email changes
  useEffect(() => {
    if (email) {
      loadContactData();
    }
  }, [email]);

  const loadContactData = async () => {
    if (!window.electronAPI) return;
    
    setLoading(prev => ({ ...prev, contact: true, history: true }));
    
    try {
      // Lookup contact
      const contactResult = await window.electronAPI.lookupContact(email);
      if (contactResult.success) {
        setContactData({
          found: contactResult.found,
          source: contactResult.source,
          ...contactResult.contact,
          // If not found, use sender name from email
          fullName: contactResult.contact?.fullName || senderName || 'Unknown',
          email: email,
        });
      }
      
      // Get email history
      const historyResult = await window.electronAPI.getEmailHistory(email);
      if (historyResult.success) {
        setEmailHistory(historyResult);
      }
    } catch (error) {
      console.error('Error loading contact:', error);
      showMessage?.('error', 'Failed to load contact info');
    } finally {
      setLoading(prev => ({ ...prev, contact: false, history: false }));
    }
  };

  const handleResearch = async () => {
    if (!window.electronAPI || !contactData) return;
    
    setLoading(prev => ({ ...prev, research: true }));
    setShowResearch(true);
    
    try {
      const result = await window.electronAPI.researchContact(
        contactData.fullName,
        email,
        contactData.company,
        contactData.jobTitle
      );
      
      if (result.success) {
        setResearch({
          content: result.research,
          generatedAt: result.generatedAt,
        });
        showMessage?.('success', 'Research complete!');
      } else {
        showMessage?.('error', result.error || 'Research failed');
      }
    } catch (error) {
      console.error('Error researching contact:', error);
      showMessage?.('error', 'Research failed');
    } finally {
      setLoading(prev => ({ ...prev, research: false }));
    }
  };

  const handleSaveToNotes = async () => {
    if (!window.electronAPI || !research) return;
    
    setSavingNotes(true);
    try {
      const result = await window.electronAPI.saveContactNotes(
        email,
        research.content,
        true // append
      );
      
      if (result.success) {
        showMessage?.('success', 'Research saved to Outlook contact notes!');
      } else {
        showMessage?.('warning', result.message || 'Could not save notes');
      }
    } catch (error) {
      showMessage?.('error', 'Failed to save notes');
    } finally {
      setSavingNotes(false);
    }
  };

  const handleCopyResearch = () => {
    if (research?.content) {
      navigator.clipboard.writeText(research.content);
      showMessage?.('success', 'Research copied to clipboard');
    }
  };

  const formatPhoneNumber = (phone) => {
    if (!phone) return null;
    return phone;
  };

  return (
    <Paper
      sx={{
        background: 'linear-gradient(135deg, #18181b 0%, #1f1f23 100%)',
        border: '1px solid #27272a',
        borderRadius: 2,
        overflow: 'hidden',
        maxHeight: '80vh',
        display: 'flex',
        flexDirection: 'column',
      }}
    >
      {/* Header */}
      <Box
        sx={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          p: 2,
          borderBottom: '1px solid #27272a',
          background: 'rgba(249, 115, 22, 0.05)',
        }}
      >
        <Stack direction="row" alignItems="center" spacing={1}>
          <PersonIcon sx={{ color: '#f97316' }} />
          <Typography variant="h6" sx={{ fontWeight: 600 }}>
            Contact Profile
          </Typography>
        </Stack>
        <IconButton onClick={onClose} size="small" sx={{ color: '#71717a' }}>
          <CloseIcon />
        </IconButton>
      </Box>

      {/* Content */}
      <Box sx={{ flex: 1, overflow: 'auto', p: 2 }}>
        {loading.contact ? (
          <Box sx={{ display: 'flex', justifyContent: 'center', py: 4 }}>
            <CircularProgress size={32} />
          </Box>
        ) : (
          <Stack spacing={2}>
            {/* Contact Info Section */}
            <Box sx={{ display: 'flex', gap: 2 }}>
              {/* Left: Outlook Data */}
              <Box sx={{ flex: 1 }}>
                <Typography variant="caption" sx={{ color: '#71717a', textTransform: 'uppercase', letterSpacing: 1 }}>
                  {contactData?.found ? `From ${contactData.source}` : 'From Email'}
                </Typography>
                
                <Typography variant="h5" sx={{ fontWeight: 700, mt: 0.5 }}>
                  {contactData?.fullName || senderName || 'Unknown'}
                </Typography>
                
                {(contactData?.jobTitle || contactData?.company) && (
                  <Typography variant="body2" sx={{ color: '#a1a1aa', mt: 0.5 }}>
                    {contactData.jobTitle}{contactData.jobTitle && contactData.company ? ' @ ' : ''}{contactData.company}
                  </Typography>
                )}
                
                {contactData?.department && (
                  <Chip
                    label={contactData.department}
                    size="small"
                    sx={{ mt: 1, background: 'rgba(59, 130, 246, 0.2)', color: '#60a5fa' }}
                  />
                )}
              </Box>
              
              {/* Right: Quick Stats */}
              <Box sx={{ textAlign: 'right' }}>
                {emailHistory && (
                  <Box>
                    <Typography variant="h4" sx={{ fontWeight: 700, color: '#f97316' }}>
                      {emailHistory.totalCount}
                    </Typography>
                    <Typography variant="caption" sx={{ color: '#71717a' }}>
                      emails exchanged
                    </Typography>
                    {emailHistory.lastEmailDate && (
                      <Typography variant="caption" sx={{ display: 'block', color: '#52525b' }}>
                        Last: {emailHistory.lastEmailDate}
                      </Typography>
                    )}
                  </Box>
                )}
              </Box>
            </Box>

            <Divider sx={{ borderColor: '#27272a' }} />

            {/* Contact Details */}
            <Stack spacing={1}>
              <Stack direction="row" alignItems="center" spacing={1}>
                <EmailIcon sx={{ fontSize: 18, color: '#71717a' }} />
                <Typography variant="body2" sx={{ color: '#a1a1aa' }}>
                  {/* Hide ugly Exchange X500 format, show clean email */}
                  {email && !email.startsWith('/O=') ? email : (contactData?.email && !contactData.email.startsWith('/O=') ? contactData.email : 'Exchange user')}
                </Typography>
              </Stack>
              
              {contactData?.phoneWork && (
                <Stack direction="row" alignItems="center" spacing={1}>
                  <PhoneIcon sx={{ fontSize: 18, color: '#71717a' }} />
                  <Typography variant="body2" sx={{ color: '#a1a1aa' }}>
                    {contactData.phoneWork} <span style={{ color: '#52525b' }}>(Work)</span>
                  </Typography>
                </Stack>
              )}
              
              {contactData?.phoneMobile && (
                <Stack direction="row" alignItems="center" spacing={1}>
                  <MobileIcon sx={{ fontSize: 18, color: '#71717a' }} />
                  <Typography variant="body2" sx={{ color: '#a1a1aa' }}>
                    {contactData.phoneMobile} <span style={{ color: '#52525b' }}>(Mobile)</span>
                  </Typography>
                </Stack>
              )}
              
              {contactData?.address && (
                <Stack direction="row" alignItems="flex-start" spacing={1}>
                  <LocationIcon sx={{ fontSize: 18, color: '#71717a', mt: 0.3 }} />
                  <Typography variant="body2" sx={{ color: '#a1a1aa' }}>
                    {contactData.address}
                  </Typography>
                </Stack>
              )}
              
              {contactData?.notes && (
                <Box sx={{ mt: 1, p: 1.5, background: '#0f0f12', borderRadius: 1, border: '1px solid #27272a' }}>
                  <Stack direction="row" alignItems="center" spacing={0.5} sx={{ mb: 0.5 }}>
                    <NotesIcon sx={{ fontSize: 14, color: '#71717a' }} />
                    <Typography variant="caption" sx={{ color: '#71717a' }}>Notes</Typography>
                  </Stack>
                  <Typography variant="body2" sx={{ color: '#a1a1aa', whiteSpace: 'pre-wrap', fontSize: '0.8rem' }}>
                    {contactData.notes.length > 200 ? contactData.notes.substring(0, 200) + '...' : contactData.notes}
                  </Typography>
                </Box>
              )}
            </Stack>

            <Divider sx={{ borderColor: '#27272a' }} />

            {/* Research Section */}
            <Box>
              <Button
                fullWidth
                variant={showResearch ? 'outlined' : 'contained'}
                startIcon={loading.research ? <CircularProgress size={18} color="inherit" /> : <SearchIcon />}
                onClick={showResearch ? () => setShowResearch(false) : handleResearch}
                disabled={loading.research}
                sx={{
                  background: showResearch ? 'transparent' : 'linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%)',
                  '&:hover': {
                    background: showResearch ? 'rgba(59, 130, 246, 0.1)' : 'linear-gradient(135deg, #60a5fa 0%, #3b82f6 100%)',
                  },
                }}
              >
                {loading.research ? 'Researching...' : showResearch ? 'Hide Research' : '🔍 Deep Research (AI)'}
              </Button>
              
              <Collapse in={showResearch}>
                <Box sx={{ mt: 2 }}>
                  {loading.research ? (
                    <Box sx={{ textAlign: 'center', py: 4 }}>
                      <CircularProgress size={40} />
                      <Typography variant="body2" sx={{ color: '#71717a', mt: 2 }}>
                        Searching the web for information...
                      </Typography>
                    </Box>
                  ) : research ? (
                    <Box>
                      {/* Research Actions */}
                      <Stack direction="row" spacing={1} sx={{ mb: 2 }}>
                        <Button
                          size="small"
                          startIcon={<CopyIcon />}
                          onClick={handleCopyResearch}
                          sx={{ color: '#a1a1aa' }}
                        >
                          Copy
                        </Button>
                        <Button
                          size="small"
                          startIcon={savingNotes ? <CircularProgress size={14} /> : <SaveIcon />}
                          onClick={handleSaveToNotes}
                          disabled={savingNotes || !contactData?.found}
                          sx={{ color: '#a1a1aa' }}
                        >
                          Save to Outlook
                        </Button>
                        <Button
                          size="small"
                          startIcon={<RefreshIcon />}
                          onClick={handleResearch}
                          sx={{ color: '#a1a1aa' }}
                        >
                          Refresh
                        </Button>
                      </Stack>
                      
                      {/* Research Content */}
                      <Box
                        sx={{
                          p: 2,
                          background: '#0f0f12',
                          borderRadius: 1,
                          border: '1px solid #27272a',
                          maxHeight: 400,
                          overflow: 'auto',
                        }}
                      >
                        <Typography
                          variant="body2"
                          sx={{
                            color: '#d4d4d8',
                            whiteSpace: 'pre-wrap',
                            fontFamily: 'inherit',
                            lineHeight: 1.7,
                            '& strong': { color: '#f97316' },
                          }}
                          dangerouslySetInnerHTML={{
                            __html: research.content
                              .replace(/\*\*(.*?)\*\*/g, '<strong>$1</strong>')
                              .replace(/^### (.*$)/gm, '<h4 style="color: #f97316; margin: 16px 0 8px 0;">$1</h4>')
                              .replace(/^## (.*$)/gm, '<h3 style="color: #f97316; margin: 16px 0 8px 0;">$1</h3>')
                              .replace(/^# (.*$)/gm, '<h2 style="color: #f97316; margin: 16px 0 8px 0;">$1</h2>')
                              .replace(/^- (.*$)/gm, '<li style="margin-left: 16px;">$1</li>')
                              .replace(/^\d+\. (.*$)/gm, '<li style="margin-left: 16px;">$1</li>')
                          }}
                        />
                        <Typography variant="caption" sx={{ display: 'block', mt: 2, color: '#52525b' }}>
                          Generated: {new Date(research.generatedAt).toLocaleString()}
                        </Typography>
                      </Box>
                    </Box>
                  ) : (
                    <Alert severity="info" sx={{ background: 'rgba(59, 130, 246, 0.1)' }}>
                      Click "Deep Research" to search for information about this contact online.
                    </Alert>
                  )}
                </Box>
              </Collapse>
            </Box>

            {/* Email History Details */}
            {emailHistory && emailHistory.totalCount > 0 && (
              <Box sx={{ p: 1.5, background: '#0f0f12', borderRadius: 1, border: '1px solid #27272a' }}>
                <Stack direction="row" alignItems="center" spacing={1}>
                  <HistoryIcon sx={{ fontSize: 16, color: '#71717a' }} />
                  <Typography variant="caption" sx={{ color: '#71717a' }}>
                    Email History: {emailHistory.inboxCount} received, {emailHistory.sentCount} sent
                  </Typography>
                </Stack>
              </Box>
            )}
          </Stack>
        )}
      </Box>
    </Paper>
  );
}

export default ContactCard;

