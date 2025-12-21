import React, { useState, useCallback } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  Stack,
  IconButton,
  Tooltip,
  CircularProgress,
  Chip,
  Divider,
  List,
  ListItem,
  ListItemIcon,
  ListItemText,
  ListItemSecondaryAction,
  Checkbox,
  Menu,
  MenuItem,
  ListItemButton,
} from '@mui/material';
import {
  AttachFile as AttachmentIcon,
  Upload as UploadIcon,
  Description as DocIcon,
  Image as ImageIcon,
  TableChart as CsvIcon,
  Code as CodeIcon,
  PictureAsPdf as PdfIcon,
  Delete as DeleteIcon,
  AutoAwesome as AIIcon,
  ContentCopy as CopyIcon,
  Summarize as SummarizeIcon,
  ListAlt as ExtractIcon,
  QuestionAnswer as QAIcon,
  Close as CloseIcon,
  FolderOpen as FolderIcon,
  CloudUpload as CloudUploadIcon,
  CheckCircle as CheckIcon,
  TextFields as OcrIcon,
  Warning as WarningIcon,
} from '@mui/icons-material';

// Get icon for file type
function getFileIcon(extension) {
  const ext = extension?.toLowerCase();
  if (['.pdf'].includes(ext)) return <PdfIcon sx={{ color: '#ef4444' }} />;
  if (['.docx', '.doc'].includes(ext)) return <DocIcon sx={{ color: '#3b82f6' }} />;
  if (['.png', '.jpg', '.jpeg', '.gif', '.webp'].includes(ext)) return <ImageIcon sx={{ color: '#22c55e' }} />;
  if (['.csv', '.xlsx', '.xls'].includes(ext)) return <CsvIcon sx={{ color: '#22c55e' }} />;
  if (['.js', '.ts', '.py', '.java', '.json', '.xml', '.html'].includes(ext)) return <CodeIcon sx={{ color: '#f59e0b' }} />;
  return <DocIcon sx={{ color: '#71717a' }} />;
}

// Format file size
function formatFileSize(bytes) {
  if (!bytes) return '0 B';
  const k = 1024;
  const sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}

function FilesTab({ activeEmail, selectedEmail, showMessage, isElectron }) {
  const [files, setFiles] = useState([]);
  const [attachments, setAttachments] = useState([]);
  const [selectedFiles, setSelectedFiles] = useState([]);
  const [loading, setLoading] = useState({
    attachments: false,
    upload: false,
    analyze: false,
  });
  const [analysisResult, setAnalysisResult] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const [analysisMenuAnchor, setAnalysisMenuAnchor] = useState(null);

  // Get current email (active or selected)
  const currentEmail = activeEmail || selectedEmail;

  // Fetch attachments from current email
  const handleFetchAttachments = async () => {
    if (!isElectron || !currentEmail?.entryId) {
      showMessage('warning', 'Please select an email with attachments first');
      return;
    }

    setLoading(prev => ({ ...prev, attachments: true }));
    try {
      const result = await window.electronAPI.getEmailAttachments(currentEmail.entryId);
      if (result.success) {
        setAttachments(result.attachments || []);
        if (result.attachments?.length > 0) {
          showMessage('success', `Found ${result.attachments.length} attachment(s)`);
        } else {
          showMessage('info', 'No attachments found in this email');
        }
      } else {
        showMessage('error', result.message);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, attachments: false }));
    }
  };

  // Open file dialog
  const handleBrowseFiles = async () => {
    if (!isElectron) return;
    
    try {
      const result = await window.electronAPI.openFileDialog();
      if (result.success && result.files.length > 0) {
        setFiles(prev => [...prev, ...result.files]);
        showMessage('success', `Added ${result.files.length} file(s)`);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Handle drag and drop
  const handleDragOver = useCallback((e) => {
    e.preventDefault();
    setDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
  }, []);

  const handleDrop = useCallback((e) => {
    e.preventDefault();
    setDragOver(false);
    
    const droppedFiles = Array.from(e.dataTransfer.files).map(file => ({
      path: file.path,
      name: file.name,
      size: file.size,
      extension: '.' + file.name.split('.').pop().toLowerCase(),
    }));
    
    if (droppedFiles.length > 0) {
      setFiles(prev => [...prev, ...droppedFiles]);
      showMessage('success', `Added ${droppedFiles.length} file(s)`);
    }
  }, [showMessage]);

  // Toggle file selection
  const toggleFileSelection = (filePath) => {
    setSelectedFiles(prev => 
      prev.includes(filePath) 
        ? prev.filter(f => f !== filePath)
        : [...prev, filePath]
    );
  };

  // Remove file from list
  const removeFile = (filePath) => {
    setFiles(prev => prev.filter(f => f.path !== filePath));
    setSelectedFiles(prev => prev.filter(f => f !== filePath));
  };

  // Analyze selected files
  const handleAnalyze = async (analysisType) => {
    setAnalysisMenuAnchor(null);
    
    const allFiles = [
      ...attachments.map(a => ({ path: a.savedPath, name: a.fileName, type: 'attachment' })),
      ...files.map(f => ({ path: f.path, name: f.name, type: 'uploaded' })),
    ].filter(f => selectedFiles.includes(f.path));

    if (allFiles.length === 0) {
      showMessage('warning', 'Please select at least one file to analyze');
      return;
    }

    setLoading(prev => ({ ...prev, analyze: true }));
    setAnalysisResult(null);

    try {
      const results = [];
      
      for (const file of allFiles) {
        const result = await window.electronAPI.analyzeFile(file.path, analysisType);
        results.push({
          fileName: file.name,
          ...result,
        });
      }

      setAnalysisResult({
        type: analysisType,
        files: results,
      });

      const successCount = results.filter(r => r.success).length;
      showMessage('success', `Analyzed ${successCount} of ${allFiles.length} file(s)`);
      
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    } finally {
      setLoading(prev => ({ ...prev, analyze: false }));
    }
  };

  // Copy analysis result
  const copyResult = () => {
    if (analysisResult?.files) {
      const text = analysisResult.files
        .map(r => `=== ${r.fileName} ===\n${r.content || r.error || 'No content'}`)
        .join('\n\n');
      navigator.clipboard.writeText(text);
      showMessage('success', 'Copied to clipboard!');
    }
  };

  // All files combined
  const allFiles = [
    ...attachments.map(a => ({ 
      path: a.savedPath, 
      name: a.fileName, 
      size: a.size, 
      extension: a.extension,
      source: 'attachment' 
    })),
    ...files.map(f => ({ 
      ...f, 
      source: 'uploaded' 
    })),
  ];

  return (
    <Stack spacing={3} sx={{ height: '100%', overflow: 'auto' }}>
      {/* Action Buttons */}
      <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap' }}>
        <Button
          variant="contained"
          startIcon={loading.attachments ? <CircularProgress size={18} color="inherit" /> : <AttachmentIcon />}
          onClick={handleFetchAttachments}
          disabled={loading.attachments || !currentEmail?.hasAttachments}
        >
          {loading.attachments ? 'Extracting...' : 'Get Email Attachments'}
        </Button>
        <Button
          variant="outlined"
          startIcon={<FolderIcon />}
          onClick={handleBrowseFiles}
          sx={{
            borderColor: '#3f3f46',
            color: 'text.secondary',
            '&:hover': { borderColor: 'primary.main', color: 'primary.main' },
          }}
        >
          Browse Files
        </Button>
        <Button
          variant="outlined"
          startIcon={loading.analyze ? <CircularProgress size={18} color="inherit" /> : <AIIcon />}
          onClick={(e) => setAnalysisMenuAnchor(e.currentTarget)}
          disabled={selectedFiles.length === 0 || loading.analyze}
          sx={{
            borderColor: selectedFiles.length > 0 ? 'primary.main' : '#3f3f46',
            color: selectedFiles.length > 0 ? 'primary.main' : 'text.secondary',
            '&:hover': { borderColor: 'primary.main', color: 'primary.main' },
          }}
        >
          {loading.analyze ? 'Analyzing...' : `Analyze Selected (${selectedFiles.length})`}
        </Button>
        
        {/* Analysis Menu */}
        <Menu
          anchorEl={analysisMenuAnchor}
          open={Boolean(analysisMenuAnchor)}
          onClose={() => setAnalysisMenuAnchor(null)}
          PaperProps={{
            sx: { background: '#18181b', border: '1px solid #27272a', minWidth: 220 }
          }}
        >
          <MenuItem onClick={() => handleAnalyze('summarize')}>
            <ListItemIcon><SummarizeIcon sx={{ color: 'primary.main' }} /></ListItemIcon>
            <ListItemText>Summarize</ListItemText>
          </MenuItem>
          <MenuItem onClick={() => handleAnalyze('extract')}>
            <ListItemIcon><ExtractIcon sx={{ color: '#22c55e' }} /></ListItemIcon>
            <ListItemText>Extract Key Info</ListItemText>
          </MenuItem>
          <MenuItem onClick={() => handleAnalyze('questions')}>
            <ListItemIcon><QAIcon sx={{ color: '#3b82f6' }} /></ListItemIcon>
            <ListItemText>Generate Q&A</ListItemText>
          </MenuItem>
          <Divider sx={{ my: 1, borderColor: '#27272a' }} />
          <MenuItem disabled sx={{ opacity: 0.7 }}>
            <ListItemIcon><OcrIcon sx={{ color: '#a855f7' }} /></ListItemIcon>
            <ListItemText 
              primary="OCR Enabled" 
              secondary="Images auto-scanned for text"
              secondaryTypographyProps={{ sx: { fontSize: '0.7rem', color: '#71717a' } }}
            />
          </MenuItem>
        </Menu>
      </Box>

      {/* Drag & Drop Zone */}
      <Paper
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        sx={{
          p: 4,
          textAlign: 'center',
          border: dragOver ? '2px dashed #f97316' : '2px dashed #3f3f46',
          background: dragOver ? 'rgba(249, 115, 22, 0.1)' : 'transparent',
          borderRadius: 2,
          transition: 'all 0.2s ease',
          cursor: 'pointer',
          '&:hover': {
            borderColor: '#f97316',
            background: 'rgba(249, 115, 22, 0.05)',
          },
        }}
        onClick={handleBrowseFiles}
      >
        <CloudUploadIcon sx={{ fontSize: 48, color: dragOver ? 'primary.main' : '#3f3f46', mb: 2 }} />
        <Typography variant="h6" sx={{ color: dragOver ? 'primary.main' : 'text.secondary', mb: 1 }}>
          {dragOver ? 'Drop files here' : 'Drag & drop files here'}
        </Typography>
        <Typography variant="body2" sx={{ color: '#71717a' }}>
          or click to browse • PDF, Word, Text, CSV, Images
        </Typography>
      </Paper>

      {/* File List */}
      {allFiles.length > 0 && (
        <Paper sx={{ background: '#1a1a1f', border: '1px solid #27272a', borderRadius: 2 }}>
          <Box sx={{ p: 2, borderBottom: '1px solid #27272a', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
              Files ({allFiles.length})
            </Typography>
            <Stack direction="row" spacing={1}>
              <Button
                size="small"
                onClick={() => setSelectedFiles(allFiles.map(f => f.path))}
                sx={{ color: 'text.secondary' }}
              >
                Select All
              </Button>
              <Button
                size="small"
                onClick={() => setSelectedFiles([])}
                sx={{ color: 'text.secondary' }}
              >
                Clear Selection
              </Button>
            </Stack>
          </Box>
          <List dense sx={{ maxHeight: 300, overflow: 'auto' }}>
            {allFiles.map((file, index) => (
              <ListItem
                key={file.path + index}
                sx={{
                  borderBottom: index < allFiles.length - 1 ? '1px solid #27272a' : 'none',
                  background: selectedFiles.includes(file.path) ? 'rgba(249, 115, 22, 0.1)' : 'transparent',
                }}
              >
                <ListItemButton onClick={() => toggleFileSelection(file.path)} sx={{ py: 1.5 }}>
                  <Checkbox
                    checked={selectedFiles.includes(file.path)}
                    sx={{ mr: 1 }}
                  />
                  <ListItemIcon sx={{ minWidth: 40 }}>
                    {getFileIcon(file.extension)}
                  </ListItemIcon>
                  <ListItemText
                    primary={file.name}
                    secondary={
                      <Stack direction="row" spacing={1} alignItems="center">
                        <Typography variant="caption" sx={{ color: '#71717a' }}>
                          {formatFileSize(file.size)}
                        </Typography>
                        <Chip
                          label={file.source === 'attachment' ? 'Attachment' : 'Uploaded'}
                          size="small"
                          sx={{
                            height: 18,
                            fontSize: '0.65rem',
                            background: file.source === 'attachment' ? 'rgba(59, 130, 246, 0.2)' : 'rgba(34, 197, 94, 0.2)',
                            color: file.source === 'attachment' ? '#3b82f6' : '#22c55e',
                          }}
                        />
                      </Stack>
                    }
                  />
                </ListItemButton>
                <ListItemSecondaryAction>
                  {file.source === 'uploaded' && (
                    <IconButton size="small" onClick={() => removeFile(file.path)}>
                      <DeleteIcon fontSize="small" sx={{ color: '#71717a' }} />
                    </IconButton>
                  )}
                </ListItemSecondaryAction>
              </ListItem>
            ))}
          </List>
        </Paper>
      )}

      {/* Analysis Results */}
      {analysisResult && (
        <Paper
          sx={{
            p: 3,
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
                {analysisResult.type === 'summarize' && '📄 File Summary'}
                {analysisResult.type === 'extract' && '📋 Extracted Information'}
                {analysisResult.type === 'questions' && '❓ Q&A'}
              </Typography>
            </Stack>
            <Stack direction="row" spacing={1}>
              <Tooltip title="Copy to clipboard">
                <IconButton onClick={copyResult} size="small" sx={{ color: 'primary.main' }}>
                  <CopyIcon />
                </IconButton>
              </Tooltip>
              <Tooltip title="Close">
                <IconButton onClick={() => setAnalysisResult(null)} size="small" sx={{ color: 'text.secondary' }}>
                  <CloseIcon />
                </IconButton>
              </Tooltip>
            </Stack>
          </Stack>
          
          <Divider sx={{ mb: 2, borderColor: 'rgba(249, 115, 22, 0.3)' }} />
          
          {analysisResult.files.map((result, index) => (
            <Box key={index} sx={{ mb: index < analysisResult.files.length - 1 ? 3 : 0 }}>
              <Stack direction="row" alignItems="center" spacing={1} mb={1} flexWrap="wrap">
                {getFileIcon('.' + result.fileName?.split('.').pop())}
                <Typography variant="subtitle2" sx={{ color: 'primary.main', fontWeight: 600 }}>
                  {result.fileName}
                </Typography>
                {result.success ? (
                  <CheckIcon sx={{ fontSize: 16, color: '#22c55e' }} />
                ) : (
                  <Chip label="Error" size="small" color="error" sx={{ height: 18 }} />
                )}
                {result.ocrUsed && (
                  <Chip 
                    icon={<OcrIcon sx={{ fontSize: 14 }} />}
                    label="OCR" 
                    size="small" 
                    sx={{ 
                      height: 20, 
                      fontSize: '0.65rem',
                      background: 'rgba(168, 85, 247, 0.2)',
                      color: '#a855f7',
                    }} 
                  />
                )}
                {result.isScanned && (
                  <Chip 
                    icon={<WarningIcon sx={{ fontSize: 14 }} />}
                    label="Scanned PDF" 
                    size="small" 
                    sx={{ 
                      height: 20, 
                      fontSize: '0.65rem',
                      background: 'rgba(245, 158, 11, 0.2)',
                      color: '#f59e0b',
                    }} 
                  />
                )}
                {result.pages && (
                  <Chip 
                    label={`${result.pages} pages`} 
                    size="small" 
                    sx={{ 
                      height: 18, 
                      fontSize: '0.65rem',
                      background: 'rgba(59, 130, 246, 0.2)',
                      color: '#3b82f6',
                    }} 
                  />
                )}
              </Stack>
              
              <Box
                sx={{
                  p: 2.5,
                  background: result.isScanned ? 'rgba(245, 158, 11, 0.1)' : 'rgba(15, 15, 18, 0.8)',
                  border: result.isScanned ? '1px solid rgba(245, 158, 11, 0.3)' : 'none',
                  borderRadius: 2,
                  maxHeight: 350,
                  overflow: 'auto',
                  fontFamily: '"Plus Jakarta Sans", sans-serif',
                  fontSize: '0.95rem',
                  lineHeight: 1.9,
                  whiteSpace: 'pre-wrap',
                  wordBreak: 'break-word',
                  color: result.success ? (result.isScanned ? '#f59e0b' : '#e4e4e7') : '#ef4444',
                }}
              >
                {result.content || result.error || 'No content'}
              </Box>
            </Box>
          ))}
        </Paper>
      )}

      {/* Empty State */}
      {allFiles.length === 0 && !analysisResult && (
        <Paper
          sx={{
            p: 4,
            background: '#1a1a1f',
            border: '1px dashed #3f3f46',
            textAlign: 'center',
            borderRadius: 2,
          }}
        >
          <AttachmentIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
          <Typography variant="h6" sx={{ color: 'text.secondary', mb: 1 }}>
            No Files Selected
          </Typography>
          <Typography variant="body2" sx={{ color: '#71717a', mb: 2 }}>
            {currentEmail?.hasAttachments 
              ? 'Click "Get Email Attachments" to extract files from the current email, or drag & drop files here.'
              : 'Drag & drop files here or browse to upload files for AI analysis.'}
          </Typography>
          <Typography variant="caption" sx={{ color: '#52525b' }}>
            Supported: PDF, Word (.docx), Text, CSV, JSON, Images (PNG, JPG)
          </Typography>
        </Paper>
      )}
    </Stack>
  );
}

export default FilesTab;

