import React, { useState, useEffect, useCallback } from 'react';
import {
  Box,
  Paper,
  Typography,
  Button,
  Stack,
  IconButton,
  Tooltip,
  TextField,
  Chip,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  FormControl,
  InputLabel,
  Select,
  MenuItem,
  List,
  ListItem,
  ListItemText,
  ListItemSecondaryAction,
  InputAdornment,
  Alert,
  Divider,
  ToggleButton,
  ToggleButtonGroup,
} from '@mui/material';
import {
  Add as AddIcon,
  Edit as EditIcon,
  Delete as DeleteIcon,
  Star as StarIcon,
  StarBorder as StarBorderIcon,
  Search as SearchIcon,
  ContentCopy as CopyIcon,
  FileUpload as ImportIcon,
  FileDownload as ExportIcon,
  Summarize as SummarizeIcon,
  Reply as ReplyIcon,
  Lightbulb as InsightIcon,
  Code as CustomIcon,
  Close as CloseIcon,
  AutoAwesome as AIIcon,
} from '@mui/icons-material';

// Category icons and colors
const CATEGORIES = {
  summarize: { icon: SummarizeIcon, color: '#f97316', label: 'Summarize' },
  reply: { icon: ReplyIcon, color: '#3b82f6', label: 'Reply' },
  insights: { icon: InsightIcon, color: '#22c55e', label: 'Insights' },
  custom: { icon: CustomIcon, color: '#a855f7', label: 'Custom' },
};

// Placeholder info
const PLACEHOLDERS = [
  { key: '{email_body}', desc: 'Full email text' },
  { key: '{subject}', desc: 'Email subject' },
  { key: '{sender}', desc: "Sender's name" },
  { key: '{sender_email}', desc: "Sender's email" },
  { key: '{to}', desc: 'Recipients' },
  { key: '{date}', desc: 'Email date' },
];

function PromptsTab({ showMessage, isElectron, onPromptsChange }) {
  const [prompts, setPrompts] = useState([]);
  const [loading, setLoading] = useState(true);
  const [searchQuery, setSearchQuery] = useState('');
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [dialogOpen, setDialogOpen] = useState(false);
  const [editingPrompt, setEditingPrompt] = useState(null);
  const [deleteConfirmOpen, setDeleteConfirmOpen] = useState(false);
  const [promptToDelete, setPromptToDelete] = useState(null);

  // Form state
  const [formData, setFormData] = useState({
    name: '',
    description: '',
    template: '',
    category: 'custom',
  });

  // Load prompts
  const loadPrompts = useCallback(async () => {
    if (!isElectron) return;
    
    setLoading(true);
    try {
      const result = await window.electronAPI.getPrompts();
      if (result.success) {
        setPrompts(result.prompts);
      } else {
        showMessage('error', result.error);
      }
    } catch (error) {
      showMessage('error', `Error loading prompts: ${error.message}`);
    } finally {
      setLoading(false);
    }
  }, [isElectron, showMessage]);

  useEffect(() => {
    loadPrompts();
  }, [loadPrompts]);

  // Filter prompts
  const filteredPrompts = prompts.filter(p => {
    const matchesSearch = !searchQuery || 
      p.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      p.description?.toLowerCase().includes(searchQuery.toLowerCase()) ||
      p.template.toLowerCase().includes(searchQuery.toLowerCase());
    
    const matchesCategory = categoryFilter === 'all' || p.category === categoryFilter;
    
    return matchesSearch && matchesCategory;
  });

  // Sort: favorites first, then by usage count
  const sortedPrompts = [...filteredPrompts].sort((a, b) => {
    if (a.isFavorite && !b.isFavorite) return -1;
    if (!a.isFavorite && b.isFavorite) return 1;
    return (b.usageCount || 0) - (a.usageCount || 0);
  });

  // Handle form changes
  const handleFormChange = (field, value) => {
    setFormData(prev => ({ ...prev, [field]: value }));
  };

  // Open dialog for new prompt
  const handleNewPrompt = () => {
    setEditingPrompt(null);
    setFormData({
      name: '',
      description: '',
      template: '',
      category: 'custom',
    });
    setDialogOpen(true);
  };

  // Open dialog for editing
  const handleEditPrompt = (prompt) => {
    setEditingPrompt(prompt);
    setFormData({
      name: prompt.name,
      description: prompt.description || '',
      template: prompt.template,
      category: prompt.category,
    });
    setDialogOpen(true);
  };

  // Save prompt
  const handleSavePrompt = async () => {
    if (!formData.name.trim() || !formData.template.trim()) {
      showMessage('warning', 'Name and template are required');
      return;
    }

    try {
      let result;
      if (editingPrompt) {
        result = await window.electronAPI.updatePrompt(editingPrompt.id, formData);
      } else {
        result = await window.electronAPI.addPrompt(formData);
      }

      if (result.success) {
        showMessage('success', editingPrompt ? 'Prompt updated!' : 'Prompt created!');
        setDialogOpen(false);
        loadPrompts();
        if (onPromptsChange) onPromptsChange();
      } else {
        showMessage('error', result.error);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Toggle favorite
  const handleToggleFavorite = async (prompt) => {
    try {
      const result = await window.electronAPI.updatePrompt(prompt.id, {
        isFavorite: !prompt.isFavorite,
      });
      if (result.success) {
        loadPrompts();
        if (onPromptsChange) onPromptsChange();
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Delete prompt
  const handleDeletePrompt = async () => {
    if (!promptToDelete) return;

    try {
      const result = await window.electronAPI.deletePrompt(promptToDelete.id);
      if (result.success) {
        showMessage('success', 'Prompt deleted');
        setDeleteConfirmOpen(false);
        setPromptToDelete(null);
        loadPrompts();
        if (onPromptsChange) onPromptsChange();
      } else {
        showMessage('error', result.error);
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Export prompts
  const handleExport = async () => {
    try {
      const result = await window.electronAPI.exportPrompts();
      if (result.success) {
        navigator.clipboard.writeText(result.data);
        showMessage('success', 'Prompts copied to clipboard! Paste into a .json file to save.');
      }
    } catch (error) {
      showMessage('error', `Error: ${error.message}`);
    }
  };

  // Import prompts
  const handleImport = async () => {
    try {
      const clipboardText = await navigator.clipboard.readText();
      const result = await window.electronAPI.importPrompts(clipboardText);
      if (result.success) {
        showMessage('success', `Imported ${result.count} prompt(s)!`);
        loadPrompts();
      } else {
        showMessage('error', result.error);
      }
    } catch (error) {
      showMessage('error', 'Failed to import. Make sure you have valid JSON in your clipboard.');
    }
  };

  // Insert placeholder into template
  const insertPlaceholder = (placeholder) => {
    setFormData(prev => ({
      ...prev,
      template: prev.template + placeholder,
    }));
  };

  return (
    <Stack spacing={3} sx={{ height: '100%', overflow: 'auto' }}>
      {/* Header Actions */}
      <Box sx={{ display: 'flex', gap: 2, flexWrap: 'wrap', alignItems: 'center' }}>
        <Button
          variant="contained"
          startIcon={<AddIcon />}
          onClick={handleNewPrompt}
        >
          New Prompt
        </Button>
        
        <TextField
          size="small"
          placeholder="Search prompts..."
          value={searchQuery}
          onChange={(e) => setSearchQuery(e.target.value)}
          InputProps={{
            startAdornment: (
              <InputAdornment position="start">
                <SearchIcon sx={{ color: 'text.secondary' }} />
              </InputAdornment>
            ),
          }}
          sx={{ minWidth: 200 }}
        />

        <ToggleButtonGroup
          value={categoryFilter}
          exclusive
          onChange={(e, v) => v && setCategoryFilter(v)}
          size="small"
        >
          <ToggleButton value="all">All</ToggleButton>
          {Object.entries(CATEGORIES).map(([key, { label }]) => (
            <ToggleButton key={key} value={key}>{label}</ToggleButton>
          ))}
        </ToggleButtonGroup>

        <Box sx={{ flexGrow: 1 }} />

        <Tooltip title="Import from clipboard">
          <IconButton onClick={handleImport} sx={{ color: 'text.secondary' }}>
            <ImportIcon />
          </IconButton>
        </Tooltip>
        <Tooltip title="Export to clipboard">
          <IconButton onClick={handleExport} sx={{ color: 'text.secondary' }}>
            <ExportIcon />
          </IconButton>
        </Tooltip>
      </Box>

      {/* Prompts List */}
      {sortedPrompts.length > 0 ? (
        <Paper sx={{ background: '#1a1a1f', border: '1px solid #27272a', borderRadius: 2 }}>
          <List sx={{ py: 0 }}>
            {sortedPrompts.map((prompt, index) => {
              const CategoryIcon = CATEGORIES[prompt.category]?.icon || CustomIcon;
              const categoryColor = CATEGORIES[prompt.category]?.color || '#71717a';
              
              return (
                <ListItem
                  key={prompt.id}
                  sx={{
                    borderBottom: index < sortedPrompts.length - 1 ? '1px solid #27272a' : 'none',
                    py: 2,
                    '&:hover': { background: 'rgba(255,255,255,0.02)' },
                  }}
                >
                  <Box sx={{ display: 'flex', alignItems: 'flex-start', gap: 2, flex: 1 }}>
                    <IconButton
                      size="small"
                      onClick={() => handleToggleFavorite(prompt)}
                      sx={{ color: prompt.isFavorite ? '#f59e0b' : '#3f3f46', mt: 0.5 }}
                    >
                      {prompt.isFavorite ? <StarIcon /> : <StarBorderIcon />}
                    </IconButton>
                    
                    <Box sx={{ flex: 1, minWidth: 0 }}>
                      <Stack direction="row" alignItems="center" spacing={1} mb={0.5}>
                        <CategoryIcon sx={{ fontSize: 18, color: categoryColor }} />
                        <Typography variant="subtitle1" sx={{ fontWeight: 600 }}>
                          {prompt.name}
                        </Typography>
                        {prompt.isDefault && (
                          <Chip label="Default" size="small" sx={{ height: 18, fontSize: '0.65rem', background: '#27272a' }} />
                        )}
                        {prompt.usageCount > 0 && (
                          <Chip 
                            label={`Used ${prompt.usageCount}x`} 
                            size="small" 
                            sx={{ height: 18, fontSize: '0.65rem', background: 'rgba(34, 197, 94, 0.2)', color: '#22c55e' }} 
                          />
                        )}
                      </Stack>
                      
                      {prompt.description && (
                        <Typography variant="body2" sx={{ color: 'text.secondary', mb: 1 }}>
                          {prompt.description}
                        </Typography>
                      )}
                      
                      <Typography 
                        variant="body2" 
                        sx={{ 
                          color: '#71717a', 
                          fontFamily: 'monospace',
                          fontSize: '0.8rem',
                          background: '#0f0f12',
                          p: 1,
                          borderRadius: 1,
                          maxHeight: 60,
                          overflow: 'hidden',
                          textOverflow: 'ellipsis',
                        }}
                      >
                        {prompt.template.substring(0, 150)}{prompt.template.length > 150 ? '...' : ''}
                      </Typography>
                    </Box>
                  </Box>
                  
                  <ListItemSecondaryAction>
                    <Stack direction="row" spacing={0.5}>
                      <Tooltip title="Edit">
                        <IconButton size="small" onClick={() => handleEditPrompt(prompt)}>
                          <EditIcon fontSize="small" sx={{ color: '#71717a' }} />
                        </IconButton>
                      </Tooltip>
                      {!prompt.isDefault && (
                        <Tooltip title="Delete">
                          <IconButton 
                            size="small" 
                            onClick={() => {
                              setPromptToDelete(prompt);
                              setDeleteConfirmOpen(true);
                            }}
                          >
                            <DeleteIcon fontSize="small" sx={{ color: '#71717a' }} />
                          </IconButton>
                        </Tooltip>
                      )}
                    </Stack>
                  </ListItemSecondaryAction>
                </ListItem>
              );
            })}
          </List>
        </Paper>
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
          <AIIcon sx={{ fontSize: 48, color: '#3f3f46', mb: 2 }} />
          <Typography variant="h6" sx={{ color: 'text.secondary', mb: 1 }}>
            {searchQuery || categoryFilter !== 'all' ? 'No matching prompts' : 'No Prompts Yet'}
          </Typography>
          <Typography variant="body2" sx={{ color: '#71717a', mb: 2 }}>
            {searchQuery || categoryFilter !== 'all' 
              ? 'Try adjusting your search or filter'
              : 'Create your first custom prompt to get started'}
          </Typography>
          {!searchQuery && categoryFilter === 'all' && (
            <Button variant="outlined" startIcon={<AddIcon />} onClick={handleNewPrompt}>
              Create Prompt
            </Button>
          )}
        </Paper>
      )}

      {/* Add/Edit Dialog */}
      <Dialog 
        open={dialogOpen} 
        onClose={() => setDialogOpen(false)}
        maxWidth="md"
        fullWidth
        PaperProps={{
          sx: { background: '#18181b', border: '1px solid #27272a' }
        }}
      >
        <DialogTitle sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
          <Typography variant="h6">
            {editingPrompt ? 'Edit Prompt' : 'New Prompt'}
          </Typography>
          <IconButton onClick={() => setDialogOpen(false)} size="small">
            <CloseIcon />
          </IconButton>
        </DialogTitle>
        
        <DialogContent>
          <Stack spacing={3} sx={{ pt: 1 }}>
            <Stack direction="row" spacing={2}>
              <TextField
                fullWidth
                label="Prompt Name"
                value={formData.name}
                onChange={(e) => handleFormChange('name', e.target.value)}
                placeholder="e.g., Bullet Point Summary"
              />
              
              <FormControl sx={{ minWidth: 150 }}>
                <InputLabel>Category</InputLabel>
                <Select
                  value={formData.category}
                  label="Category"
                  onChange={(e) => handleFormChange('category', e.target.value)}
                >
                  {Object.entries(CATEGORIES).map(([key, { label }]) => (
                    <MenuItem key={key} value={key}>{label}</MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Stack>

            <TextField
              fullWidth
              label="Description (optional)"
              value={formData.description}
              onChange={(e) => handleFormChange('description', e.target.value)}
              placeholder="Brief description of what this prompt does"
              inputProps={{ spellCheck: true }}
            />

            <Box>
              <Typography variant="subtitle2" sx={{ mb: 1, color: 'text.secondary' }}>
                Available Placeholders (click to insert):
              </Typography>
              <Stack direction="row" spacing={1} flexWrap="wrap" gap={1}>
                {PLACEHOLDERS.map(({ key, desc }) => (
                  <Tooltip key={key} title={desc}>
                    <Chip
                      label={key}
                      size="small"
                      onClick={() => insertPlaceholder(key)}
                      sx={{ 
                        cursor: 'pointer',
                        fontFamily: 'monospace',
                        '&:hover': { background: 'rgba(249, 115, 22, 0.2)' },
                      }}
                    />
                  </Tooltip>
                ))}
              </Stack>
            </Box>

            <TextField
              fullWidth
              label="Prompt Template"
              value={formData.template}
              onChange={(e) => handleFormChange('template', e.target.value)}
              placeholder="Enter your prompt template here. Use placeholders like {email_body} for dynamic content."
              multiline
              rows={6}
              inputProps={{ spellCheck: true }}
              sx={{
                '& .MuiInputBase-input': {
                  fontFamily: 'monospace',
                  fontSize: '0.9rem',
                },
              }}
            />
          </Stack>
        </DialogContent>

        <DialogActions sx={{ px: 3, pb: 3 }}>
          <Button onClick={() => setDialogOpen(false)} sx={{ color: 'text.secondary' }}>
            Cancel
          </Button>
          <Button variant="contained" onClick={handleSavePrompt}>
            {editingPrompt ? 'Save Changes' : 'Create Prompt'}
          </Button>
        </DialogActions>
      </Dialog>

      {/* Delete Confirmation */}
      <Dialog 
        open={deleteConfirmOpen} 
        onClose={() => setDeleteConfirmOpen(false)}
        PaperProps={{
          sx: { background: '#18181b', border: '1px solid #27272a' }
        }}
      >
        <DialogTitle>Delete Prompt?</DialogTitle>
        <DialogContent>
          <Typography>
            Are you sure you want to delete "{promptToDelete?.name}"? This cannot be undone.
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setDeleteConfirmOpen(false)}>Cancel</Button>
          <Button onClick={handleDeletePrompt} color="error" variant="contained">
            Delete
          </Button>
        </DialogActions>
      </Dialog>
    </Stack>
  );
}

export default PromptsTab;

