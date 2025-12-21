import React, { useState, useEffect, useCallback } from 'react';
import {
  Box,
  Paper,
  Typography,
  Chip,
  Stack,
  IconButton,
  Tooltip,
  Checkbox,
  Button,
  Menu,
  MenuItem,
  ListItemIcon,
  ListItemText,
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  TextField,
  Divider,
  Badge,
  FormControlLabel,
  Switch,
} from '@mui/material';
import {
  Lock as LockIcon,
  DragIndicator as DragIcon,
  Add as AddIcon,
  Save as SaveIcon,
  BookmarkBorder as PresetIcon,
  Bookmark as PresetFilledIcon,
  Delete as DeleteIcon,
  ExpandMore as ExpandMoreIcon,
  ExpandLess as ExpandLessIcon,
  Check as CheckIcon,
  Clear as ClearIcon,
  Summarize as SummarizeIcon,
  Reply as ReplyIcon,
  Lightbulb as InsightIcon,
  Code as CustomIcon,
  EditNote as QuickNoteIcon,
} from '@mui/icons-material';

// Category icons
const CATEGORY_ICONS = {
  summarize: SummarizeIcon,
  reply: ReplyIcon,
  insights: InsightIcon,
  custom: CustomIcon,
};

const CATEGORY_COLORS = {
  summarize: '#f97316',
  reply: '#3b82f6',
  insights: '#22c55e',
  custom: '#a855f7',
};

function PromptSelector({
  prompts,
  selectedPromptIds,
  setSelectedPromptIds,
  hasPersonalContext,
  personalContextEnabled = true,
  companyContextEnabled = false,
  presets,
  onSavePreset,
  onLoadPreset,
  onDeletePreset,
  showMessage,
  quickNotes = '',
  setQuickNotes,
  keepQuickNotes = false,
  setKeepQuickNotes,
}) {
  const [expanded, setExpanded] = useState(true);
  const [presetMenuAnchor, setPresetMenuAnchor] = useState(null);
  const [savePresetOpen, setSavePresetOpen] = useState(false);
  const [newPresetName, setNewPresetName] = useState('');
  const [draggedIndex, setDraggedIndex] = useState(null);

  // Get selected prompts in order
  const selectedPrompts = selectedPromptIds
    .map(id => prompts.find(p => p.id === id))
    .filter(Boolean);

  // Toggle prompt selection
  const handleTogglePrompt = (promptId) => {
    setSelectedPromptIds(prev => {
      if (prev.includes(promptId)) {
        return prev.filter(id => id !== promptId);
      } else {
        return [...prev, promptId];
      }
    });
  };

  // Clear all selections
  const handleClearAll = () => {
    setSelectedPromptIds([]);
  };

  // Drag and drop handlers
  const handleDragStart = (e, index) => {
    setDraggedIndex(index);
    e.dataTransfer.effectAllowed = 'move';
  };

  const handleDragOver = (e, index) => {
    e.preventDefault();
    if (draggedIndex === null || draggedIndex === index) return;

    const newOrder = [...selectedPromptIds];
    const [dragged] = newOrder.splice(draggedIndex, 1);
    newOrder.splice(index, 0, dragged);
    setSelectedPromptIds(newOrder);
    setDraggedIndex(index);
  };

  const handleDragEnd = () => {
    setDraggedIndex(null);
  };

  // Save preset
  const handleSavePreset = () => {
    if (!newPresetName.trim()) {
      showMessage('warning', 'Please enter a preset name');
      return;
    }
    if (selectedPromptIds.length === 0) {
      showMessage('warning', 'Select at least one prompt to save as preset');
      return;
    }
    onSavePreset(newPresetName.trim(), selectedPromptIds);
    setNewPresetName('');
    setSavePresetOpen(false);
  };

  return (
    <Paper
      sx={{
        background: '#1a1a1f',
        border: '1px solid #27272a',
        borderRadius: 2,
        overflow: 'hidden',
      }}
    >
      {/* Header */}
      <Box
        sx={{
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          p: 1.5,
          borderBottom: expanded ? '1px solid #27272a' : 'none',
          cursor: 'pointer',
          '&:hover': { background: 'rgba(255,255,255,0.02)' },
        }}
        onClick={() => setExpanded(!expanded)}
      >
        <Stack direction="row" alignItems="center" spacing={1}>
          <Typography variant="subtitle2" sx={{ fontWeight: 600 }}>
            📋 Active Prompts
          </Typography>
          <Badge 
            badgeContent={selectedPromptIds.length} 
            color="primary"
            sx={{ '& .MuiBadge-badge': { fontSize: '0.7rem' } }}
          >
            <Box />
          </Badge>
        </Stack>
        <IconButton size="small" sx={{ color: 'text.secondary' }}>
          {expanded ? <ExpandLessIcon /> : <ExpandMoreIcon />}
        </IconButton>
      </Box>

      {expanded && (
        <Box sx={{ p: 1.5 }}>
          {/* Context Indicators */}
          <Stack spacing={0.5} sx={{ mb: 1 }}>
            {/* Personal Context Indicator */}
            <Box
              sx={{
                display: 'flex',
                alignItems: 'center',
                gap: 0.75,
                px: 1,
                py: 0.5,
                background: !personalContextEnabled 
                  ? 'rgba(239, 68, 68, 0.08)' 
                  : hasPersonalContext 
                    ? 'rgba(34, 197, 94, 0.08)' 
                    : 'rgba(113, 113, 122, 0.08)',
                borderRadius: 0.75,
                border: `1px solid ${!personalContextEnabled 
                  ? 'rgba(239, 68, 68, 0.25)' 
                  : hasPersonalContext 
                    ? 'rgba(34, 197, 94, 0.25)' 
                    : '#3f3f46'}`,
              }}
            >
              <LockIcon sx={{ fontSize: 12, color: !personalContextEnabled ? '#ef4444' : hasPersonalContext ? '#22c55e' : '#71717a' }} />
              <Typography sx={{ fontSize: '0.7rem', color: !personalContextEnabled ? '#ef4444' : hasPersonalContext ? '#22c55e' : '#71717a', flex: 1 }}>
                Personal Context {!personalContextEnabled ? '(OFF)' : hasPersonalContext ? '(ON)' : '(not set)'}
              </Typography>
              {personalContextEnabled && hasPersonalContext && <CheckIcon sx={{ fontSize: 10, color: '#22c55e' }} />}
            </Box>

            {/* Company Context Indicator */}
            <Box
              sx={{
                display: 'flex',
                alignItems: 'center',
                gap: 0.75,
                px: 1,
                py: 0.5,
                background: companyContextEnabled 
                  ? 'rgba(59, 130, 246, 0.08)' 
                  : 'rgba(239, 68, 68, 0.08)',
                borderRadius: 0.75,
                border: `1px solid ${companyContextEnabled 
                  ? 'rgba(59, 130, 246, 0.25)' 
                  : 'rgba(239, 68, 68, 0.25)'}`,
              }}
            >
              <LockIcon sx={{ fontSize: 12, color: companyContextEnabled ? '#3b82f6' : '#ef4444' }} />
              <Typography sx={{ fontSize: '0.7rem', color: companyContextEnabled ? '#3b82f6' : '#ef4444', flex: 1 }}>
                Company Context {companyContextEnabled ? '(ON)' : '(OFF)'}
              </Typography>
              {companyContextEnabled && <CheckIcon sx={{ fontSize: 10, color: '#3b82f6' }} />}
            </Box>
          </Stack>

          {/* Selected Prompts (Draggable) */}
          {selectedPrompts.length > 0 ? (
            <Box sx={{ mb: 1.5 }}>
              <Typography variant="caption" sx={{ color: '#71717a', display: 'block', mb: 0.5 }}>
                Selected (drag to reorder):
              </Typography>
              <Stack spacing={0.5}>
                {selectedPrompts.map((prompt, index) => {
                  const Icon = CATEGORY_ICONS[prompt.category] || CustomIcon;
                  const color = CATEGORY_COLORS[prompt.category] || '#71717a';
                  
                  return (
                    <Box
                      key={prompt.id}
                      draggable
                      onDragStart={(e) => handleDragStart(e, index)}
                      onDragOver={(e) => handleDragOver(e, index)}
                      onDragEnd={handleDragEnd}
                      sx={{
                        display: 'flex',
                        alignItems: 'center',
                        gap: 1,
                        p: 0.75,
                        background: draggedIndex === index ? 'rgba(249, 115, 22, 0.2)' : '#0f0f12',
                        borderRadius: 1,
                        border: '1px solid #27272a',
                        cursor: 'grab',
                        '&:hover': { borderColor: '#3f3f46' },
                        '&:active': { cursor: 'grabbing' },
                      }}
                    >
                      <DragIcon sx={{ fontSize: 16, color: '#3f3f46' }} />
                      <Typography 
                        variant="caption" 
                        sx={{ 
                          color: '#a1a1aa',
                          background: 'rgba(249, 115, 22, 0.2)',
                          px: 0.5,
                          borderRadius: 0.5,
                          fontWeight: 600,
                        }}
                      >
                        {index + 1}
                      </Typography>
                      <Icon sx={{ fontSize: 14, color }} />
                      <Typography variant="body2" sx={{ flex: 1, fontSize: '0.8rem' }}>
                        {prompt.name}
                      </Typography>
                      <IconButton 
                        size="small" 
                        onClick={(e) => { e.stopPropagation(); handleTogglePrompt(prompt.id); }}
                        sx={{ p: 0.25 }}
                      >
                        <ClearIcon sx={{ fontSize: 14, color: '#71717a' }} />
                      </IconButton>
                    </Box>
                  );
                })}
              </Stack>
            </Box>
          ) : (
            <Box
              sx={{
                p: 2,
                mb: 1.5,
                textAlign: 'center',
                background: '#0f0f12',
                borderRadius: 1,
                border: '1px dashed #3f3f46',
              }}
            >
              <Typography variant="caption" sx={{ color: '#71717a' }}>
                No prompts selected — raw email only
              </Typography>
            </Box>
          )}

          {/* Available Prompts */}
          <Typography variant="caption" sx={{ color: '#71717a', display: 'block', mb: 0.5 }}>
            Available prompts:
          </Typography>
          <Box sx={{ maxHeight: 200, overflow: 'auto', mb: 1.5 }}>
            <Stack spacing={0.5}>
              {prompts.map((prompt) => {
                const isSelected = selectedPromptIds.includes(prompt.id);
                const Icon = CATEGORY_ICONS[prompt.category] || CustomIcon;
                const color = CATEGORY_COLORS[prompt.category] || '#71717a';
                
                return (
                  <Box
                    key={prompt.id}
                    onClick={() => handleTogglePrompt(prompt.id)}
                    sx={{
                      display: 'flex',
                      alignItems: 'center',
                      gap: 1,
                      p: 0.75,
                      background: isSelected ? 'rgba(249, 115, 22, 0.1)' : 'transparent',
                      borderRadius: 1,
                      cursor: 'pointer',
                      '&:hover': { background: isSelected ? 'rgba(249, 115, 22, 0.15)' : 'rgba(255,255,255,0.02)' },
                    }}
                  >
                    <Checkbox
                      checked={isSelected}
                      size="small"
                      sx={{ 
                        p: 0, 
                        color: '#3f3f46',
                        '&.Mui-checked': { color: 'primary.main' },
                      }}
                    />
                    <Icon sx={{ fontSize: 14, color }} />
                    <Typography variant="body2" sx={{ flex: 1, fontSize: '0.8rem' }}>
                      {prompt.name}
                    </Typography>
                    {prompt.isFavorite && (
                      <Typography sx={{ fontSize: '0.7rem' }}>⭐</Typography>
                    )}
                  </Box>
                );
              })}
            </Stack>
          </Box>

          {/* Quick Notes - One-time Disposable Prompt */}
          <Box sx={{ mt: 1.5, mb: 1 }}>
            <Stack direction="row" alignItems="center" justifyContent="space-between" sx={{ mb: 0.5 }}>
              <Stack direction="row" alignItems="center" spacing={0.5}>
                <QuickNoteIcon sx={{ fontSize: 14, color: quickNotes.trim() ? '#22c55e' : '#71717a' }} />
                <Typography variant="caption" sx={{ color: quickNotes.trim() ? '#22c55e' : '#71717a', fontWeight: 500 }}>
                  Quick Notes {quickNotes.trim() ? '✓' : '(optional)'}
                </Typography>
              </Stack>
              <FormControlLabel
                control={
                  <Switch
                    checked={keepQuickNotes}
                    onChange={(e) => setKeepQuickNotes?.(e.target.checked)}
                    size="small"
                    sx={{ transform: 'scale(0.75)' }}
                  />
                }
                label={<Typography variant="caption" sx={{ color: '#71717a', fontSize: '0.65rem' }}>Keep</Typography>}
                labelPlacement="start"
                sx={{ m: 0, ml: 'auto' }}
              />
            </Stack>
            <TextField
              fullWidth
              multiline
              rows={2}
              placeholder="Add one-time notes... e.g., 'mention the deadline is Friday' or 'keep it brief'"
              value={quickNotes}
              onChange={(e) => setQuickNotes?.(e.target.value)}
              variant="outlined"
              size="small"
              sx={{
                '& .MuiOutlinedInput-root': {
                  fontSize: '0.8rem',
                  background: '#0f0f12',
                  '& fieldset': { borderColor: quickNotes.trim() ? 'rgba(34, 197, 94, 0.4)' : '#27272a' },
                  '&:hover fieldset': { borderColor: quickNotes.trim() ? 'rgba(34, 197, 94, 0.6)' : '#3f3f46' },
                  '&.Mui-focused fieldset': { borderColor: '#f97316' },
                },
              }}
              inputProps={{ maxLength: 500, spellCheck: true }}
            />
            <Typography variant="caption" sx={{ color: '#52525b', display: 'block', textAlign: 'right', mt: 0.25, fontSize: '0.65rem' }}>
              {quickNotes.length}/500 • Clears after use {keepQuickNotes && '(keep enabled)'}
            </Typography>
          </Box>

          <Divider sx={{ my: 1, borderColor: '#27272a' }} />

          {/* Actions */}
          <Stack direction="row" spacing={1} justifyContent="space-between">
            <Stack direction="row" spacing={0.5}>
              <Tooltip title="Clear all">
                <IconButton 
                  size="small" 
                  onClick={handleClearAll}
                  disabled={selectedPromptIds.length === 0}
                  sx={{ color: '#71717a' }}
                >
                  <ClearIcon fontSize="small" />
                </IconButton>
              </Tooltip>
              <Tooltip title="Save as preset">
                <IconButton 
                  size="small" 
                  onClick={() => setSavePresetOpen(true)}
                  disabled={selectedPromptIds.length === 0}
                  sx={{ color: '#71717a' }}
                >
                  <SaveIcon fontSize="small" />
                </IconButton>
              </Tooltip>
            </Stack>

            {/* Presets Menu */}
            <Button
              size="small"
              startIcon={<PresetIcon />}
              onClick={(e) => setPresetMenuAnchor(e.currentTarget)}
              sx={{ color: '#a1a1aa', fontSize: '0.75rem' }}
            >
              Presets {presets.length > 0 && `(${presets.length})`}
            </Button>
            <Menu
              anchorEl={presetMenuAnchor}
              open={Boolean(presetMenuAnchor)}
              onClose={() => setPresetMenuAnchor(null)}
              PaperProps={{
                sx: { background: '#18181b', border: '1px solid #27272a', minWidth: 180 }
              }}
            >
              {presets.length === 0 ? (
                <MenuItem disabled>
                  <Typography variant="caption" sx={{ color: '#71717a' }}>
                    No presets saved
                  </Typography>
                </MenuItem>
              ) : (
                presets.map((preset) => (
                  <MenuItem 
                    key={preset.id}
                    sx={{ display: 'flex', justifyContent: 'space-between' }}
                  >
                    <Box onClick={() => { onLoadPreset(preset); setPresetMenuAnchor(null); }} sx={{ flex: 1 }}>
                      <Typography variant="body2">{preset.name}</Typography>
                      <Typography variant="caption" sx={{ color: '#71717a' }}>
                        {preset.promptIds.length} prompts
                      </Typography>
                    </Box>
                    <IconButton
                      size="small"
                      onClick={(e) => { e.stopPropagation(); onDeletePreset(preset.id); }}
                      sx={{ color: '#71717a', ml: 1 }}
                    >
                      <DeleteIcon fontSize="small" />
                    </IconButton>
                  </MenuItem>
                ))
              )}
            </Menu>
          </Stack>
        </Box>
      )}

      {/* Save Preset Dialog */}
      <Dialog 
        open={savePresetOpen} 
        onClose={() => setSavePresetOpen(false)}
        PaperProps={{ sx: { background: '#18181b', border: '1px solid #27272a' } }}
      >
        <DialogTitle>Save Prompt Preset</DialogTitle>
        <DialogContent>
          <TextField
            autoFocus
            fullWidth
            label="Preset Name"
            value={newPresetName}
            onChange={(e) => setNewPresetName(e.target.value)}
            placeholder="e.g., Client Email Response"
            sx={{ mt: 1 }}
          />
          <Typography variant="caption" sx={{ color: '#71717a', mt: 1, display: 'block' }}>
            Saving {selectedPromptIds.length} prompt(s)
          </Typography>
        </DialogContent>
        <DialogActions>
          <Button onClick={() => setSavePresetOpen(false)}>Cancel</Button>
          <Button variant="contained" onClick={handleSavePreset}>Save Preset</Button>
        </DialogActions>
      </Dialog>
    </Paper>
  );
}

export default PromptSelector;

