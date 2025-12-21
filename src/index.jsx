import React from 'react';
import ReactDOM from 'react-dom/client';
import { ThemeProvider, createTheme, CssBaseline } from '@mui/material';
import App from './App';
import './index.css';

// Error Boundary to catch rendering errors
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error) {
    return { hasError: true, error };
  }

  componentDidCatch(error, errorInfo) {
    console.error('React Error:', error, errorInfo);
    
    // Report to crash reporting system (if available)
    if (window.electronAPI?.reportRendererError) {
      window.electronAPI.reportRendererError(
        error?.toString() || 'Unknown error',
        errorInfo?.componentStack || 'No component stack',
        window.location.href
      );
    }
  }

  render() {
    if (this.state.hasError) {
      return (
        <div style={{ 
          padding: 40, 
          backgroundColor: '#0a0a0f', 
          color: '#fafafa',
          minHeight: '100vh',
          fontFamily: 'sans-serif'
        }}>
          <h1 style={{ color: '#f97316' }}>Something went wrong</h1>
          <p>The app encountered an error. Please try refreshing (Ctrl+R).</p>
          <pre style={{ 
            background: '#18181b', 
            padding: 20, 
            borderRadius: 8,
            overflow: 'auto',
            fontSize: 12
          }}>
            {this.state.error?.toString()}
          </pre>
          <button 
            onClick={() => window.location.reload()}
            style={{
              marginTop: 20,
              padding: '10px 20px',
              background: '#f97316',
              border: 'none',
              borderRadius: 8,
              color: '#000',
              fontWeight: 'bold',
              cursor: 'pointer'
            }}
          >
            Reload App
          </button>
        </div>
      );
    }
    return this.props.children;
  }
}

// Custom dark theme with orange accent (Grok-inspired)
const darkTheme = createTheme({
  palette: {
    mode: 'dark',
    primary: {
      main: '#f97316', // Vibrant orange
      light: '#fb923c',
      dark: '#ea580c',
      contrastText: '#000000',
    },
    secondary: {
      main: '#06b6d4', // Cyan accent
      light: '#22d3ee',
      dark: '#0891b2',
    },
    background: {
      default: '#0a0a0f',
      paper: '#18181b',
    },
    text: {
      primary: '#fafafa',
      secondary: '#a1a1aa',
    },
    divider: '#27272a',
    error: {
      main: '#ef4444',
    },
    warning: {
      main: '#f59e0b',
    },
    success: {
      main: '#22c55e',
    },
    info: {
      main: '#3b82f6',
    },
  },
  typography: {
    fontFamily: '"Plus Jakarta Sans", -apple-system, BlinkMacSystemFont, sans-serif',
    h1: {
      fontWeight: 700,
      letterSpacing: '-0.02em',
    },
    h2: {
      fontWeight: 700,
      letterSpacing: '-0.01em',
    },
    h3: {
      fontWeight: 600,
    },
    h4: {
      fontWeight: 600,
    },
    h5: {
      fontWeight: 600,
    },
    h6: {
      fontWeight: 600,
    },
    button: {
      fontWeight: 600,
      textTransform: 'none',
    },
    code: {
      fontFamily: '"JetBrains Mono", monospace',
    },
  },
  shape: {
    borderRadius: 12,
  },
  components: {
    MuiButton: {
      styleOverrides: {
        root: {
          borderRadius: 8,
          padding: '10px 20px',
          boxShadow: 'none',
          '&:hover': {
            boxShadow: '0 4px 12px rgba(249, 115, 22, 0.25)',
          },
        },
        contained: {
          background: 'linear-gradient(135deg, #f97316 0%, #ea580c 100%)',
          '&:hover': {
            background: 'linear-gradient(135deg, #fb923c 0%, #f97316 100%)',
          },
        },
      },
    },
    MuiPaper: {
      styleOverrides: {
        root: {
          backgroundImage: 'none',
          border: '1px solid #27272a',
        },
      },
    },
    MuiCard: {
      styleOverrides: {
        root: {
          backgroundImage: 'none',
          border: '1px solid #27272a',
          transition: 'border-color 0.2s ease, box-shadow 0.2s ease',
          '&:hover': {
            borderColor: '#3f3f46',
            boxShadow: '0 8px 32px rgba(0, 0, 0, 0.4)',
          },
        },
      },
    },
    MuiTab: {
      styleOverrides: {
        root: {
          textTransform: 'none',
          fontWeight: 500,
          fontSize: '0.95rem',
        },
      },
    },
    MuiChip: {
      styleOverrides: {
        root: {
          fontWeight: 500,
        },
      },
    },
  },
});

const root = ReactDOM.createRoot(document.getElementById('root'));
root.render(
  <React.StrictMode>
    <ErrorBoundary>
      <ThemeProvider theme={darkTheme}>
        <CssBaseline />
        <App />
      </ThemeProvider>
    </ErrorBoundary>
  </React.StrictMode>
);

// Global error handler for uncaught errors in renderer
window.onerror = (message, source, lineno, colno, error) => {
  console.error('Global error:', message, source, lineno, colno);
  if (window.electronAPI?.reportRendererError) {
    window.electronAPI.reportRendererError(
      `${message} at ${source}:${lineno}:${colno}`,
      error?.stack || 'No stack trace',
      window.location.href
    );
  }
};

// Global handler for unhandled promise rejections
window.onunhandledrejection = (event) => {
  console.error('Unhandled rejection:', event.reason);
  if (window.electronAPI?.reportRendererError) {
    window.electronAPI.reportRendererError(
      event.reason?.message || String(event.reason),
      event.reason?.stack || 'No stack trace',
      window.location.href
    );
  }
};

