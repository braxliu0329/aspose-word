import { useState, useEffect, useCallback, ChangeEvent, useMemo, useRef } from 'react';
import debounce from 'lodash.debounce';
import {
  fetchInitialHtml,
  updateDocumentStyle,
  updateRangeStyle,
  uploadDocument,
  downloadDocument,
  redoDocument,
  undoDocument,
  type DocumentResponse,
} from './api';
import Preview from './components/Preview';
import Controls from './components/Controls';
import './App.css';

interface SelectionData {
  text: string;
  startNodeId: string;
  startOffset: number;
  endNodeId: string;
  endOffset: number;
}

function findBookmarkAnchor(node: Node | null): HTMLAnchorElement | null {
  let el: HTMLElement | null = null;
  if (!node) return null;
  if (node.nodeType === Node.TEXT_NODE) el = node.parentElement;
  else el = node as HTMLElement;

  while (el) {
    if (el.tagName === 'A') {
      const name = el.getAttribute('name') || '';
      if (name.startsWith('Run_')) return el as HTMLAnchorElement;
    }
    el = el.parentElement;
  }

  if (node.nodeType === Node.TEXT_NODE && node.parentElement) {
    let prev: ChildNode | null = node as ChildNode;
    while (prev && prev.parentNode === node.parentNode) {
      prev = prev.previousSibling;
      if (!prev) break;
      if ((prev as HTMLElement).nodeType === Node.ELEMENT_NODE) {
        const tag = (prev as HTMLElement).tagName;
        if (tag === 'A') {
          const name = (prev as HTMLElement).getAttribute('name') || '';
          if (name.startsWith('Run_')) return prev as HTMLAnchorElement;
        }
      }
    }
  }

  return null;
}

function offsetWithinAnchor(anchor: HTMLAnchorElement, rangeContainer: Node, rangeOffset: number): number {
  const r = document.createRange();
  r.selectNodeContents(anchor);
  r.setEnd(rangeContainer, rangeOffset);
  return r.toString().length;
}

function App() {
  const [htmlContent, setHtmlContent] = useState<string>('');
  const [fontName, setFontName] = useState<string>('Arial');
  const [fontSize, setFontSize] = useState<number>(12);
  const [fontColor, setFontColor] = useState<string>('#000000');
  const [loading, setLoading] = useState<boolean>(false);
  const [downloading, setDownloading] = useState<boolean>(false);
  const [currentSelection, setCurrentSelection] = useState<SelectionData | null>(null);
  const [history, setHistory] = useState<{ canUndo: boolean; canRedo: boolean }>({ canUndo: false, canRedo: false });

  // Keep track of current style state to debounce correctly
  const styleStateRef = useRef({ fontName, fontSize, fontColor });
  styleStateRef.current = { fontName, fontSize, fontColor };

  // Initial load
  useEffect(() => {
    const loadInitial = async () => {
      try {
        setLoading(true);
        const res = await fetchInitialHtml();
        setHtmlContent(res.html);
        setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
      } catch (error) {
        console.error("Failed to load initial document", error);
      } finally {
        setLoading(false);
      }
    };
    loadInitial();
  }, []);

  // Handle Text Selection in Preview
  useEffect(() => {
    const handleSelectionChange = () => {
      const selection = window.getSelection();
      if (!selection || selection.rangeCount === 0 || selection.isCollapsed) {
        return;
      }

      const range = selection.getRangeAt(0);
      const startAnchor = findBookmarkAnchor(range.startContainer);
      const endAnchor = findBookmarkAnchor(range.endContainer);
      if (!startAnchor || !endAnchor) return;

      const text = selection.toString();
      if (!text) return;

      const startNodeId = startAnchor.getAttribute('name') || '';
      const endNodeId = endAnchor.getAttribute('name') || '';
      if (!startNodeId || !endNodeId) return;

      const startOffset = offsetWithinAnchor(startAnchor, range.startContainer, range.startOffset);
      const endOffset = offsetWithinAnchor(endAnchor, range.endContainer, range.endOffset);

      setCurrentSelection({ text, startNodeId, startOffset, endNodeId, endOffset });
        
        // Try to sync UI with computed style
        const element = selection.anchorNode?.parentElement;
        if (element) {
             const computed = window.getComputedStyle(element);
             const family = computed.fontFamily.split(',')[0].replace(/['"]/g, '');
             const size = parseFloat(computed.fontSize);
             const sizePt = Math.round(size * 0.75);
             
             setFontName(family);
             setFontSize(sizePt);
        }
    };

    document.addEventListener('mouseup', handleSelectionChange);
    return () => document.removeEventListener('mouseup', handleSelectionChange);
  }, []);

  // Debounced update function
  const applyDocResponse = useCallback((res: DocumentResponse) => {
    setHtmlContent(res.html);
    setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
  }, []);

  const debouncedUpdate = useMemo(() =>
    debounce(async (sel: SelectionData | null, name: string, size: number, color: string) => {
      try {
        setLoading(true);
        let res: DocumentResponse;
        if (sel && sel.startNodeId && sel.endNodeId) {
          res = await updateRangeStyle(sel.startNodeId, sel.startOffset, sel.endNodeId, sel.endOffset, name, size, color);
          setCurrentSelection(null);
        } else {
          res = await updateDocumentStyle(name, size, color);
        }
        applyDocResponse(res);
      } catch (error) {
        console.error("Failed to update document", error);
      } finally {
        setLoading(false);
      }
    }, 500)
  , [applyDocResponse]);

  useEffect(() => {
    return () => {
      debouncedUpdate.cancel();
    };
  }, [debouncedUpdate]);

  const handleStyleChange = (name: string, size: number, color: string) => {
    setFontName(name);
    setFontSize(size);
    setFontColor(color);
    debouncedUpdate(currentSelection, name, size, color);
  };

  const handleUndo = useCallback(async () => {
    if (!history.canUndo || loading) return;
    debouncedUpdate.cancel();
    try {
      setLoading(true);
      const res = await undoDocument();
      applyDocResponse(res);
      setCurrentSelection(null);
    } catch (error) {
      console.error('Failed to undo', error);
    } finally {
      setLoading(false);
    }
  }, [history.canUndo, loading, debouncedUpdate, applyDocResponse]);

  const handleRedo = useCallback(async () => {
    if (!history.canRedo || loading) return;
    debouncedUpdate.cancel();
    try {
      setLoading(true);
      const res = await redoDocument();
      applyDocResponse(res);
      setCurrentSelection(null);
    } catch (error) {
      console.error('Failed to redo', error);
    } finally {
      setLoading(false);
    }
  }, [history.canRedo, loading, debouncedUpdate, applyDocResponse]);

  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      const target = e.target as HTMLElement | null;
      const tag = target?.tagName?.toLowerCase();
      const isFormField = tag === 'input' || tag === 'textarea' || tag === 'select' || target?.isContentEditable;
      if (isFormField) return;

      const isMod = e.ctrlKey || e.metaKey;
      if (!isMod) return;

      const key = e.key.toLowerCase();
      if (key === 'z') {
        e.preventDefault();
        if (e.shiftKey) handleRedo();
        else handleUndo();
      } else if (key === 'y') {
        e.preventDefault();
        handleRedo();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [handleRedo, handleUndo]);

  const handleFileUpload = async (e: ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    try {
      setLoading(true);
      debouncedUpdate.cancel();
      const res = await uploadDocument(file);
      applyDocResponse(res);
      setCurrentSelection(null);
    } catch (error) {
      console.error("Failed to upload document", error);
      alert("Upload failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  const handleDownload = async () => {
    try {
      setDownloading(true);
      await downloadDocument();
    } catch (error) {
      console.error("Failed to download document", error);
      alert("Download failed. Please try again.");
    } finally {
      setDownloading(false);
    }
  };

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>Aspose.Words Web Prototype</h1>
      </header>
      <main className="app-main">
        <div className="preview-section">
          <Controls
            fontName={fontName}
            fontSize={fontSize}
            fontColor={fontColor}
            canUndo={history.canUndo}
            canRedo={history.canRedo}
            loading={loading}
            selectionLabel={
              currentSelection
                ? `范围编辑：${currentSelection.startNodeId} → ${currentSelection.endNodeId}（${currentSelection.startOffset} → ${currentSelection.endOffset}）`
                : '全局样式编辑'
            }
            onChange={handleStyleChange}
            onUndo={handleUndo}
            onRedo={handleRedo}
            onClearSelection={currentSelection ? () => setCurrentSelection(null) : undefined}
          />
          <div className="preview-scroll">
            <Preview html={htmlContent} loading={loading} />
          </div>
        </div>
        <div className="controls-section">
          <div className="action-box" style={{ marginBottom: '20px', padding: '15px', border: '1px dashed #ccc', borderRadius: '8px' }}>
            <h3>Document Actions</h3>
            
            <div style={{ marginBottom: '15px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>Upload:</label>
              <input 
                type="file" 
                accept=".docx,.doc" 
                onChange={handleFileUpload}
                style={{ width: '100%' }}
              />
              <p style={{ fontSize: '0.8rem', color: '#666', marginTop: '5px' }}>
                Supports .docx, .doc
              </p>
            </div>

            <div style={{ borderTop: '1px solid #eee', paddingTop: '15px' }}>
              <button 
                onClick={handleDownload}
                disabled={downloading}
                style={{
                  width: '100%',
                  padding: '10px',
                  backgroundColor: '#4CAF50',
                  color: 'white',
                  border: 'none',
                  borderRadius: '4px',
                  cursor: downloading ? 'wait' : 'pointer',
                  fontSize: '16px'
                }}
              >
                {downloading ? 'Downloading...' : 'Download Document'}
              </button>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;
