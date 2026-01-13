import { useState, useEffect, useCallback, ChangeEvent, useMemo, useRef } from 'react';
import debounce from 'lodash.debounce';
import {
  fetchInitialHtml,
  fetchPageHtml,
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

function getActiveRangeInElement(root: HTMLElement): Range | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0 || sel.isCollapsed) return null;
  const range = sel.getRangeAt(0);
  if (!root.contains(range.commonAncestorContainer)) return null;
  return range;
}

function selectionToSelectionData(range: Range): SelectionData | null {
  const selection = window.getSelection();
  const text = selection?.toString() ?? '';
  if (!text) return null;

  const startAnchor = findBookmarkAnchor(range.startContainer);
  const endAnchor = findBookmarkAnchor(range.endContainer);
  if (!startAnchor || !endAnchor) return null;

  const startNodeId = startAnchor.getAttribute('name') || '';
  const endNodeId = endAnchor.getAttribute('name') || '';
  if (!startNodeId || !endNodeId) return null;

  const startOffset = offsetWithinAnchor(startAnchor, range.startContainer, range.startOffset);
  const endOffset = offsetWithinAnchor(endAnchor, range.endContainer, range.endOffset);

  return { text, startNodeId, startOffset, endNodeId, endOffset };
}

function findBookmarkAnchor(node: Node | null): HTMLAnchorElement | null {
  if (!node) return null;

  const getRunAnchor = (el: Element | null): HTMLAnchorElement | null => {
    if (!el) return null;
    if (el.tagName === 'A') {
      const name = el.getAttribute('name') || '';
      if (name.startsWith('Run_')) return el as HTMLAnchorElement;
    }
    const anchors = el.querySelectorAll('a[name^="Run_"]');
    if (anchors.length > 0) return anchors[anchors.length - 1] as HTMLAnchorElement;
    return null;
  };

  let el: HTMLElement | null = node.nodeType === Node.TEXT_NODE ? node.parentElement : (node as HTMLElement);
  while (el) {
    const direct = getRunAnchor(el);
    if (direct) return direct;

    let prev: Element | null = el.previousElementSibling;
    while (prev) {
      const found = getRunAnchor(prev);
      if (found) return found;
      prev = prev.previousElementSibling;
    }

    el = el.parentElement;
  }

  return null;
}

function offsetWithinAnchor(anchor: HTMLAnchorElement, rangeContainer: Node, rangeOffset: number): number {
  const r = document.createRange();
  r.selectNodeContents(anchor);
  r.setEnd(rangeContainer, rangeOffset);
  return r.toString().length;
}

function getDomPointAtTextOffset(root: Node, textOffset: number): { node: Text; offset: number } | null {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let remaining = Math.max(0, textOffset);
  let lastText: Text | null = null;

  while (walker.nextNode()) {
    const textNode = walker.currentNode as Text;
    lastText = textNode;
    const len = textNode.textContent?.length ?? 0;
    if (remaining <= len) return { node: textNode, offset: remaining };
    remaining -= len;
  }

  if (!lastText) return null;
  const end = lastText.textContent?.length ?? 0;
  return { node: lastText, offset: end };
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
  const [pageIndex, setPageIndex] = useState<number>(1);
  const [pageCount, setPageCount] = useState<number>(1);
  const [pageJumpInput, setPageJumpInput] = useState<string>('1');

  const documentRenderRef = useRef<HTMLDivElement | null>(null);
  const pendingRangeSelectionRef = useRef<SelectionData | null>(null);

  // Keep track of current style state to debounce correctly
  const styleStateRef = useRef({ fontName, fontSize, fontColor });
  styleStateRef.current = { fontName, fontSize, fontColor };

  // Initial load
  useEffect(() => {
    const loadInitial = async () => {
      try {
        setLoading(true);
        const res = await fetchInitialHtml(1);
        setHtmlContent(res.html);
        setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
        const pc = res.pageCount ?? 1;
        const pi = res.pageIndex ?? 1;
        setPageCount(pc);
        setPageIndex(pi);
        setPageJumpInput(String(pi));
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
    const handleMouseUp = (e: MouseEvent) => {
      const docEl = documentRenderRef.current;
      if (!docEl) return;

      const targetNode = e.target as Node | null;
      const isInDoc = !!targetNode && docEl.contains(targetNode);
      const range = getActiveRangeInElement(docEl);
      const isRangeInDoc = !!range;
      if (!isInDoc && !isRangeInDoc) return;

      if (!range) {
        if (isInDoc) setCurrentSelection(null);
        return;
      }

      const selData = selectionToSelectionData(range);
      if (!selData) return;

      setCurrentSelection(selData);
        
        // Try to sync UI with computed style
        const element = window.getSelection()?.anchorNode?.parentElement;
        if (element) {
             const computed = window.getComputedStyle(element);
             const family = computed.fontFamily.split(',')[0].replace(/['"]/g, '');
             const size = parseFloat(computed.fontSize);
             const sizePt = Math.round(size * 0.75);
             
             setFontName(family);
             setFontSize(sizePt);
        }
    };

    document.addEventListener('mouseup', handleMouseUp);
    return () => document.removeEventListener('mouseup', handleMouseUp);
  }, []);

  useEffect(() => {
    const handleKeyUp = (e: KeyboardEvent) => {
      const docEl = documentRenderRef.current;
      if (!docEl) return;
      if (pendingRangeSelectionRef.current) return;
      if (loading) return;

      const target = e.target as Node | null;
      const isInDoc = !!target && docEl.contains(target);
      if (!isInDoc) return;

      const range = getActiveRangeInElement(docEl);
      if (!range) {
        setCurrentSelection(null);
        return;
      }

      const selData = selectionToSelectionData(range);
      if (!selData) return;
      setCurrentSelection(selData);
    };

    document.addEventListener('keyup', handleKeyUp);
    return () => document.removeEventListener('keyup', handleKeyUp);
  }, [loading]);

  useEffect(() => {
    const pending = pendingRangeSelectionRef.current;
    if (!pending) return;
    const docEl = documentRenderRef.current;
    if (!docEl) {
      pendingRangeSelectionRef.current = null;
      return;
    }

    const selectorStart = `a[name="${CSS.escape(pending.startNodeId)}"]`;
    const selectorEnd = `a[name="${CSS.escape(pending.endNodeId)}"]`;
    const startAnchor = docEl.querySelector(selectorStart) as HTMLAnchorElement | null;
    const endAnchor = docEl.querySelector(selectorEnd) as HTMLAnchorElement | null;
    if (!startAnchor || !endAnchor) {
      pendingRangeSelectionRef.current = null;
      return;
    }

    setCurrentSelection(pending);

    const startPoint = getDomPointAtTextOffset(startAnchor, pending.startOffset);
    const endPoint = getDomPointAtTextOffset(endAnchor, pending.endOffset);
    if (startPoint && endPoint) {
      const r = document.createRange();
      r.setStart(startPoint.node, startPoint.offset);
      r.setEnd(endPoint.node, endPoint.offset);
      const sel = window.getSelection();
      if (sel) {
        sel.removeAllRanges();
        sel.addRange(r);
      }
    }

    pendingRangeSelectionRef.current = null;
  }, [htmlContent]);

  // Debounced update function
  const applyDocResponse = useCallback((res: DocumentResponse) => {
    setHtmlContent(res.html);
    setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
    const pc = res.pageCount ?? 1;
    const pi = res.pageIndex ?? 1;
    setPageCount(pc);
    setPageIndex(pi);
    setPageJumpInput(String(pi));
  }, []);

  const debouncedUpdate = useMemo(() =>
    debounce(async (sel: SelectionData | null, name: string, size: number, color: string) => {
      try {
        setLoading(true);
        let res: DocumentResponse;
        if (sel && sel.startNodeId && sel.endNodeId) {
          res = await updateRangeStyle(sel.startNodeId, sel.startOffset, sel.endNodeId, sel.endOffset, name, size, color, pageIndex);
          pendingRangeSelectionRef.current = sel;
        } else {
          res = await updateDocumentStyle(name, size, color, pageIndex);
        }
        applyDocResponse(res);
      } catch (error) {
        console.error("Failed to update document", error);
      } finally {
        setLoading(false);
      }
    }, 500)
  , [applyDocResponse, pageIndex]);

  useEffect(() => {
    return () => {
      debouncedUpdate.cancel();
    };
  }, [debouncedUpdate]);

  const handleStyleChange = (name: string, size: number, color: string) => {
    if (!currentSelection) return;
    setFontName(name);
    setFontSize(size);
    setFontColor(color);
    debouncedUpdate(currentSelection, name, size, color);
  };

  const loadPage = useCallback(async (nextPage: number) => {
    const clamped = Math.max(1, Math.min(nextPage, pageCount));
    if (loading) return;
    debouncedUpdate.cancel();
    setCurrentSelection(null);
    try {
      setLoading(true);
      const res = await fetchPageHtml(clamped);
      applyDocResponse(res);
    } catch (error) {
      console.error('Failed to load page', error);
    } finally {
      setLoading(false);
    }
  }, [applyDocResponse, debouncedUpdate, loading, pageCount]);

  const handlePrevPage = useCallback(() => {
    loadPage(pageIndex - 1);
  }, [loadPage, pageIndex]);

  const handleNextPage = useCallback(() => {
    loadPage(pageIndex + 1);
  }, [loadPage, pageIndex]);

  const handleUndo = useCallback(async () => {
    if (!history.canUndo || loading) return;
    debouncedUpdate.cancel();
    try {
      setLoading(true);
      setCurrentSelection(null);
      const res = await undoDocument(pageIndex);
      applyDocResponse(res);
    } catch (error) {
      console.error('Failed to undo', error);
    } finally {
      setLoading(false);
    }
  }, [history.canUndo, loading, debouncedUpdate, applyDocResponse, pageIndex]);

  const handleRedo = useCallback(async () => {
    if (!history.canRedo || loading) return;
    debouncedUpdate.cancel();
    try {
      setLoading(true);
      setCurrentSelection(null);
      const res = await redoDocument(pageIndex);
      applyDocResponse(res);
    } catch (error) {
      console.error('Failed to redo', error);
    } finally {
      setLoading(false);
    }
  }, [history.canRedo, loading, debouncedUpdate, applyDocResponse, pageIndex]);

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
      setCurrentSelection(null);
      const res = await uploadDocument(file, 1);
      applyDocResponse(res);
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
              currentSelection ? '范围编辑（已选中文字，可连续操作；点击其他地方取消）' : '请先选择文字'
            }
            onChange={handleStyleChange}
            onUndo={handleUndo}
            onRedo={handleRedo}
            onClearSelection={currentSelection ? () => setCurrentSelection(null) : undefined}
            rangeEnabled={!!currentSelection}
            pageIndex={pageIndex}
            pageCount={pageCount}
            onPrevPage={handlePrevPage}
            onNextPage={handleNextPage}
          />
          <div className="preview-scroll">
            <Preview html={htmlContent} loading={loading} documentRenderRef={documentRenderRef} />
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

            <div style={{ borderTop: '1px solid #eee', paddingTop: '15px', marginTop: '15px' }}>
              <h4 style={{ margin: '0 0 10px 0' }}>预览跳页</h4>

              <div style={{ display: 'flex', gap: '8px', alignItems: 'center', marginBottom: '10px' }}>
                <input
                  type="number"
                  min={1}
                  max={pageCount}
                  value={pageJumpInput}
                  onChange={(e) => setPageJumpInput(e.target.value)}
                  style={{ width: 90, padding: '8px', borderRadius: 4, border: '1px solid #ccc' }}
                  disabled={loading}
                />
                <button
                  type="button"
                  onClick={() => {
                    const v = Number(pageJumpInput);
                    if (!Number.isFinite(v)) return;
                    loadPage(v);
                  }}
                  disabled={loading}
                  style={{ padding: '8px 12px', borderRadius: 4, border: '1px solid #ccc', background: '#fff', cursor: loading ? 'wait' : 'pointer' }}
                >
                  跳转
                </button>
                <div style={{ color: '#666', fontSize: 12 }}>当前：{pageIndex}/{pageCount}</div>
              </div>

              <div style={{ maxHeight: 180, overflow: 'auto', border: '1px solid #eee', borderRadius: 6, padding: 8 }}>
                <div style={{ display: 'flex', flexWrap: 'wrap', gap: 6 }}>
                  {(() => {
                    const maxButtons = 200;
                    const total = Math.min(pageCount, maxButtons);
                    const items: JSX.Element[] = [];
                    for (let p = 1; p <= total; p += 1) {
                      items.push(
                        <button
                          key={p}
                          type="button"
                          onClick={() => loadPage(p)}
                          disabled={loading}
                          style={{
                            padding: '6px 8px',
                            borderRadius: 6,
                            border: '1px solid #ddd',
                            background: p === pageIndex ? '#2563eb' : '#fff',
                            color: p === pageIndex ? '#fff' : '#111',
                            cursor: loading ? 'wait' : 'pointer',
                          }}
                        >
                          {p}
                        </button>
                      );
                    }
                    return items;
                  })()}
                  {pageCount > 200 ? (
                    <div style={{ color: '#666', fontSize: 12, padding: '6px 4px' }}>页数较多，支持上方输入跳转</div>
                  ) : null}
                </div>
              </div>
            </div>
          </div>
        </div>
      </main>
    </div>
  );
}

export default App;
