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
  insertText,
  deleteRange,
  deleteBackward,
  deleteForward,
  insertBreak,
  type DocumentResponse,
  type HtmlPatch,
  type VersionedMeta,
} from './api';
import Preview from './components/Preview';
import Controls from './components/Controls';
import './App.css';

const DEBUG_FONT = true;

function makeClientOpId(): string {
  const c = globalThis.crypto as Crypto | undefined;
  const maybeRandomUUID = (c as any)?.randomUUID;
  if (typeof maybeRandomUUID === 'function') return maybeRandomUUID.call(c);

  if (c && typeof c.getRandomValues === 'function') {
    const bytes = new Uint8Array(16);
    c.getRandomValues(bytes);
    bytes[6] = (bytes[6] & 0x0f) | 0x40;
    bytes[8] = (bytes[8] & 0x3f) | 0x80;
    const hex = Array.from(bytes, (b) => b.toString(16).padStart(2, '0')).join('');
    return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
  }

  return `op_${Date.now().toString(36)}_${Math.random().toString(36).slice(2)}${Math.random().toString(36).slice(2)}`;
}

interface SelectionData {
  text: string;
  startNodeId: string;
  startOffset: number;
  endNodeId: string;
  endOffset: number;
  isGap?: boolean;
}

function getActiveRangeInElement(root: HTMLElement): Range | null {
  const sel = window.getSelection();
  if (!sel || sel.rangeCount === 0) return null;
  const range = sel.getRangeAt(0);
  if (!root.contains(range.commonAncestorContainer)) return null;
  return range;
}

function findRunAnchorAtPoint(
  root: HTMLElement,
  container: Node,
  offset: number,
  kind: 'start' | 'end'
): HTMLAnchorElement | null {
  // Optimization: First try to find the anchor by traversing up the DOM tree
  let current: Node | null = container;
  if (current.nodeType === Node.TEXT_NODE) {
    current = current.parentElement;
  }
  
  if (current instanceof HTMLElement) {
    const directAnchor = current.closest('a[name^="Run_"]') as HTMLAnchorElement | null;
    if (directAnchor && root.contains(directAnchor)) {
      return directAnchor;
    }
  }

  // Fallback: search by position (for points between anchors or in complex layouts)
  const anchors = Array.from(root.querySelectorAll('a[name^="Run_"]')) as HTMLAnchorElement[];
  if (anchors.length === 0) return null;

  const point = document.createRange();
  try {
    point.setStart(container, offset);
    point.collapse(true);
  } catch {
    return anchors[0] ?? null;
  }

  let lastAnchor: HTMLAnchorElement | null = null;
  for (const anchor of anchors) {
    const anchorRange = document.createRange();
    anchorRange.selectNodeContents(anchor);

    const cmpToStart = point.compareBoundaryPoints(Range.START_TO_START, anchorRange);
    if (cmpToStart <= 0) {
      // point is before this anchor starts
      // Regardless of start/end, if we are in the gap after lastAnchor, 
      // we should attach to lastAnchor to maintain consistency and avoid split ranges.
      if (lastAnchor) {
         // Auto-stitch: if we are in a text node after lastAnchor, 
          // and there are no other anchors in between, move all intermediate nodes into lastAnchor.
          if (container.nodeType === Node.TEXT_NODE) {
             let startNode: Node = lastAnchor;
             
             // If lastAnchor is last child of its parent (and parent is inline), start searching from parent's sibling
             const anchorParent = lastAnchor.parentElement;
             const isBlock = (n: Node) => n.nodeType === Node.ELEMENT_NODE && /^(P|DIV|LI|TD|TH|H[1-6]|BODY)$/i.test(n.nodeName);
             
             if (anchorParent && !isBlock(anchorParent) && anchorParent.lastChild === lastAnchor) {
                 startNode = anchorParent;
             }

             let runner = startNode.nextSibling;
             const nodesToMove: Node[] = [];
             let safeToStitch = false;
             
             // Limit search depth to avoid infinite loops or huge scans
             let steps = 0;
             while (runner && steps < 50) {
                 if (runner === container) {
                     safeToStitch = true;
                     nodesToMove.push(runner);
                     break;
                 }
                 // If we hit a block boundary or BR, stop.
                 if (isBlock(runner) || runner.nodeName === 'BR') {
                     break;
                 }
                 // If we hit another Anchor, stop.
                 if (runner.nodeType === Node.ELEMENT_NODE && (runner as Element).matches('a[name^="Run_"]')) {
                     break;
                 }
                 
                 nodesToMove.push(runner);
                 runner = runner.nextSibling;
                 steps++;
             }
             
             if (safeToStitch) {
                 if (DEBUG_FONT) console.log(`[font-debug] Auto-stitching ${nodesToMove.length} nodes into lastAnchor`);
                 nodesToMove.forEach(n => lastAnchor!.appendChild(n));
                 return lastAnchor;
             }
          }
         return lastAnchor;
      }
      // If no lastAnchor (beginning of doc), attach to this anchor (next).
      return anchor;
    }

    const cmpToEnd = point.compareBoundaryPoints(Range.START_TO_END, anchorRange);
    if (cmpToEnd <= 0) return anchor; // Inside this anchor

    lastAnchor = anchor;
  }

  return lastAnchor;
}

function selectionToSelectionData(range: Range, root: HTMLElement): SelectionData | null {
  const selection = window.getSelection();
  const text = selection?.toString() ?? '';
  const startAnchor = findRunAnchorAtPoint(root, range.startContainer, range.startOffset, 'start');
  const endAnchor = findRunAnchorAtPoint(root, range.endContainer, range.endOffset, 'end');
  
  if (DEBUG_FONT) {
    console.log('[font-debug] selectionToSelectionData range:', {
      startContainer: range.startContainer,
      startOffset: range.startOffset,
      endContainer: range.endContainer,
      endOffset: range.endOffset,
      text,
      startAnchorName: startAnchor?.getAttribute('name'),
      endAnchorName: endAnchor?.getAttribute('name')
    });
  }

  if (!startAnchor || !endAnchor) return null;

  const startNodeId = startAnchor.getAttribute('name') || '';
  const endNodeId = endAnchor.getAttribute('name') || '';
  if (!startNodeId || !endNodeId) return null;

  let startOffset = offsetWithinAnchor(startAnchor, range.startContainer, range.startOffset, 'start');
  let endOffset = offsetWithinAnchor(endAnchor, range.endContainer, range.endOffset, 'end');

  let isGap = false;

  // Fix for Gap offsets: if point is outside anchor (fallback case), add distance
  if (!startAnchor.contains(range.startContainer)) {
     const dist = getTextDistance(startAnchor, range.startContainer, range.startOffset);
     if (DEBUG_FONT) console.log(`[font-debug] gap distance start: ${dist}`);
     startOffset += dist;
     // Only treat as Gap if there is actual distance. If dist is 0 (adjacent), 
     // it's likely a boundary case that backend can handle or doesn't require delete+insert.
     if (dist > 0) isGap = true;
  }
  if (!endAnchor.contains(range.endContainer)) {
     const dist = getTextDistance(endAnchor, range.endContainer, range.endOffset);
     if (DEBUG_FONT) console.log(`[font-debug] gap distance end: ${dist}`);
     endOffset += dist;
     if (dist > 0) isGap = true;
  }

  if (DEBUG_FONT) {
    console.log('[font-debug] computed offsets:', { startOffset, endOffset });
  }

  return { text, startNodeId, startOffset, endNodeId, endOffset, isGap };
}

function getTextDistance(anchor: HTMLElement, targetNode: Node, targetOffset: number): number {
  // Calculate text length between anchor end and target point
  const range = document.createRange();
  range.setStartAfter(anchor);
  range.setEnd(targetNode, targetOffset);
  return range.toString().length;
}

function offsetWithinAnchor(anchor: HTMLAnchorElement, rangeContainer: Node, rangeOffset: number, label: string = ''): number {
  const anchorTextLen = anchor.textContent?.length ?? 0;
  const point = document.createRange();
  try {
    point.setStart(rangeContainer, rangeOffset);
    point.collapse(true);
  } catch (e) {
    if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] setStart failed`, e);
    // Fallback logic
    return 0;
  }

  // Remove early boundary checks that were causing false positives
  // Instead, rely on TreeWalker to find the exact position

  let total = 0;
  let found = false;
  const walker = document.createTreeWalker(anchor, NodeFilter.SHOW_TEXT);
  
  while (walker.nextNode()) {
    const textNode = walker.currentNode as Text;
    const len = textNode.data.length;
    
    // Check if point is inside or at boundary of this text node
    const nodeRange = document.createRange();
    nodeRange.selectNodeContents(textNode);
    
    // Check strict containment by position, not just object identity
    // This handles cases where rangeContainer might be different but refers to the same logical position,
    // or if rangeContainer is a parent element pointing into this text node.
    const startCmp = point.compareBoundaryPoints(Range.START_TO_START, nodeRange);
    const endCmp = point.compareBoundaryPoints(Range.START_TO_END, nodeRange);
    
    if (startCmp >= 0 && endCmp <= 0) {
       // Point is strictly inside or at boundaries of this text node
       // Calculate local offset
       // We can't use point.startOffset directly if containers differ.
       // We need the distance from node start to point.
       
       // Create a temp range from node start to point
       const temp = document.createRange();
       temp.setStart(textNode, 0);
       temp.setEnd(point.startContainer, point.startOffset);
       const localOffset = temp.toString().length;
       
       if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] matched by position`, total + localOffset);
       return total + localOffset;
    }

    if (rangeContainer === textNode) {
        // Exact match fallback (should be covered by above, but keep for safety)
        const offset = Math.max(0, Math.min(rangeOffset, len));
        if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] matched exact textNode`, total + offset);
        return total + offset;
    }
    
    // Check if point is before this node
    if (startCmp <= 0) {
        // point starts before (or at start of) this text node
        // If it was equal, it might mean offset 0 in this node?
        // But if rangeContainer === textNode, we handled it above.
        // So this means rangeContainer is some other node BEFORE this textNode.
        if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] point found before textNode`, total);
        return total;
    }
    
    // point starts after start of this node.
    // Does it start after end?
    if (point.compareBoundaryPoints(Range.START_TO_END, nodeRange) >= 0) {
        // Point is after this node.
        total += len;
        continue;
    }
    
    // Point is inside this node (but rangeContainer != textNode)?
    // This happens if rangeContainer is an element wrapping the text, and offset points into it?
    // But TreeWalker only gives Text nodes.
    // If rangeContainer is the parent element, setStart(parent, offset) might point to this text node.
    
    // Let's trust total += len for passed nodes.
    // If we are here, it means point is "inside" the range of this node but container != node.
    // This is rare for pure text selection but possible.
    // Let's return total for safety if we can't pinpoint.
    if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] weird overlap`, total);
    return total;
  }

  // If we finished loop, point is after all text nodes
  if (DEBUG_FONT) console.log(`[font-debug] offsetWithinAnchor [${label}] end of anchor`, total);
  return total;
}

function sanitizeRenderedHtml(html: string): string {
  if (!html) return html;
  if (typeof DOMParser === 'undefined') return html;
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const body = doc.body;

  const licenseLinks = Array.from(body.querySelectorAll('a[href*="temporary-license"]'));
  for (const a of licenseLinks) a.remove();

  const paras = Array.from(body.querySelectorAll('p'));
  for (const p of paras) {
    const hasRun = p.querySelector('a[name^="Run_"]') !== null;
    if (hasRun) continue;
    const text = (p.textContent ?? '').trim();
    if (!text) continue;
    if (/Evaluation Only/i.test(text) || /evaluation copy of aspose\.words/i.test(text) || /Created with an evaluation copy/i.test(text)) {
      p.remove();
    }
  }

  const bodyStyle = body.getAttribute('style') ?? '';
  const safeStyle = bodyStyle.replace(/"/g, '&quot;');
  return `<div data-aspose-body style="${safeStyle}">${body.innerHTML}</div>`;
}

function sanitizeFragmentHtml(html: string): string {
  if (!html) return html;
  if (typeof DOMParser === 'undefined') return html;
  const parser = new DOMParser();
  const doc = parser.parseFromString(html, 'text/html');
  const body = doc.body;

  const licenseLinks = Array.from(body.querySelectorAll('a[href*="temporary-license"]'));
  for (const a of licenseLinks) a.remove();

  const paras = Array.from(body.querySelectorAll('p'));
  for (const p of paras) {
    const hasRun = p.querySelector('a[name^="Run_"]') !== null;
    if (hasRun) continue;
    const text = (p.textContent ?? '').trim();
    if (!text) continue;
    if (/Evaluation Only/i.test(text) || /evaluation copy of aspose\.words/i.test(text) || /Created with an evaluation copy/i.test(text)) {
      p.remove();
    }
  }

  return body.innerHTML;
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
  const [docId, setDocId] = useState<string>('');
  const [docVersion, setDocVersion] = useState<number>(0);
  const [htmlContent, setHtmlContent] = useState<string>('');
  const [fontName, setFontName] = useState<string>('Arial');
  const [fontSize, setFontSize] = useState<number>(12);
  const [fontColor, setFontColor] = useState<string>('#000000');
  const [alignment, setAlignment] = useState<string>('left');
  const [firstLineIndent, setFirstLineIndent] = useState<number>(0);
  const [bold, setBold] = useState<boolean>(false);
  const [italic, setItalic] = useState<boolean>(false);
  const [loading, setLoading] = useState<boolean>(false);
  const [downloading, setDownloading] = useState<boolean>(false);
  const [currentSelection, setCurrentSelection] = useState<SelectionData | null>(null);
  const [history, setHistory] = useState<{ canUndo: boolean; canRedo: boolean }>({ canUndo: false, canRedo: false });
  const [pageIndex, setPageIndex] = useState<number>(1);
  const [pageCount, setPageCount] = useState<number>(1);
  const [pageJumpInput, setPageJumpInput] = useState<string>('1');

  const documentRenderRef = useRef<HTMLDivElement | null>(null);
  const pendingRangeSelectionRef = useRef<SelectionData | null>(null);
  const selectionFallbackRef = useRef<SelectionData | null>(null);
  const docIdRef = useRef<string>('');
  const docVersionRef = useRef<number>(0);
  const keyHandlingRef = useRef<boolean>(false);
  const composingRef = useRef<boolean>(false);
  const selectionVersionRef = useRef<number>(0);

  const opQueueRef = useRef<Promise<void>>(Promise.resolve());
  const opSeqRef = useRef<number>(0);
  const latestAppliedOpRef = useRef<number>(0);
  const pendingOpsCountRef = useRef<number>(0);
  const opTypeRef = useRef<Map<number, string>>(new Map());

  const busyDepthRef = useRef<number>(0);
  const loadingTimerRef = useRef<number | null>(null);

  const pendingInsertRef = useRef<{ nodeId: string; offset: number; text: string } | null>(null);
  const pendingInsertTimerRef = useRef<number | null>(null);
  const lastInsertOpIdRef = useRef<number>(0);

  const pendingDeleteRef = useRef<{ kind: 'backspace' | 'delete'; nodeId: string; offset: number; count: number } | null>(null);
  const pendingDeleteTimerRef = useRef<number | null>(null);

  // Keep track of current style state to debounce correctly
  const styleStateRef = useRef({ fontName, fontSize, fontColor, bold, italic });
  styleStateRef.current = { fontName, fontSize, fontColor, bold, italic };

  // Initial load
  useEffect(() => {
    const loadInitial = async () => {
      try {
        setLoading(true);
        const res = await fetchInitialHtml(1);
        setDocId(res.docId);
        docIdRef.current = res.docId;
        setDocVersion(res.version);
        docVersionRef.current = res.version;
        setHtmlContent(sanitizeRenderedHtml(res.html ?? ''));
        setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
        const pc = res.pageCount ?? 1;
        const pi = res.pageIndex ?? 1;
        setPageCount(pc);
        setPageIndex(pi);
        setPageJumpInput(String(pi));
      } catch (error) {
        console.error("加载初始文档失败", error);
      } finally {
        setLoading(false);
      }
    };
    loadInitial();
  }, []);

  // Handle Text Selection in Preview
  useEffect(() => {
    return;
  }, []);

  const syncSelectionFromDom = useCallback(() => {
    const docEl = documentRenderRef.current;
    if (!docEl) return;
    const range = getActiveRangeInElement(docEl);
    if (!range) {
      if (DEBUG_FONT) console.log('[font] syncSelectionFromDom: no range');
      setCurrentSelection(null);
      return;
    }
    const selData = selectionToSelectionData(range, docEl);
    if (!selData) {
      if (DEBUG_FONT) console.log('[font] syncSelectionFromDom: selectionToSelectionData=null');
      return;
    }
    setCurrentSelection(selData);
    selectionVersionRef.current += 1;
    if (DEBUG_FONT) {
      console.log('[font] syncSelectionFromDom', {
        text: selData.text,
        startNodeId: selData.startNodeId,
        startOffset: selData.startOffset,
        endNodeId: selData.endNodeId,
        endOffset: selData.endOffset,
        selectionVersion: selectionVersionRef.current,
      });
    }

    if (selData.text) {
      const element = window.getSelection()?.anchorNode?.parentElement;
      if (element) {
        const computed = window.getComputedStyle(element);
        const family = computed.fontFamily.split(',')[0].replace(/['"]/g, '');
        const size = parseFloat(computed.fontSize);
        const sizePt = Math.round(size * 0.75);
        const align = computed.textAlign;
        const indentStr = computed.textIndent;
        const indentPx = parseFloat(indentStr) || 0;
        const indentPt = Math.round(indentPx * 0.75);
        const isBold = computed.fontWeight === 'bold' || parseInt(computed.fontWeight) >= 700;
        const isItalic = computed.fontStyle === 'italic';

        setFontName(family);
        setFontSize(sizePt);
        setAlignment(align);
        setFirstLineIndent(indentPt);
        setBold(isBold);
        setItalic(isItalic);
      }
    }
  }, []);

  const handlePreviewMouseUp = useCallback(() => {
    const docEl = documentRenderRef.current;
    if (!docEl) return;
    if (pendingInsertRef.current || pendingDeleteRef.current) {
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
      });
    }
    syncSelectionFromDom();
    docEl.focus();
  }, [syncSelectionFromDom]);

  const handlePreviewKeyUp = useCallback(() => {
    if (pendingRangeSelectionRef.current) return;
    if (pendingInsertRef.current || pendingDeleteRef.current) {
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
      });
    }
    syncSelectionFromDom();
  }, [syncSelectionFromDom]);

  const beginBusy = useCallback(() => {
    busyDepthRef.current += 1;
    if (busyDepthRef.current === 1) {
      if (loadingTimerRef.current) {
        window.clearTimeout(loadingTimerRef.current);
      }
      loadingTimerRef.current = window.setTimeout(() => setLoading(true), 150);
    }
  }, []);

  const endBusy = useCallback(() => {
    busyDepthRef.current = Math.max(0, busyDepthRef.current - 1);
    if (busyDepthRef.current === 0) {
      if (loadingTimerRef.current) {
        window.clearTimeout(loadingTimerRef.current);
        loadingTimerRef.current = null;
      }
      setLoading(false);
    }
  }, []);

  const enqueueOp = useCallback((fn: (opId: number) => Promise<void>) => {
    const opId = (opSeqRef.current += 1);
    pendingOpsCountRef.current += 1;
    opQueueRef.current = opQueueRef.current
      .then(() => fn(opId))
      .catch((error) => {
        console.error('操作失败', error);
      })
      .finally(() => {
        pendingOpsCountRef.current -= 1;
      });
    return opQueueRef.current;
  }, []);

  const makeMeta = useCallback((): VersionedMeta => {
    return {
      docId: docIdRef.current,
      baseVersion: docVersionRef.current,
      clientOpId: makeClientOpId(),
    };
  }, []);

  const applyPendingRangeSelectionToDom = useCallback(() => {
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
    
    const r = document.createRange();
    
    if (startPoint) {
      r.setStart(startPoint.node, startPoint.offset);
    } else {
      // Fallback: if no text node found (e.g. empty anchor), set to start of anchor element
      r.setStart(startAnchor, 0);
    }

    if (endPoint) {
      r.setEnd(endPoint.node, endPoint.offset);
    } else {
      // Fallback
      r.setEnd(endAnchor, 0);
    }

    const sel = window.getSelection();
    if (sel) {
      sel.removeAllRanges();
      sel.addRange(r);
    }

    pendingRangeSelectionRef.current = null;
    docEl.focus();
  }, []);

  useEffect(() => {
    applyPendingRangeSelectionToDom();
  }, [applyPendingRangeSelectionToDom, htmlContent]);

  const applyHtmlPatches = useCallback((patches: HtmlPatch[]): boolean => {
    const docEl = documentRenderRef.current;
    if (!docEl) return false;
    if (!Array.isArray(patches) || patches.length === 0) return false;

    let applied = 0;
    let missingAnchor = 0;
    let missingTarget = 0;
    let missingHtml = 0;
    if (DEBUG_FONT) {
      console.log('[font] applyHtmlPatches input', {
        count: patches.length,
        sample: patches.slice(0, 5).map((p) => ({
          kind: p?.kind,
          anchorName: p?.anchorName,
          htmlLen: typeof p?.html === 'string' ? p.html.length : null,
        })),
      });
    }
    for (const patch of patches) {
      if (!patch || patch.kind !== 'paragraph') continue;
      const selector = `a[name="${CSS.escape(patch.anchorName)}"]`;
      const anchor = docEl.querySelector(selector) as HTMLAnchorElement | null;
      if (!anchor) {
        missingAnchor += 1;
        continue;
      }

      const target = anchor.closest('p,li,div,td,th') as HTMLElement | null;
      if (!target || !target.parentNode) {
        missingTarget += 1;
        continue;
      }

      if (DEBUG_FONT) {
        console.log(`[font] applying patch for ${patch.anchorName}`, {
          html: patch.html,
          targetTag: target.tagName,
          targetHtmlBefore: target.outerHTML.slice(0, 100)
        });
      }

      const template = document.createElement('template');
      template.innerHTML = sanitizeFragmentHtml(patch.html).trim();
      if (!template.content.firstChild) {
        missingHtml += 1;
        continue;
      }
      target.replaceWith(template.content);
      applied += 1;
    }
    if (DEBUG_FONT) {
      console.log('[font] applyHtmlPatches result', { applied, missingAnchor, missingTarget, missingHtml });
    }
    return applied > 0;
  }, []);

  // Debounced update function
  const applyDocResponse = useCallback((res: DocumentResponse, opId: number) => {
    setDocId(res.docId);
    docIdRef.current = res.docId;
    setDocVersion(res.version);
    docVersionRef.current = res.version;

    setHistory({ canUndo: res.history.canUndo, canRedo: res.history.canRedo });
    const pc = res.pageCount ?? 1;
    const pi = res.pageIndex ?? 1;
    setPageCount(pc);
    setPageIndex(pi);
    setPageJumpInput(String(pi));

    const isTyping = pendingInsertRef.current !== null || composingRef.current;
    // Also check if this op is older than the latest insert op (meaning there are newer inserts inflight)
    const isOutdatedInsert = opId < lastInsertOpIdRef.current;
    // Check if there are more ops in the queue waiting to be processed
    const hasMoreOps = pendingOpsCountRef.current > 1;
    
    if (DEBUG_FONT) {
        const hasHtmlPatch = (res.patches && res.patches.length > 0) || typeof res.html === 'string';
        console.log(`[font] applyDocResponse opId=${opId} lastInsert=${lastInsertOpIdRef.current} count=${pendingOpsCountRef.current} pending=${pendingInsertRef.current ? 'YES' : 'NO'} hasHtmlPatch=${hasHtmlPatch}`);
    }

    const hasHtmlPatch = (res.patches && res.patches.length > 0) || typeof res.html === 'string';
    const opType = opTypeRef.current.get(opId);
    if (opType) opTypeRef.current.delete(opId);
    const isInsertOp = opType === 'insert';

    if ((isTyping || isOutdatedInsert || hasMoreOps || isInsertOp) && hasHtmlPatch) {
        if (DEBUG_FONT) console.log(`[font] Skipping HTML patch for opId ${opId} (type=${opType}) due to active typing/queue or insert-strategy`);
        return;
    }

    const preservedText =
      pendingRangeSelectionRef.current?.text ??
      selectionFallbackRef.current?.text ??
      currentSelection?.text ??
      '';

    if (res.selection) {
      pendingRangeSelectionRef.current = {
        text: preservedText,
        startNodeId: res.selection.startNodeId,
        startOffset: res.selection.startOffset,
        endNodeId: res.selection.endNodeId,
        endOffset: res.selection.endOffset,
      };
    } else if (!pendingRangeSelectionRef.current) {
      const fallback = selectionFallbackRef.current ?? currentSelection;
      if (fallback) pendingRangeSelectionRef.current = fallback;
    }

    selectionFallbackRef.current = null;

    const didPatch = res.patches ? applyHtmlPatches(res.patches) : false;
    if (!didPatch && typeof res.html === 'string') {
      setHtmlContent(sanitizeRenderedHtml(res.html));
    }

    if (didPatch) {
      applyPendingRangeSelectionToDom();
    }
  }, [applyHtmlPatches, applyPendingRangeSelectionToDom, currentSelection]);

  const applyDocResponseSafe = useCallback((res: DocumentResponse, opId: number) => {
    if (opId < latestAppliedOpRef.current) return;
    latestAppliedOpRef.current = opId;
    applyDocResponse(res, opId);
  }, [applyDocResponse]);

  const handleConflictIfNeeded = useCallback(async (error: unknown, opId: number) => {
    const status = (error as any)?.response?.status;
    if (status !== 409) return false;

    const data = (error as any)?.response?.data;
    if (data && typeof data.docId === 'string' && typeof data.version === 'number' && data.history) {
      applyDocResponseSafe(data as DocumentResponse, opId);
      return true;
    }

    const res = await fetchPageHtml(pageIndex);
    applyDocResponseSafe(res, opId);
    return true;
  }, [applyDocResponseSafe, pageIndex]);

  const debouncedUpdate = useMemo(() =>
    debounce((sel: SelectionData | null, name: string, size: number, color: string, alignment: string, firstLineIndent: number, bold: boolean, italic: boolean, selectionVersion: number) => {
      if (DEBUG_FONT) {
        console.log('[font] debouncedUpdate called', {
          sel,
          name,
          size,
          color,
          alignment,
          firstLineIndent,
          bold,
          italic,
          pageIndex,
          selectionVersion,
          selectionVersionNow: selectionVersionRef.current,
          docId: docIdRef.current,
          baseVersion: docVersionRef.current,
        });
      }
      if (sel && !pendingRangeSelectionRef.current) {
        pendingRangeSelectionRef.current = sel;
        selectionFallbackRef.current = sel;
      } else if (!selectionFallbackRef.current && currentSelection && !pendingRangeSelectionRef.current) {
        pendingRangeSelectionRef.current = currentSelection;
        selectionFallbackRef.current = currentSelection;
      }
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
        if (selectionVersion !== selectionVersionRef.current) {
          if (DEBUG_FONT) {
            console.log('[font] debouncedUpdate skipped: stale selectionVersion', {
              opId,
              selectionVersion,
              selectionVersionNow: selectionVersionRef.current,
            });
          }
          return;
        }
        try {
          beginBusy();
          let res: DocumentResponse;
          const meta = makeMeta();
          if (sel && sel.startNodeId && sel.endNodeId) {
            // Only use delete+insert for Gap within the SAME node. 
            // For cross-node selections, updateRangeStyle is safer and delete+insert is too destructive.
            const isSingleNodeGap = sel.isGap && sel.startNodeId === sel.endNodeId;
            
            if (isSingleNodeGap) {
               // Strategy for Gap: delete then insert with style to force backend to recognize the text run
               if (DEBUG_FONT) console.log('[font] Single-node Gap detected, using delete+insert strategy');
               await deleteRange(meta, sel.startNodeId, sel.startOffset, sel.endNodeId, sel.endOffset, pageIndex);
               res = await insertText(
                  makeMeta(), // New op meta
                  sel.startNodeId, 
                  sel.startOffset, 
                  sel.text, 
                  { fontName: name, fontSize: size, color: color, alignment: alignment, firstLineIndent: firstLineIndent, bold: bold, italic: italic }, 
                  pageIndex
               );
            } else {
               if (DEBUG_FONT) {
                 console.log('[font] updateRangeStyle request', {
                   opId,
                   startNodeId: sel.startNodeId,
                   startOffset: sel.startOffset,
                   endNodeId: sel.endNodeId,
                   endOffset: sel.endOffset,
                   isGap: sel.isGap,
                   name,
                   size,
                   color,
                   alignment,
                   firstLineIndent,
                   bold,
                   italic,
                   pageIndex,
                   meta: { docId: meta.docId, baseVersion: meta.baseVersion },
                 });
               }
               res = await updateRangeStyle(meta, sel.startNodeId, sel.startOffset, sel.endNodeId, sel.endOffset, name, size, color, alignment, firstLineIndent, bold, italic, pageIndex);
            }
          } else {
            if (DEBUG_FONT) {
              console.log('[font] updateDocumentStyle request', {
                opId,
                name,
                size,
                color,
                alignment,
                firstLineIndent,
                bold,
                italic,
                pageIndex,
                meta: { docId: meta.docId, baseVersion: meta.baseVersion },
              });
            }
            res = await updateDocumentStyle(meta, name, size, color, alignment, firstLineIndent, bold, italic, pageIndex);
          }
          if (DEBUG_FONT) {
            console.log('[font] style update response', {
              opId,
              docId: res.docId,
              version: res.version,
              patches: Array.isArray(res.patches) ? res.patches.length : 0,
              hasHtml: typeof res.html === 'string',
              selection: res.selection,
              history: res.history,
              pageIndex: res.pageIndex,
              pageCount: res.pageCount,
            });
          }
          applyDocResponseSafe(res, opId);
        } catch (error) {
          if (!(await handleConflictIfNeeded(error, opId))) console.error("更新文档失败", error);
        } finally {
          endBusy();
        }
      });
    }, 250)
  , [applyDocResponseSafe, beginBusy, currentSelection, endBusy, enqueueOp, pageIndex]);

  useEffect(() => {
    return () => {
      debouncedUpdate.cancel();
    };
  }, [debouncedUpdate]);

  const handleStyleChange = (name: string, size: number, color: string, alignment: string, firstLineIndent: number, bold: boolean, italic: boolean) => {
    if (DEBUG_FONT) console.log('[font] toolbar change', { name, size, color, alignment, firstLineIndent, bold, italic, currentSelection });
    if (!currentSelection || !currentSelection.text) return;
    setFontName(name);
    setFontSize(size);
    setFontColor(color);
    setAlignment(alignment);
    setFirstLineIndent(firstLineIndent);
    setBold(bold);
    setItalic(italic);
    debouncedUpdate(currentSelection, name, size, color, alignment, firstLineIndent, bold, italic, selectionVersionRef.current);
  };

  const loadPage = useCallback(async (nextPage: number) => {
    const clamped = Math.max(1, Math.min(nextPage, pageCount));
    debouncedUpdate.cancel();
    const snapshot = currentSelection;
    setCurrentSelection(null);
    if (snapshot && !pendingRangeSelectionRef.current) {
      pendingRangeSelectionRef.current = snapshot;
      selectionFallbackRef.current = snapshot;
    }
    enqueueOp(async (opId) => {
      await flushPendingInsert(opId);
      await flushPendingDelete(opId);
      try {
        beginBusy();
        const res = await fetchPageHtml(clamped);
        applyDocResponseSafe(res, opId);
      } catch (error) {
        console.error('加载页面失败', error);
      } finally {
        endBusy();
      }
    });
  }, [applyDocResponseSafe, beginBusy, currentSelection, debouncedUpdate, endBusy, enqueueOp, pageCount]);

  const handlePrevPage = useCallback(() => {
    loadPage(pageIndex - 1);
  }, [loadPage, pageIndex]);

  const handleNextPage = useCallback(() => {
    loadPage(pageIndex + 1);
  }, [loadPage, pageIndex]);

  const handleUndo = useCallback(async () => {
    if (!history.canUndo) return;
    debouncedUpdate.cancel();
    const snapshot = currentSelection;
    setCurrentSelection(null);
    if (snapshot && !pendingRangeSelectionRef.current) {
      pendingRangeSelectionRef.current = snapshot;
      selectionFallbackRef.current = snapshot;
    }
    enqueueOp(async (opId) => {
      await flushPendingInsert(opId);
      await flushPendingDelete(opId);
      try {
        beginBusy();
        const res = await undoDocument(makeMeta(), pageIndex);
        applyDocResponseSafe(res, opId);
      } catch (error) {
        if (!(await handleConflictIfNeeded(error, opId))) console.error('撤销操作失败', error);
      } finally {
        endBusy();
      }
    });
  }, [applyDocResponseSafe, beginBusy, currentSelection, debouncedUpdate, endBusy, enqueueOp, history.canUndo, pageIndex]);

  const handleRedo = useCallback(async () => {
    if (!history.canRedo) return;
    debouncedUpdate.cancel();
    const snapshot = currentSelection;
    setCurrentSelection(null);
    if (snapshot && !pendingRangeSelectionRef.current) {
      pendingRangeSelectionRef.current = snapshot;
      selectionFallbackRef.current = snapshot;
    }
    enqueueOp(async (opId) => {
      await flushPendingInsert(opId);
      await flushPendingDelete(opId);
      try {
        beginBusy();
        const res = await redoDocument(makeMeta(), pageIndex);
        applyDocResponseSafe(res, opId);
      } catch (error) {
        if (!(await handleConflictIfNeeded(error, opId))) console.error('重做操作失败', error);
      } finally {
        endBusy();
      }
    });
  }, [applyDocResponseSafe, beginBusy, currentSelection, debouncedUpdate, endBusy, enqueueOp, history.canRedo, pageIndex]);

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

    debouncedUpdate.cancel();
    setCurrentSelection(null);
    enqueueOp(async (opId) => {
      await flushPendingInsert(opId);
      await flushPendingDelete(opId);
      try {
        beginBusy();
        const res = await uploadDocument(makeMeta(), file, 1);
        applyDocResponseSafe(res, opId);
      } catch (error) {
        if (!(await handleConflictIfNeeded(error, opId))) {
          console.error("上传文档失败", error);
          alert("上传文档失败，请重试。");
        }
      } finally {
        endBusy();
      }
    });
  };

  const handleDownload = async () => {
    try {
      setDownloading(true);
      await enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
      });
      await downloadDocument();
    } catch (error) {
      console.error("下载文档失败", error);
      alert("下载文档失败，请重试。");
    } finally {
      setDownloading(false);
    }
  };

  const ensureCaretInsideRun = useCallback(() => {
    const docEl = documentRenderRef.current;
    if (!docEl) return;
    const first = docEl.querySelector('a[name^="Run_"]') as HTMLAnchorElement | null;
    if (!first) return;
    const len = first.textContent?.length ?? 0;
    const pt = getDomPointAtTextOffset(first, len);
    if (!pt) return;
    const r = document.createRange();
    r.setStart(pt.node, pt.offset);
    r.setEnd(pt.node, pt.offset);
    const sel = window.getSelection();
    if (!sel) return;
    sel.removeAllRanges();
    sel.addRange(r);
    docEl.focus();
  }, []);

  const getCaretSelectionData = useCallback((): SelectionData | null => {
    const docEl = documentRenderRef.current;
    if (!docEl) return null;
    const range = getActiveRangeInElement(docEl);
    if (!range) return null;
    const selData = selectionToSelectionData(range, docEl);
    if (selData) return selData;
    ensureCaretInsideRun();
    const range2 = getActiveRangeInElement(docEl);
    if (!range2) return null;
    return selectionToSelectionData(range2, docEl);
  }, [ensureCaretInsideRun]);

  const clearPendingInsertTimer = useCallback(() => {
    if (pendingInsertTimerRef.current) {
      window.clearTimeout(pendingInsertTimerRef.current);
      pendingInsertTimerRef.current = null;
    }
  }, []);

  const clearPendingDeleteTimer = useCallback(() => {
    if (pendingDeleteTimerRef.current) {
      window.clearTimeout(pendingDeleteTimerRef.current);
      pendingDeleteTimerRef.current = null;
    }
  }, []);

  const flushPendingInsert = useCallback(async (opId: number) => {
    clearPendingInsertTimer();
    const pending = pendingInsertRef.current;
    if (!pending) return;
    if (DEBUG_FONT) console.log(`[font] flushing insert: offset=${pending.offset} text="${pending.text.replace(/\n/g, '\\n')}" opId=${opId}`);
    pendingInsertRef.current = null;
    lastInsertOpIdRef.current = opId; // Mark this opId as the latest insert
    opTypeRef.current.set(opId, 'insert');
    const apply = async (buf: { nodeId: string; offset: number; text: string }) => {
      const parts = buf.text.replace(/\r\n/g, '\n').split('\n');
      let currentNodeId = buf.nodeId;
      let currentOffset = buf.offset;

      for (let i = 0; i < parts.length; i += 1) {
        const seg = parts[i];
        if (seg) {
          const nextOffset = currentOffset + seg.length;
          pendingRangeSelectionRef.current = {
            text: '',
            startNodeId: currentNodeId,
            endNodeId: currentNodeId,
            startOffset: nextOffset,
            endOffset: nextOffset,
          };
          const res = await insertText(
            makeMeta(),
            currentNodeId,
            currentOffset,
            seg,
            { fontName, fontSize, color: fontColor, alignment: alignment, bold: bold, italic: italic },
            pageIndex
          );
          applyDocResponseSafe(res, opId);
          currentOffset = nextOffset;
        }
        if (i < parts.length - 1) {
          const brRes = await insertBreak(makeMeta(), currentNodeId, currentOffset, pageIndex);
          applyDocResponseSafe(brRes, opId);
          currentNodeId = brRes.selection?.startNodeId ?? currentNodeId;
          currentOffset = brRes.selection?.startOffset ?? 0;
        }
      }
    };
    try {
      beginBusy();
      debouncedUpdate.cancel();
      await apply(pending);
    } catch (error) {
      await handleConflictIfNeeded(error, opId);
    } finally {
      endBusy();
    }
  }, [applyDocResponseSafe, beginBusy, clearPendingInsertTimer, debouncedUpdate, endBusy, fontColor, fontName, fontSize, alignment, bold, italic, handleConflictIfNeeded, makeMeta, pageIndex]);

  const flushPendingDelete = useCallback(async (opId: number) => {
    clearPendingDeleteTimer();
    const pending = pendingDeleteRef.current;
    if (!pending) return;
    pendingDeleteRef.current = null;
    try {
      beginBusy();
      debouncedUpdate.cancel();
      if (pending.kind === 'backspace') {
        const res = await deleteBackward(makeMeta(), pending.nodeId, pending.offset, pending.count, pageIndex);
        applyDocResponseSafe(res, opId);
      } else {
        const res = await deleteForward(makeMeta(), pending.nodeId, pending.offset, pending.count, pageIndex);
        applyDocResponseSafe(res, opId);
      }
    } catch (error) {
      await handleConflictIfNeeded(error, opId);
    } finally {
      endBusy();
    }
  }, [applyDocResponseSafe, beginBusy, clearPendingDeleteTimer, debouncedUpdate, endBusy, handleConflictIfNeeded, makeMeta, pageIndex]);

  const scheduleInsertFlush = useCallback(() => {
    clearPendingInsertTimer();
    pendingInsertTimerRef.current = window.setTimeout(() => {
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
      });
    }, 80);
  }, [clearPendingInsertTimer, enqueueOp, flushPendingInsert]);

  const scheduleDeleteFlush = useCallback(() => {
    clearPendingDeleteTimer();
    pendingDeleteTimerRef.current = window.setTimeout(() => {
      enqueueOp(async (opId) => {
        await flushPendingDelete(opId);
      });
    }, 60);
  }, [clearPendingDeleteTimer, enqueueOp, flushPendingDelete]);

  const performDelete = useCallback(async (opId: number, kind: 'backspace' | 'delete') => {
    const selData = getCaretSelectionData();
    if (!selData) return;

    const isRange = !!selData.text;

    try {
      beginBusy();
      debouncedUpdate.cancel();
      if (isRange) {
        const res = await deleteRange(makeMeta(), selData.startNodeId, selData.startOffset, selData.endNodeId, selData.endOffset, pageIndex);
        pendingRangeSelectionRef.current = {
          text: '',
          startNodeId: selData.startNodeId,
          endNodeId: selData.startNodeId,
          startOffset: selData.startOffset,
          endOffset: selData.startOffset,
        };
        applyDocResponseSafe(res, opId);
        return;
      }

      if (kind === 'backspace') {
        const res = await deleteBackward(makeMeta(), selData.startNodeId, selData.startOffset, 1, pageIndex);
        applyDocResponseSafe(res, opId);
      } else {
        const res = await deleteForward(makeMeta(), selData.startNodeId, selData.startOffset, 1, pageIndex);
        applyDocResponseSafe(res, opId);
      }
    } catch (error) {
      if (!(await handleConflictIfNeeded(error, opId))) console.error('删除失败', error);
    } finally {
      endBusy();
    }
  }, [applyDocResponseSafe, beginBusy, debouncedUpdate, endBusy, getCaretSelectionData, handleConflictIfNeeded, makeMeta, pageIndex]);

  const performInsert = useCallback(async (opId: number, textToInsert: string) => {
    if (!textToInsert) return;
    const selData = getCaretSelectionData();
    if (!selData) return;

    const isRange = !!selData.text;
    try {
      beginBusy();
      debouncedUpdate.cancel();

      if (isRange) {
        const delRes = await deleteRange(makeMeta(), selData.startNodeId, selData.startOffset, selData.endNodeId, selData.endOffset, pageIndex);
        applyDocResponseSafe(delRes, opId);
        const nextSelAfterDel: SelectionData = {
          text: '',
          startNodeId: selData.startNodeId,
          endNodeId: selData.startNodeId,
          startOffset: selData.startOffset,
          endOffset: selData.startOffset,
        };
        pendingRangeSelectionRef.current = nextSelAfterDel;
      }

      const parts = textToInsert.replace(/\r\n/g, '\n').split('\n');
      let currentNodeId = selData.startNodeId;
      let currentOffset = selData.startOffset;

      for (let i = 0; i < parts.length; i += 1) {
        const seg = parts[i];
        if (seg) {
          const nextOffset = currentOffset + seg.length;
          pendingRangeSelectionRef.current = {
            text: '',
            startNodeId: currentNodeId,
            endNodeId: currentNodeId,
            startOffset: nextOffset,
            endOffset: nextOffset,
          };
          const res = await insertText(
            makeMeta(),
            currentNodeId,
            currentOffset,
            seg,
            { fontName, fontSize, color: fontColor, alignment: alignment, bold: bold, italic: italic },
            pageIndex
          );
          applyDocResponseSafe(res, opId);
          currentOffset = nextOffset;
        }
        if (i < parts.length - 1) {
          const brRes = await insertBreak(makeMeta(), currentNodeId, currentOffset, pageIndex);
          applyDocResponseSafe(brRes, opId);
          currentNodeId = brRes.selection?.startNodeId ?? currentNodeId;
          currentOffset = brRes.selection?.startOffset ?? 0;
        }
      }
    } catch (error) {
      if (!(await handleConflictIfNeeded(error, opId))) console.error('插入失败', error);
    } finally {
      endBusy();
    }
  }, [applyDocResponseSafe, beginBusy, debouncedUpdate, endBusy, fontColor, fontName, fontSize, alignment, bold, italic, getCaretSelectionData, handleConflictIfNeeded, makeMeta, pageIndex]);

  const handlePreviewKeyDown = useCallback((e: React.KeyboardEvent<HTMLDivElement>) => {
    if (composingRef.current) return;

    if (!e.ctrlKey && !e.metaKey && !e.altKey && e.key.length === 1) {
      e.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performInsert(opId, e.key);
        });
        return;
      }
      if (!pendingInsertRef.current) {
        pendingInsertRef.current = { nodeId: selData.startNodeId, offset: selData.startOffset, text: '' };
      }
      pendingInsertRef.current.text += e.key;
      scheduleInsertFlush();
      return;
    }

    if (e.key === 'Enter') {
      e.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performInsert(opId, '\n');
        });
        return;
      }
      if (!pendingInsertRef.current) {
        pendingInsertRef.current = { nodeId: selData.startNodeId, offset: selData.startOffset, text: '' };
      }
      pendingInsertRef.current.text += '\n';
      scheduleInsertFlush();
      return;
    }

    if (e.key === 'Backspace') {
      e.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performDelete(opId, 'backspace');
        });
        return;
      }

      // Check for boundary crossing (offset 0)
      if (selData.startOffset === 0) {
        const docEl = documentRenderRef.current;
        const currentAnchor = docEl?.querySelector(`a[name="${CSS.escape(selData.startNodeId)}"]`);
        if (currentAnchor) {
            let prev = currentAnchor.previousSibling;
            // Skip non-element or non-Run nodes (e.g. whitespace text nodes that are not gaps, or other tags)
            // Actually, we should be careful about Gaps. If prev is text node (Gap), we should probably delete from it?
            // But our Gap logic attaches gaps to LastAnchor. So if we are at offset 0 of CurrentAnchor,
            // we are strictly AFTER the gap (if any).
            // So we should look for the previous Anchor.
            while (prev && (prev.nodeType !== Node.ELEMENT_NODE || !(prev as Element).matches('a[name^="Run_"]'))) {
                prev = prev.previousSibling;
            }
            
            if (prev) {
                const prevAnchor = prev as HTMLAnchorElement;
                const prevId = prevAnchor.getAttribute('name');
                if (prevId) {
                    // Found previous run. Execute delete on it directly.
                    // First flush any pending ops
                    enqueueOp(async (opId) => {
                        await flushPendingInsert(opId);
                        await flushPendingDelete(opId); // This handles current pendingDeleteRef
                        
                        try {
                            beginBusy();
                            // Delete last char of previous run
                            const len = prevAnchor.textContent?.length ?? 0;
                            const res = await deleteBackward(makeMeta(), prevId, len, 1, pageIndex);
                            applyDocResponseSafe(res, opId);
                        } catch (error) {
                            if (!(await handleConflictIfNeeded(error, opId))) console.error('跨节点删除失败', error);
                        } finally {
                            endBusy();
                        }
                    });
                    return;
                }
            }
        }
      }

      pendingDeleteRef.current = pendingDeleteRef.current ?? { kind: 'backspace', nodeId: selData.startNodeId, offset: selData.startOffset, count: 0 };
      if (pendingDeleteRef.current.kind !== 'backspace') return;
      pendingDeleteRef.current.count += 1;
      scheduleDeleteFlush();
      return;
    }

    if (e.key === 'Delete') {
      e.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performDelete(opId, 'delete');
        });
        return;
      }
      pendingDeleteRef.current = pendingDeleteRef.current ?? { kind: 'delete', nodeId: selData.startNodeId, offset: selData.startOffset, count: 0 };
      if (pendingDeleteRef.current.kind !== 'delete') return;
      pendingDeleteRef.current.count += 1;
      scheduleDeleteFlush();
      return;
    }
  }, [enqueueOp, flushPendingDelete, flushPendingInsert, getCaretSelectionData, performDelete, performInsert, scheduleDeleteFlush, scheduleInsertFlush]);

  const handlePreviewBeforeInput = useCallback((e: React.FormEvent<HTMLDivElement>) => {
    const nativeEvent = e.nativeEvent as InputEvent;
    if (!nativeEvent) return;

    const inputType = nativeEvent.inputType;
    
    if (skipNextInputRef.current) {
      if (inputType === 'insertText' || inputType === 'insertCompositionText') {
        nativeEvent.preventDefault();
        skipNextInputRef.current = false;
        return;
      }
    }

    if (inputType === 'insertText') {
      const data = nativeEvent.data ?? '';
      if (!data) return;
      nativeEvent.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performInsert(opId, data);
        });
        return;
      }
      if (!pendingInsertRef.current) {
        pendingInsertRef.current = { nodeId: selData.startNodeId, offset: selData.startOffset, text: '' };
      }
      pendingInsertRef.current.text += data;
      scheduleInsertFlush();
      return;
    }

    if (inputType === 'insertCompositionText') {
      nativeEvent.preventDefault();
      return;
    }

    if (inputType === 'insertParagraph' || inputType === 'insertLineBreak') {
      nativeEvent.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performInsert(opId, '\n');
        });
        return;
      }
      if (!pendingInsertRef.current) {
        pendingInsertRef.current = { nodeId: selData.startNodeId, offset: selData.startOffset, text: '' };
      }
      pendingInsertRef.current.text += '\n';
      scheduleInsertFlush();
      return;
    }

    if (inputType === 'deleteContentBackward') {
      nativeEvent.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performDelete(opId, 'backspace');
        });
        return;
      }
      pendingDeleteRef.current = pendingDeleteRef.current ?? { kind: 'backspace', nodeId: selData.startNodeId, offset: selData.startOffset, count: 0 };
      if (pendingDeleteRef.current.kind !== 'backspace') return;
      pendingDeleteRef.current.count += 1;
      scheduleDeleteFlush();
      return;
    }

    if (inputType === 'deleteContentForward') {
      nativeEvent.preventDefault();
      const selData = getCaretSelectionData();
      if (!selData) return;
      if (selData.text) {
        enqueueOp(async (opId) => {
          await flushPendingInsert(opId);
          await flushPendingDelete(opId);
          await performDelete(opId, 'delete');
        });
        return;
      }
      pendingDeleteRef.current = pendingDeleteRef.current ?? { kind: 'delete', nodeId: selData.startNodeId, offset: selData.startOffset, count: 0 };
      if (pendingDeleteRef.current.kind !== 'delete') return;
      pendingDeleteRef.current.count += 1;
      scheduleDeleteFlush();
      return;
    }

    if (inputType.startsWith('insert') || inputType.startsWith('delete')) {
      nativeEvent.preventDefault();
    }
  }, [enqueueOp, flushPendingDelete, flushPendingInsert, getCaretSelectionData, performDelete, performInsert, scheduleDeleteFlush, scheduleInsertFlush]);

  const skipNextInputRef = useRef<boolean>(false);

  const handlePreviewCompositionStart = useCallback(() => {
    composingRef.current = true;
  }, []);

  const handlePreviewCompositionEnd = useCallback((e: React.CompositionEvent<HTMLDivElement>) => {
    composingRef.current = false;
    const text = e.data ?? '';
    if (!text) return;
    
    // Mark that we handled this input via composition end
    skipNextInputRef.current = true;
    // Clear flag after a short delay in case beforeinput doesn't fire (rare but possible)
    setTimeout(() => { skipNextInputRef.current = false; }, 0);

    const selData = getCaretSelectionData();
    if (!selData) return;
    if (selData.text) {
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
        await performInsert(opId, text);
      });
      return;
    }
    if (!pendingInsertRef.current) {
      const startOffset = Math.max(0, selData.startOffset - text.length);
      pendingInsertRef.current = { nodeId: selData.startNodeId, offset: startOffset, text: '' };
    }
    pendingInsertRef.current.text += text;
    scheduleInsertFlush();
  }, [enqueueOp, flushPendingDelete, flushPendingInsert, getCaretSelectionData, performInsert, scheduleInsertFlush]);

  const handlePreviewPaste = useCallback((e: React.ClipboardEvent<HTMLDivElement>) => {
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    if (!text) return;
    const selData = getCaretSelectionData();
    if (!selData) return;
    if (selData.text) {
      enqueueOp(async (opId) => {
        await flushPendingInsert(opId);
        await flushPendingDelete(opId);
        await performInsert(opId, text.replace(/\r\n/g, '\n'));
      });
      return;
    }
    if (!pendingInsertRef.current) {
      pendingInsertRef.current = { nodeId: selData.startNodeId, offset: selData.startOffset, text: '' };
    }
    pendingInsertRef.current.text += text.replace(/\r\n/g, '\n');
    scheduleInsertFlush();
  }, [enqueueOp, flushPendingDelete, flushPendingInsert, getCaretSelectionData, performInsert, scheduleInsertFlush]);

  return (
    <div className="app-container">
      <header className="app-header">
        <h1>Aspose.Words Web编辑器</h1>
      </header>
      <main className="app-main">
        <div className="preview-section">
          <Controls
            fontName={fontName}
            fontSize={fontSize}
            fontColor={fontColor}
            alignment={alignment}
            firstLineIndent={firstLineIndent}
            bold={bold}
            italic={italic}
            canUndo={history.canUndo}
            canRedo={history.canRedo}
            loading={loading}
            selectionLabel={
              currentSelection?.text ? '范围编辑（已选中文字，可连续操作；点击其他地方取消）' : '请先选择文字'
            }
            onChange={handleStyleChange}
            onUndo={handleUndo}
            onRedo={handleRedo}
            onClearSelection={currentSelection ? () => setCurrentSelection(null) : undefined}
            rangeEnabled={!!currentSelection?.text}
            pageIndex={pageIndex}
            pageCount={pageCount}
            onPrevPage={handlePrevPage}
            onNextPage={handleNextPage}
          />
          <div className="preview-scroll">
            <Preview
              html={htmlContent}
              loading={loading}
              documentRenderRef={documentRenderRef}
              onBeforeInput={handlePreviewBeforeInput}
              onKeyDown={handlePreviewKeyDown}
              onKeyUp={handlePreviewKeyUp}
              onMouseUp={handlePreviewMouseUp}
              onPaste={handlePreviewPaste}
              onCompositionStart={handlePreviewCompositionStart}
              onCompositionEnd={handlePreviewCompositionEnd}
            />
          </div>
        </div>
        <div className="controls-section">
          <div className="action-box" style={{ marginBottom: '20px', padding: '15px', border: '1px dashed #ccc', borderRadius: '8px' }}>

            
            <div style={{ marginBottom: '15px' }}>
              <label style={{ display: 'block', marginBottom: '5px', fontWeight: 'bold' }}>文档上传：</label>
              <input 
                type="file" 
                accept=".docx,.doc"  
                onChange={handleFileUpload}
                style={{ width: '100%' }}
              />
              <p style={{ fontSize: '0.8rem', color: '#666', marginTop: '5px' }}>
                支持上传 .docx, .doc 格式的文档
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
                {downloading ? '下载中...' : '下载文档'}
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
