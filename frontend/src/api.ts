import axios from 'axios';

const API_BASE_URL = '/api';

export const api = axios.create({
  baseURL: API_BASE_URL,
});

export type HistoryInfo = {
  canUndo: boolean;
  canRedo: boolean;
  undoDepth: number;
  redoDepth: number;
};

export type HtmlPatch = {
  kind: 'paragraph';
  anchorName: string;
  html: string;
};

export type VersionedMeta = {
  docId: string;
  baseVersion: number;
  clientOpId: string;
};

export type DocumentResponse = {
  docId: string;
  version: number;
  html?: string;
  patches?: HtmlPatch[];
  history: HistoryInfo;
  pageIndex?: number;
  pageCount?: number;
  didUndo?: boolean;
  didRedo?: boolean;
  selection?: {
    startNodeId: string;
    startOffset: number;
    endNodeId: string;
    endOffset: number;
  };
};

export const fetchInitialHtml = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.get('/init', { params: { page } });
  return response.data as DocumentResponse;
};

export const fetchPageHtml = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.get('/render', { params: { page } });
  return response.data as DocumentResponse;
};

export const updateDocumentStyle = async (meta: VersionedMeta, fontName?: string, fontSize?: number, color?: string, alignment?: string, firstLineIndent?: number, bold?: boolean, italic?: boolean, page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/update', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    font_name: fontName,
    font_size: fontSize,
    color: color,
    alignment: alignment,
    first_line_indent: firstLineIndent,
    bold: bold,
    italic: italic,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const updateNodeStyle = async (
  meta: VersionedMeta,
  nodeId: string,
  startOffset: number,
  endOffset: number,
  fontName?: string, 
  fontSize?: number, 
  color?: string,
  alignment?: string,
  firstLineIndent?: number,
  bold?: boolean,
  italic?: boolean,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/update_node', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    node_id: nodeId,
    start_offset: startOffset,
    end_offset: endOffset,
    style: {
      font_name: fontName,
      font_size: fontSize,
      color: color,
      alignment: alignment,
      first_line_indent: firstLineIndent,
      bold: bold,
      italic: italic,
    }
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const updateRangeStyle = async (
  meta: VersionedMeta,
  startNodeId: string,
  startOffset: number,
  endNodeId: string,
  endOffset: number,
  fontName?: string,
  fontSize?: number,
  color?: string,
  alignment?: string,
  firstLineIndent?: number,
  bold?: boolean,
  italic?: boolean,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/update_range', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    start_node_id: startNodeId,
    start_offset: startOffset,
    end_node_id: endNodeId,
    end_offset: endOffset,
    style: {
      font_name: fontName,
      font_size: fontSize,
      color: color,
      alignment: alignment,
      first_line_indent: firstLineIndent,
      bold: bold,
      italic: italic,
    }
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const undoDocument = async (meta: VersionedMeta, page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/undo', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const redoDocument = async (meta: VersionedMeta, page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/redo', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const uploadDocument = async (meta: VersionedMeta, file: File, page?: number): Promise<DocumentResponse> => {
  const formData = new FormData();
  formData.append('doc_id', meta.docId);
  formData.append('base_version', String(meta.baseVersion));
  formData.append('client_op_id', meta.clientOpId);
  formData.append('file', file);
  
  const response = await api.post('/upload', formData, {
    headers: {
      'Content-Type': 'multipart/form-data',
    },
    params: { page },
  });
  return response.data as DocumentResponse;
};

export const downloadDocument = async () => {
  const response = await api.get('/download', {
    responseType: 'blob',
  });
  
  // Create download link
  const url = window.URL.createObjectURL(new Blob([response.data]));
  const link = document.createElement('a');
  link.href = url;
  
  // Extract filename from header if possible, or default
  const contentDisposition = response.headers['content-disposition'];
  let filename = 'document.docx';
  if (contentDisposition) {
    const filenameMatch = contentDisposition.match(/filename="?(.+)"?/);
    if (filenameMatch.length === 2)
      filename = filenameMatch[1];
  }
  
  link.setAttribute('download', filename);
  document.body.appendChild(link);
  link.click();
  link.remove();
  window.URL.revokeObjectURL(url);
};

export const insertText = async (
  meta: VersionedMeta,
  nodeId: string,
  offset: number,
  text: string,
  style?: { fontName?: string; fontSize?: number; color?: string; alignment?: string; firstLineIndent?: number; bold?: boolean; italic?: boolean },
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/insert_text', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    node_id: nodeId,
    offset: offset,
    text: text,
    style: style ? {
      font_name: style.fontName,
      font_size: style.fontSize,
      color: style.color,
      alignment: style.alignment,
      first_line_indent: style.firstLineIndent,
      bold: style.bold,
      italic: style.italic,
    } : undefined,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const deleteRange = async (
  meta: VersionedMeta,
  startNodeId: string,
  startOffset: number,
  endNodeId: string,
  endOffset: number,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/delete_range', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    start_node_id: startNodeId,
    start_offset: startOffset,
    end_node_id: endNodeId,
    end_offset: endOffset,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const deleteBackward = async (
  meta: VersionedMeta,
  nodeId: string,
  offset: number,
  count: number = 1,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/delete_backward', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    node_id: nodeId,
    offset,
    count,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const deleteForward = async (
  meta: VersionedMeta,
  nodeId: string,
  offset: number,
  count: number = 1,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/delete_forward', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    node_id: nodeId,
    offset,
    count,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const insertBreak = async (
  meta: VersionedMeta,
  nodeId: string,
  offset: number,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/insert_break', {
    doc_id: meta.docId,
    base_version: meta.baseVersion,
    client_op_id: meta.clientOpId,
    node_id: nodeId,
    offset,
  }, { params: { page } });
  return response.data as DocumentResponse;
};
