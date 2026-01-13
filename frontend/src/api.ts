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

export type DocumentResponse = {
  html: string;
  history: HistoryInfo;
  pageIndex?: number;
  pageCount?: number;
  didUndo?: boolean;
  didRedo?: boolean;
};

export const fetchInitialHtml = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.get('/init', { params: { page } });
  return response.data as DocumentResponse;
};

export const fetchPageHtml = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.get('/render', { params: { page } });
  return response.data as DocumentResponse;
};

export const updateDocumentStyle = async (fontName?: string, fontSize?: number, color?: string, page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/update', {
    font_name: fontName,
    font_size: fontSize,
    color: color,
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const updateNodeStyle = async (
  nodeId: string,
  startOffset: number,
  endOffset: number,
  fontName?: string, 
  fontSize?: number, 
  color?: string,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/update_node', {
    node_id: nodeId,
    start_offset: startOffset,
    end_offset: endOffset,
    style: {
      font_name: fontName,
      font_size: fontSize,
      color: color,
    }
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const updateRangeStyle = async (
  startNodeId: string,
  startOffset: number,
  endNodeId: string,
  endOffset: number,
  fontName?: string,
  fontSize?: number,
  color?: string,
  page?: number
): Promise<DocumentResponse> => {
  const response = await api.post('/update_range', {
    start_node_id: startNodeId,
    start_offset: startOffset,
    end_node_id: endNodeId,
    end_offset: endOffset,
    style: {
      font_name: fontName,
      font_size: fontSize,
      color: color,
    }
  }, { params: { page } });
  return response.data as DocumentResponse;
};

export const undoDocument = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/undo', undefined, { params: { page } });
  return response.data as DocumentResponse;
};

export const redoDocument = async (page?: number): Promise<DocumentResponse> => {
  const response = await api.post('/redo', undefined, { params: { page } });
  return response.data as DocumentResponse;
};

export const uploadDocument = async (file: File, page?: number): Promise<DocumentResponse> => {
  const formData = new FormData();
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
