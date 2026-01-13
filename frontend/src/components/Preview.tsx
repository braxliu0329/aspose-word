import React from 'react';

interface PreviewProps {
  html: string;
  loading: boolean;
  documentRenderRef?: React.Ref<HTMLDivElement>;
}

const Preview: React.FC<PreviewProps> = ({ html, loading, documentRenderRef }) => {
  return (
    <div className="preview-container">
      <div className={`preview-content ${loading ? 'loading' : ''}`}>
        {loading && <div className="loading-overlay">Updating...</div>}
        <div 
          className="document-render"
          ref={documentRenderRef}
          dangerouslySetInnerHTML={{ __html: html }} 
        />
      </div>
    </div>
  );
};

export default Preview;
