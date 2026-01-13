import React from 'react';

interface PreviewProps {
  html: string;
  loading: boolean;
}

const Preview: React.FC<PreviewProps> = ({ html, loading }) => {
  return (
    <div className="preview-container">
      <div className={`preview-content ${loading ? 'loading' : ''}`}>
        {loading && <div className="loading-overlay">Updating...</div>}
        <div 
          className="document-render"
          dangerouslySetInnerHTML={{ __html: html }} 
        />
      </div>
    </div>
  );
};

export default Preview;
