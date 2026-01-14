import React from 'react';

interface PreviewProps {
  html: string;
  loading: boolean;
  documentRenderRef?: React.Ref<HTMLDivElement>;
  onBeforeInput?: (e: React.FormEvent<HTMLDivElement>) => void;
  onKeyDown?: (e: React.KeyboardEvent<HTMLDivElement>) => void;
  onKeyUp?: (e: React.KeyboardEvent<HTMLDivElement>) => void;
  onMouseUp?: (e: React.MouseEvent<HTMLDivElement>) => void;
  onPaste?: (e: React.ClipboardEvent<HTMLDivElement>) => void;
  onCompositionStart?: (e: React.CompositionEvent<HTMLDivElement>) => void;
  onCompositionEnd?: (e: React.CompositionEvent<HTMLDivElement>) => void;
}

const Preview: React.FC<PreviewProps> = ({
  html,
  loading,
  documentRenderRef,
  onBeforeInput,
  onKeyDown,
  onKeyUp,
  onMouseUp,
  onPaste,
  onCompositionStart,
  onCompositionEnd,
}) => {
  return (
    <div className="preview-container">
      <div className={`preview-content ${loading ? 'loading' : ''}`}>
        {loading && <div className="loading-overlay">Updating...</div>}
        <div 
          className="document-render"
          ref={documentRenderRef}
          dangerouslySetInnerHTML={{ __html: html }}
          contentEditable={true}
          suppressContentEditableWarning={true}
          spellCheck={false}
          onBeforeInput={onBeforeInput}
          onKeyDown={onKeyDown}
          onKeyUp={onKeyUp}
          onMouseUp={onMouseUp}
          onPaste={onPaste}
          onCompositionStart={onCompositionStart}
          onCompositionEnd={onCompositionEnd}
          style={{ outline: 'none', cursor: 'text' }}
        />
      </div>
    </div>
  );
};

export default Preview;
