import React from 'react';
import { Redo2, Undo2, X } from 'lucide-react';

interface ControlsProps {
  fontName: string;
  fontSize: number;
  fontColor: string;
  canUndo: boolean;
  canRedo: boolean;
  loading?: boolean;
  selectionLabel?: string;
  onChange: (name: string, size: number, color: string) => void;
  onUndo: () => void;
  onRedo: () => void;
  onClearSelection?: () => void;
  rangeEnabled?: boolean;
  pageIndex?: number;
  pageCount?: number;
  onPrevPage?: () => void;
  onNextPage?: () => void;
}

const Controls: React.FC<ControlsProps> = ({
  fontName,
  fontSize,
  fontColor,
  canUndo,
  canRedo,
  loading,
  selectionLabel,
  onChange,
  onUndo,
  onRedo,
  onClearSelection,
  rangeEnabled,
  pageIndex,
  pageCount,
  onPrevPage,
  onNextPage,
}) => {
  const fonts = [
    { label: 'Arial', value: 'Arial' },
    { label: 'Calibri', value: 'Calibri' },
    { label: 'Times New Roman', value: 'Times New Roman' },
    { label: 'Verdana', value: 'Verdana' },
    { label: 'Courier New', value: 'Courier New' },
    { label: '微软雅黑 (Microsoft YaHei)', value: 'Microsoft YaHei' },
    { label: '宋体 (SimSun)', value: 'SimSun' },
    { label: '黑体 (SimHei)', value: 'SimHei' },
    { label: '楷体 (KaiTi)', value: 'KaiTi' },
    { label: '仿宋 (FangSong)', value: 'FangSong' }
  ];

  const handleChange = (key: string, value: string | number) => {
    let newName = fontName;
    let newSize = fontSize;
    let newColor = fontColor;

    if (key === 'fontName') newName = value as string;
    if (key === 'fontSize') newSize = Number(value);
    if (key === 'fontColor') newColor = value as string;

    onChange(newName, newSize, newColor);
  };

  return (
    <div className="toolbar-container">
      <div className="toolbar-row">
        <div className="toolbar-group">
          <button
            type="button"
            className="toolbar-button"
            onClick={onUndo}
            disabled={!canUndo || !!loading}
            title="撤回 (Ctrl+Z)"
          >
            <Undo2 size={18} />
            <span>撤回</span>
          </button>
          <button
            type="button"
            className="toolbar-button"
            onClick={onRedo}
            disabled={!canRedo || !!loading}
            title="取消撤回 (Ctrl+Y / Ctrl+Shift+Z)"
          >
            <Redo2 size={18} />
            <span>取消撤回</span>
          </button>
        </div>

        <div className="toolbar-spacer" />

        <div className="toolbar-group">
          <button
            type="button"
            className="toolbar-button"
            onClick={onPrevPage}
            disabled={!!loading || !pageCount || !pageIndex || pageIndex <= 1}
            title="上一页"
          >
            <span>上一页</span>
          </button>
          <button
            type="button"
            className="toolbar-button"
            onClick={onNextPage}
            disabled={!!loading || !pageCount || !pageIndex || pageIndex >= pageCount}
            title="下一页"
          >
            <span>下一页</span>
          </button>
          <div className="toolbar-field" style={{ minWidth: 120 }}>
            <label>页码</label>
            <div style={{ padding: '6px 8px', border: '1px solid #ddd', borderRadius: 6, background: '#fff' }}>
              {pageIndex ?? 1} / {pageCount ?? 1}
            </div>
          </div>
        </div>

        <div className="toolbar-group">
          <div className="toolbar-field">
            <label>字体</label>
            <select
              value={fontName}
              onChange={(e) => handleChange('fontName', e.target.value)}
              disabled={!rangeEnabled || !!loading}
            >
              {fonts.map((font) => (
                <option key={font.value} value={font.value}>
                  {font.label}
                </option>
              ))}
            </select>
          </div>

          <div className="toolbar-field">
            <label>字号</label>
            <input
              type="number"
              value={fontSize}
              min={8}
              max={72}
              onChange={(e) => handleChange('fontSize', e.target.value)}
              disabled={!rangeEnabled || !!loading}
            />
          </div>

          <div className="toolbar-field">
            <label>颜色</label>
            <input
              type="color"
              value={fontColor}
              onChange={(e) => handleChange('fontColor', e.target.value)}
              disabled={!rangeEnabled || !!loading}
            />
          </div>
        </div>
      </div>

      {selectionLabel ? (
        <div className="toolbar-selection">
          <span className="toolbar-selection-text">{selectionLabel}</span>
          {onClearSelection ? (
            <button type="button" className="toolbar-link" onClick={onClearSelection} title="取消范围编辑">
              <X size={16} />
              <span>取消</span>
            </button>
          ) : null}
        </div>
      ) : null}
    </div>
  );
};

export default Controls;
