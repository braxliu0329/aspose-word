# 页面设计说明：文档预览/编辑页（含 Undo/Redo）

## 1) Layout（桌面优先）

* 页面采用“纵向堆叠（stacked）”结构：顶部控制面板（Toolbar / Control Panel）在上，Document Preview 在下。

* 主要布局：CSS Grid（外层 2 行：toolbar + content），内容区内部用 Flex 处理横向扩展与对齐。

* Toolbar 行为：

  * 桌面端默认固定在预览区域上方（可采用 `position: sticky; top: 0;`），保证像 Word 一样滚动预览时仍可见。

  * Toolbar 与预览之间保留明确分隔（1px border 或阴影），强化层级。

## 2) Meta Information

* Title：文档预览与编辑

* Description：在预览界面进行编辑，并支持撤回/取消撤回。

* Open Graph：

  * og:title = 文档预览与编辑

  * og:description = 支持 Undo/Redo 的文档预览编辑体验

## 3) Global Styles（设计变量建议）

* 背景色：页面 #F7F8FA；预览画布 #FFFFFF

* 主色（Accent）：#2F6FED

* 字体：

  * 基础字号 14px；标题/分组标签 16px

  * 等宽用于快捷键提示（如 Ctrl+Z）

* 按钮：

  * 默认：浅底 + 边框（1px #D0D5DD），hover 加深背景

  * 主操作（若项目已有主按钮样式沿用）：Accent 填充

  * Disabled：降低不透明度（如 40%）+ 禁止点击

* 图标：Undo/Redo 使用清晰方向性图标（左回退/右前进），与文本可选其一或并存（推荐“图标 + tooltip”）。

## 4) Page Structure

1. 顶部控制面板（位于 document preview 上方，横向工具栏）
2. Document Preview 区域（可滚动）

## 5) Sections & Components

### A. 顶部控制面板（Control Panel / Toolbar）

* 位置与层级

  * 位于 Document Preview 上方，横向铺满预览容器宽度。

  * 对齐：左侧放置 Undo/Redo（符合多数编辑器习惯），右侧预留给既有控制项（如果项目已有）。

* 组件构成（本需求强制项）

  1. Undo 按钮

     * 文案/Tooltip：撤回（Ctrl+Z）

     * 状态：当 UndoStack 为空时 Disabled

     * 交互：点击立即触发撤回并刷新预览
  2. Redo 按钮

     * 文案/Tooltip：取消撤回（Ctrl+Y 或 Ctrl+Shift+Z，以项目既定为准）

     * 状态：当 RedoStack 为空时 Disabled

     * 交互：点击立即触发前进并刷新预览

* 分组与间距

  * Undo/Redo 作为一个紧凑分组（button group），两按钮间距 4–8px。

  * Toolbar 内边距：12–16px；高度建议 48–56px。

### B. Document Preview 区域

* 容器

  * 占据 Toolbar 下方剩余可用空间；垂直滚动。

  * 预览画布居中展示（类似 Word 页面居中），左右留白。

* 内容刷新

  * 当 Undo/Redo 发生时，预览内容应在同一帧或极短过渡内更新（避免闪烁）。

## 6) Responsive behavior（简要）

* 桌面优先：Undo/Redo 始终可见。

* 较窄宽度时：Toolbar 可换行或将按钮收纳为仅图标（仍需 tooltip）。

## 7) Interaction & Transition Guidelines

* Toolbar hover/active 状态应清晰；Disabled 必须不可点击。

* Undo/Redo 触发后可提供轻量反馈：例如按钮短暂高亮或预览区域无闪烁更新（避免大动画）。

