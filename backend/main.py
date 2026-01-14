from fastapi import FastAPI, HTTPException, UploadFile, File, Form
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel
import io
import os
import re
import time
import uuid
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI()

# Allow CORS for frontend
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

class StyleUpdate(BaseModel):
    font_name: str | None = None
    font_size: float | None = None
    color: str | None = None
    alignment: str | None = None
    first_line_indent: float | None = None
    bold: bool | None = None
    italic: bool | None = None


class VersionedRequest(BaseModel):
    doc_id: str
    base_version: int
    client_op_id: str


class UpdateDocumentRequest(VersionedRequest):
    font_name: str | None = None
    font_size: float | None = None
    color: str | None = None
    alignment: str | None = None
    first_line_indent: float | None = None
    bold: bool | None = None
    italic: bool | None = None

class NodeUpdate(VersionedRequest):
    node_id: str
    start_offset: int
    end_offset: int
    style: StyleUpdate


class RangeUpdate(VersionedRequest):
    start_node_id: str
    start_offset: int
    end_node_id: str
    end_offset: int
    style: StyleUpdate

class TextInsert(VersionedRequest):
    node_id: str
    offset: int
    text: str
    style: StyleUpdate | None = None

class RangeDelete(VersionedRequest):
    start_node_id: str
    start_offset: int
    end_node_id: str
    end_offset: int

class CaretPosition(VersionedRequest):
    node_id: str
    offset: int
    count: int = 1


class HistoryOpRequest(VersionedRequest):
    pass

# Try importing Aspose.Words, fallback to Mock if fails
try:
    import aspose.words as aw
    import aspose.pydrawing as drawing
    USE_MOCK = False
    print("INFO: Aspose.Words loaded successfully (Real Mode)")
except ImportError:
    USE_MOCK = True
    print("WARNING: Aspose.Words not installed or incompatible. Using MOCK mode.")

if not USE_MOCK:
    # --- Real Aspose Implementation ---
    class DocumentState:
        def __init__(self):
            self.doc_id = uuid.uuid4().hex
            self.version = 0
            self.doc = None
            self.undo_stack: list[bytes] = []
            self.redo_stack: list[bytes] = []
            self.max_history = 50
            self._last_edit_time = 0.0
            self._last_edit_kind: str | None = None
            self._applied_ops: dict[str, dict] = {}
            self._applied_ops_order: list[str] = []
            self._max_applied_ops = 5000
            self.load_default()

        def _reset_identity(self):
            self.doc_id = uuid.uuid4().hex
            self.version = 0

        def _op_cache_key(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None):
            p = 0 if page is None else int(page)
            return f"{req_doc_id}:{int(req_base_version)}:{client_op_id}:{p}"

        def get_cached_response(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None):
            key = self._op_cache_key(req_doc_id, req_base_version, client_op_id, page)
            return self._applied_ops.get(key)

        def cache_response(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None, payload: dict):
            key = self._op_cache_key(req_doc_id, req_base_version, client_op_id, page)
            if key in self._applied_ops:
                return
            self._applied_ops[key] = payload
            self._applied_ops_order.append(key)
            overflow = len(self._applied_ops_order) - self._max_applied_ops
            if overflow > 0:
                for _ in range(overflow):
                    oldest = self._applied_ops_order.pop(0)
                    self._applied_ops.pop(oldest, None)

        def validate_version(self, req_doc_id: str, req_base_version: int):
            if req_doc_id != self.doc_id:
                return {"error": "doc_conflict", "docId": self.doc_id, "version": self.version}
            if int(req_base_version) != int(self.version):
                return {"error": "version_conflict", "docId": self.doc_id, "version": self.version}
            return None

        def bump_version(self):
            self.version += 1

        def _serialize_doc(self) -> bytes:
            out_stream = io.BytesIO()
            self.doc.save(out_stream, aw.SaveFormat.DOCX)
            return out_stream.getvalue()

        def _restore_doc(self, doc_bytes: bytes):
            self.doc = aw.Document(io.BytesIO(doc_bytes))

        def _push_undo_snapshot(self, snapshot: bytes):
            self.undo_stack.append(snapshot)
            if len(self.undo_stack) > self.max_history:
                self.undo_stack.pop(0)

        def record_change(self, kind: str | None = None):
            if self.doc is None:
                return
            now = time.monotonic()
            if kind == "insert" and self._last_edit_kind == "insert" and (now - self._last_edit_time) < 2.5:
                self._last_edit_time = now
                return
            self._push_undo_snapshot(self._serialize_doc())
            self.redo_stack.clear()
            self._last_edit_time = now
            self._last_edit_kind = kind

        def can_undo(self) -> bool:
            return len(self.undo_stack) > 0

        def can_redo(self) -> bool:
            return len(self.redo_stack) > 0

        def history(self):
            return {
                "canUndo": self.can_undo(),
                "canRedo": self.can_redo(),
                "undoDepth": len(self.undo_stack),
                "redoDepth": len(self.redo_stack),
            }

        def undo(self) -> bool:
            if not self.can_undo():
                return False
            current = self._serialize_doc()
            snapshot = self.undo_stack.pop()
            self.redo_stack.append(current)
            if len(self.redo_stack) > self.max_history:
                self.redo_stack.pop(0)
            self._restore_doc(snapshot)
            return True

        def redo(self) -> bool:
            if not self.can_redo():
                return False
            current = self._serialize_doc()
            snapshot = self.redo_stack.pop()
            self._push_undo_snapshot(current)
            self._restore_doc(snapshot)
            return True
            
        def _clear_auto_bookmarks(self):
            # Remove existing auto-generated bookmarks to avoid duplication/mess
            # Iterating backwards when removing is safer
            for i in range(self.doc.range.bookmarks.count - 1, -1, -1):
                bm = self.doc.range.bookmarks[i]
                if bm.name.startswith("Run_"):
                    bm.remove()

        def _inject_bookmarks(self):
            self._clear_auto_bookmarks()
            
            # Traverse all Run nodes and wrap them in unique bookmarks
            runs = self.doc.get_child_nodes(aw.NodeType.RUN, True)
            
            run_list = [node.as_run() for node in runs]
            
            for run in run_list:
                run_id = f"Run_{uuid.uuid4().hex}"
                run.parent_node.insert_before(aw.BookmarkStart(self.doc, run_id), run)
                run.parent_node.insert_after(aw.BookmarkEnd(self.doc, run_id), run)

        def load_default(self):
            self._reset_identity()
            self.doc = aw.Document()
            builder = aw.DocumentBuilder(self.doc)
            builder.writeln("这是一个原型文档。")
            builder.writeln("您可以使用右侧的控件更改文本的字体样式。")
            builder.writeln("Aspose.Words 使文档处理变得简单！")
            self._inject_bookmarks()
            self.undo_stack.clear()
            self.redo_stack.clear()
            self._last_edit_time = 0.0
            self._last_edit_kind = None
            
        def load_from_stream(self, file_stream):
            try:
                self._reset_identity()
                self.doc = aw.Document(file_stream)
                self._inject_bookmarks()
                self._last_edit_time = 0.0
                self._last_edit_kind = None
            except Exception as e:
                print(f"Error loading document: {e}")
                raise HTTPException(status_code=400, detail="Invalid document format")

        def page_count(self) -> int:
            try:
                try:
                    self.doc.update_page_layout()
                except Exception:
                    pass
                c = int(getattr(self.doc, "page_count", 1) or 1)
                return max(1, c)
            except Exception:
                return 1

        def get_html(self, page: int | None = None):
            if page is None:
                options = aw.saving.HtmlSaveOptions()
                options.export_images_as_base64 = True
                options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
                options.pretty_format = True

                out_stream = io.BytesIO()
                self.doc.save(out_stream, options)
                return out_stream.getvalue().decode("utf-8")

            total = self.page_count()
            page = max(1, min(int(page), total))
            options = aw.saving.HtmlSaveOptions()
            options.export_images_as_base64 = True
            options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
            options.pretty_format = True

            out_stream = io.BytesIO()
            try:
                try:
                    self.doc.update_page_layout()
                except Exception:
                    pass
                page_doc = self.doc.extract_pages(page - 1, 1)
                page_doc.save(out_stream, options)
                return out_stream.getvalue().decode("utf-8")
            except Exception:
                try:
                    out_stream = io.BytesIO()
                    fixed = aw.saving.HtmlFixedSaveOptions()
                    try:
                        fixed.export_embedded_images = True
                    except Exception:
                        pass
                    try:
                        fixed.page_index = page - 1
                        fixed.page_count = 1
                    except Exception:
                        try:
                            fixed.page_set = aw.saving.PageSet(page - 1)
                        except Exception:
                            pass
                    self.doc.save(out_stream, fixed)
                    return out_stream.getvalue().decode("utf-8")
                except Exception as e:
                    raise HTTPException(status_code=500, detail=f"Failed to render page {page}: {e}")

        def _extract_body_inner_html(self, html: str) -> str:
            m = re.search(r"<body[^>]*>([\s\S]*?)</body>", html, flags=re.IGNORECASE)
            if m:
                return m.group(1)
            return html

        def _export_paragraph_html_fragment(self, para):
            # Try to export directly using to_string with options if supported (newer Aspose versions)
            try:
                options = aw.saving.HtmlSaveOptions()
                options.export_images_as_base64 = True
                options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
                options.pretty_format = True
                html = para.to_string(options)
                return self._extract_body_inner_html(html)
            except Exception:
                pass

            try:
                frag_doc = aw.Document()
                try:
                    section = frag_doc.sections[0]
                    body = section.body
                except Exception:
                    section = getattr(frag_doc, "first_section", None)
                    body = getattr(section, "body", None) if section is not None else None

                if body is None:
                    return None

                while body.first_child is not None:
                    body.first_child.remove()

                imported = None
                try:
                    imported = frag_doc.import_node(para, True, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
                except Exception:
                    try:
                        imported = frag_doc.import_node(para, True)
                    except Exception:
                        imported = None

                if imported is None:
                    return None

                body.append_child(imported)

                options = aw.saving.HtmlSaveOptions()
                options.export_images_as_base64 = True
                options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
                options.pretty_format = True

                out_stream = io.BytesIO()
                frag_doc.save(out_stream, options)
                return self._extract_body_inner_html(out_stream.getvalue().decode("utf-8"))
            except Exception:
                return None

        def _paragraph_from_run_bookmark(self, run_bookmark_name: str):
            bm = self._find_bookmark(run_bookmark_name)
            if not bm:
                return None
            run = self._run_in_bookmark(bm)
            if not run:
                return None
            return self._ancestor_paragraph(run)

        def make_paragraph_patch_by_run_id(self, run_bookmark_name: str):
            para = self._paragraph_from_run_bookmark(run_bookmark_name)
            if para is None:
                return None
            html = self._export_paragraph_html_fragment(para)
            if not html:
                return None
            print(f"DEBUG: Generated patch HTML for {run_bookmark_name}: {html[:200]}...")
            return {"kind": "paragraph", "anchorName": run_bookmark_name, "html": html}

        def make_single_paragraph_patch_if_same_paragraph(self, start_run_id: str, end_run_id: str):
            start_para = self._paragraph_from_run_bookmark(start_run_id)
            end_para = self._paragraph_from_run_bookmark(end_run_id)
            if start_para is None or end_para is None:
                return None
            if start_para != end_para:
                return None
            return self.make_paragraph_patch_by_run_id(start_run_id)
        
        def get_document_stream(self):
            out_stream = io.BytesIO()
            self.doc.save(out_stream, aw.SaveFormat.DOCX)
            out_stream.seek(0)
            return out_stream

        def _apply_style_to_font(self, font, style: StyleUpdate):
            if style.font_name:
                font.name = style.font_name
            if style.font_size:
                font.size = style.font_size
            if style.bold is not None:
                font.bold = style.bold
            if style.italic is not None:
                font.italic = style.italic
            if style.color:
                try:
                    if style.color.startswith("#"):
                        hex_color = style.color.lstrip('#')
                        r = int(hex_color[0:2], 16)
                        g = int(hex_color[2:4], 16)
                        b = int(hex_color[4:6], 16)
                        c = drawing.Color.from_argb(r, g, b)
                        font.color = c
                except Exception as e:
                    print(f"Error setting color: {e}")

        def _apply_style_to_paragraph(self, para, style: StyleUpdate):
            if style.alignment:
                val = style.alignment.lower()
                if val == 'left':
                    para.paragraph_format.alignment = aw.ParagraphAlignment.LEFT
                elif val == 'center':
                    para.paragraph_format.alignment = aw.ParagraphAlignment.CENTER
                elif val == 'right':
                    para.paragraph_format.alignment = aw.ParagraphAlignment.RIGHT
                elif val == 'justify':
                    para.paragraph_format.alignment = aw.ParagraphAlignment.JUSTIFY
            
            if style.first_line_indent is not None:
                para.paragraph_format.first_line_indent = style.first_line_indent

        def insert_text(self, node_id: str, offset: int, text: str, style: StyleUpdate | None = None):
            self.record_change("insert")
            bm = self._find_bookmark(node_id)
            if not bm:
                return
            run = self._run_in_bookmark(bm)
            if not run:
                return
            
            text_len = len(run.text)
            if offset < 0: offset = 0
            if offset > text_len: offset = text_len

            if not text:
                return

            if offset == 0 and text_len == 0:
                run.text = text
                if style is not None:
                    self._apply_style_to_font(run.font, style)
                return

            try:
                bm.remove()
            except Exception:
                pass

            text_original = run.text
            text_pre = text_original[:offset]
            text_post = text_original[offset:]
            parent = run.parent_node

            if text_pre:
                pre_run = run.clone(True).as_run()
                pre_run.text = text_pre
                parent.insert_before(pre_run, run)
                self._wrap_run_with_bookmark(pre_run, f"{node_id}_pre_{uuid.uuid4().hex}")

            inserted_run = run.clone(True).as_run()
            inserted_run.text = text
            if style is not None:
                self._apply_style_to_font(inserted_run.font, style)
            parent.insert_before(inserted_run, run)
            self._wrap_run_with_bookmark(inserted_run, node_id)

            if text_post:
                post_run = run.clone(True).as_run()
                post_run.text = text_post
                parent.insert_before(post_run, run)
                self._wrap_run_with_bookmark(post_run, f"{node_id}_post_{uuid.uuid4().hex}")

            run.remove()

        def delete_range(self, start_node_id: str, start_offset: int, end_node_id: str, end_offset: int):
            self.record_change("delete")
            start_bm = self._find_bookmark(start_node_id)
            end_bm = self._find_bookmark(end_node_id)
            if not start_bm or not end_bm:
                return

            start_run = self._run_in_bookmark(start_bm)
            end_run = self._run_in_bookmark(end_bm)
            if not start_run or not end_run:
                return

            if start_node_id == end_node_id and start_run == end_run:
                text_len = len(start_run.text)
                if start_offset < 0: start_offset = 0
                if end_offset > text_len: end_offset = text_len
                if start_offset >= end_offset: return
                
                start_run.text = start_run.text[:start_offset] + start_run.text[end_offset:]
                return

            start_run_len = len(start_run.text)
            end_run_len = len(end_run.text)
            if start_offset < 0: start_offset = 0
            if start_offset > start_run_len: start_offset = start_run_len
            if end_offset < 0: end_offset = 0
            if end_offset > end_run_len: end_offset = end_run_len

            start_run.text = start_run.text[:start_offset]
            end_run.text = end_run.text[end_offset:]
            
            node = start_run.next_pre_order(self.doc)
            nodes_to_remove = []
            while node and node != end_run:
                next_node = node.next_pre_order(self.doc) 
                if node.node_type == aw.NodeType.RUN:
                    nodes_to_remove.append(node)
                node = next_node
            
            for n in nodes_to_remove:
                n.remove()

        def _ancestor_paragraph(self, node):
            cur = node
            while cur is not None:
                if cur.node_type == aw.NodeType.PARAGRAPH:
                    return cur.as_paragraph()
                cur = cur.parent_node
            return None

        def _last_run_in_paragraph(self, para):
            node = para.last_child
            while node is not None:
                if node.node_type == aw.NodeType.RUN:
                    return node.as_run()
                node = node.previous_sibling
            return None

        def _first_run_in_paragraph(self, para):
            node = para.first_child
            while node is not None:
                if node.node_type == aw.NodeType.RUN:
                    return node.as_run()
                node = node.next_sibling
            return None

        def _prev_paragraph(self, para):
            node = para.previous_sibling
            while node is not None:
                if node.node_type == aw.NodeType.PARAGRAPH:
                    return node.as_paragraph()
                node = node.previous_sibling
            return None

        def _next_paragraph(self, para):
            node = para.next_sibling
            while node is not None:
                if node.node_type == aw.NodeType.PARAGRAPH:
                    return node.as_paragraph()
                node = node.next_sibling
            return None

        def _prev_run(self, run):
            node = run.previous_sibling
            while node is not None:
                if node.node_type == aw.NodeType.RUN:
                    return node.as_run()
                node = node.previous_sibling
            para = self._ancestor_paragraph(run)
            if para is None:
                return None
            prev_para = self._prev_paragraph(para)
            if prev_para is None:
                return None
            return self._last_run_in_paragraph(prev_para)

        def _next_run(self, run):
            node = run.next_sibling
            while node is not None:
                if node.node_type == aw.NodeType.RUN:
                    return node.as_run()
                node = node.next_sibling
            para = self._ancestor_paragraph(run)
            if para is None:
                return None
            next_para = self._next_paragraph(para)
            if next_para is None:
                return None
            return self._first_run_in_paragraph(next_para)

        def _bookmark_name_for_run(self, run) -> str | None:
            try:
                for b in self.doc.range.bookmarks:
                    if not getattr(b, "name", "").startswith("Run_"):
                        continue
                    r = self._run_in_bookmark(b)
                    if r is not None and r == run:
                        return b.name
            except Exception:
                return None
            return None

        def delete_backward(self, node_id: str, offset: int):
            self.record_change("delete")
            bm = self._find_bookmark(node_id)
            if not bm:
                return {"node_id": node_id, "offset": max(0, offset)}
            run = self._run_in_bookmark(bm)
            if not run:
                return {"node_id": node_id, "offset": max(0, offset)}

            text = run.text or ""
            offset = max(0, min(int(offset), len(text)))
            if offset > 0:
                run.text = text[: offset - 1] + text[offset:]
                return {"node_id": node_id, "offset": offset - 1}

            prev_run = self._prev_run(run)
            if prev_run is None:
                return {"node_id": node_id, "offset": 0}

            cur_para = self._ancestor_paragraph(run)
            prev_para = self._ancestor_paragraph(prev_run)
            prev_id = self._bookmark_name_for_run(prev_run) or node_id

            if cur_para is not None and prev_para is not None and cur_para != prev_para:
                node_to_move = cur_para.first_child
                while node_to_move is not None:
                    nxt = node_to_move.next_sibling
                    prev_para.append_child(node_to_move)
                    node_to_move = nxt
                cur_para.remove()
                return {"node_id": prev_id, "offset": len(prev_run.text or "")}

            prev_text = prev_run.text or ""
            if prev_text:
                prev_run.text = prev_text[:-1]
                return {"node_id": prev_id, "offset": len(prev_run.text or "")}
            return {"node_id": prev_id, "offset": 0}

        def delete_forward(self, node_id: str, offset: int):
            self.record_change("delete")
            bm = self._find_bookmark(node_id)
            if not bm:
                return {"node_id": node_id, "offset": max(0, offset)}
            run = self._run_in_bookmark(bm)
            if not run:
                return {"node_id": node_id, "offset": max(0, offset)}

            text = run.text or ""
            offset = max(0, min(int(offset), len(text)))
            if offset < len(text):
                run.text = text[:offset] + text[offset + 1 :]
                return {"node_id": node_id, "offset": offset}

            next_run = self._next_run(run)
            if next_run is None:
                return {"node_id": node_id, "offset": offset}

            cur_para = self._ancestor_paragraph(run)
            next_para = self._ancestor_paragraph(next_run)
            if cur_para is not None and next_para is not None and cur_para != next_para:
                node_to_move = next_para.first_child
                while node_to_move is not None:
                    nxt = node_to_move.next_sibling
                    cur_para.append_child(node_to_move)
                    node_to_move = nxt
                next_para.remove()
                return {"node_id": node_id, "offset": len(run.text or "")}

            next_text = next_run.text or ""
            if next_text:
                next_run.text = next_text[1:]
            return {"node_id": node_id, "offset": len(run.text or "")}

        def insert_break(self, node_id: str, offset: int):
            self.record_change("break")
            bm = self._find_bookmark(node_id)
            if not bm:
                return {"node_id": node_id, "offset": 0}
            run = self._run_in_bookmark(bm)
            if not run:
                return {"node_id": node_id, "offset": 0}

            para = self._ancestor_paragraph(run)
            if para is None:
                return {"node_id": node_id, "offset": 0}

            text_original = run.text or ""
            offset = max(0, min(int(offset), len(text_original)))
            text_pre = text_original[:offset]
            text_post = text_original[offset:]

            try:
                bm.remove()
            except Exception:
                pass

            parent = run.parent_node

            if text_pre:
                pre_run = run.clone(True).as_run()
                pre_run.text = text_pre
                parent.insert_before(pre_run, run)
                self._wrap_run_with_bookmark(pre_run, f"{node_id}_pre_{uuid.uuid4().hex}")

            new_para = aw.Paragraph(self.doc)
            para.parent_node.insert_after(new_para, para)

            post_run = run.clone(True).as_run()
            post_run.text = text_post
            new_para.append_child(post_run)
            self._wrap_run_with_bookmark(post_run, node_id)

            run.remove()
            return {"node_id": node_id, "offset": 0}


        def update_style(self, font_name: str = None, font_size: float = None, color: str = None, alignment: str = None, first_line_indent: float = None, bold: bool = None, italic: bool = None):
            self.record_change("style")
            style = StyleUpdate(font_name=font_name, font_size=font_size, color=color, alignment=alignment, first_line_indent=first_line_indent, bold=bold, italic=italic)
            
            # Update Runs (Font properties)
            if font_name or font_size or color or bold is not None or italic is not None:
                runs = self.doc.get_child_nodes(aw.NodeType.RUN, True)
                for run in runs:
                    run = run.as_run()
                    self._apply_style_to_font(run.font, style)

            # Update Paragraphs (Paragraph properties)
            if alignment or first_line_indent is not None:
                paras = self.doc.get_child_nodes(aw.NodeType.PARAGRAPH, True)
                for para in paras:
                    self._apply_style_to_paragraph(para.as_paragraph(), style)

        def update_node_style(self, node_id: str, start_offset: int, end_offset: int, style: StyleUpdate):
            return self.update_range_style(node_id, start_offset, node_id, end_offset, style)

        def _find_bookmark(self, bookmark_name: str):
            try:
                return self.doc.range.bookmarks[bookmark_name]
            except TypeError:
                for b in self.doc.range.bookmarks:
                    if b.name == bookmark_name:
                        return b
            return None

        def _run_in_bookmark(self, bookmark):
            current_node = bookmark.bookmark_start.next_sibling
            while current_node and current_node != bookmark.bookmark_end:
                if current_node.node_type == aw.NodeType.RUN:
                    return current_node.as_run()
                current_node = current_node.next_sibling
            return None

        def _wrap_run_with_bookmark(self, run, bookmark_name: str):
            parent = run.parent_node
            parent.insert_before(aw.BookmarkStart(self.doc, bookmark_name), run)
            parent.insert_after(aw.BookmarkEnd(self.doc, bookmark_name), run)

        def _split_run_keep_mid_id(self, run, bookmark_name: str, start_offset: int, end_offset: int, apply_style: StyleUpdate | None):
            text_original = run.text
            text_pre = text_original[:start_offset]
            text_mid = text_original[start_offset:end_offset]
            text_post = text_original[end_offset:]

            parent = run.parent_node

            pre_run = None
            mid_run = None
            post_run = None

            if text_pre:
                pre_run = run.clone(True).as_run()
                pre_run.text = text_pre
                parent.insert_before(pre_run, run)
                self._wrap_run_with_bookmark(pre_run, f"{bookmark_name}_pre_{uuid.uuid4().hex}")

            if text_mid:
                mid_run = run.clone(True).as_run()
                mid_run.text = text_mid
                if apply_style is not None:
                    self._apply_style_to_font(mid_run.font, apply_style)
                parent.insert_before(mid_run, run)
                self._wrap_run_with_bookmark(mid_run, bookmark_name)

            if text_post:
                post_run = run.clone(True).as_run()
                post_run.text = text_post
                parent.insert_before(post_run, run)
                self._wrap_run_with_bookmark(post_run, f"{bookmark_name}_post_{uuid.uuid4().hex}")

            run.remove()
            return pre_run, mid_run, post_run

        def update_range_style(self, start_node_id: str, start_offset: int, end_node_id: str, end_offset: int, style: StyleUpdate):
            print(f"DEBUG: update_range_style called with style: {style}")
            start_bm = self._find_bookmark(start_node_id)
            end_bm = self._find_bookmark(end_node_id)
            if not start_bm or not end_bm:
                return None

            start_run = self._run_in_bookmark(start_bm)
            end_run = self._run_in_bookmark(end_bm)
            if not start_run or not end_run:
                return None

            self.record_change("style")

            # Apply alignment to paragraphs in range
            if style.alignment or style.first_line_indent is not None:
                paragraphs_to_update = set()
                
                def add_para_from_node(n):
                    if n.node_type == aw.NodeType.PARAGRAPH:
                        paragraphs_to_update.add(n.as_paragraph())
                    else:
                        p = self._ancestor_paragraph(n)
                        if p: paragraphs_to_update.add(p)

                if start_run == end_run:
                    add_para_from_node(start_run)
                else:
                    curr = start_run
                    while curr and curr != end_run:
                        add_para_from_node(curr)
                        curr = curr.next_pre_order(self.doc)
                    add_para_from_node(end_run)
                
                for p in paragraphs_to_update:
                    self._apply_style_to_paragraph(p, style)

            if start_node_id == end_node_id and start_run == end_run:
                text_len = len(start_run.text)
                if start_offset < 0:
                    start_offset = 0
                if end_offset > text_len:
                    end_offset = text_len
                if start_offset >= end_offset:
                    return {"startNodeId": start_node_id, "startOffset": start_offset, "endNodeId": end_node_id, "endOffset": end_offset}
                if start_offset == 0 and end_offset == text_len:
                    self._apply_style_to_font(start_run.font, style)
                    return {"startNodeId": start_node_id, "startOffset": 0, "endNodeId": end_node_id, "endOffset": text_len}
                try:
                    start_bm.remove()
                except Exception:
                    pass
                self._split_run_keep_mid_id(start_run, start_node_id, start_offset, end_offset, style)
                return {"startNodeId": start_node_id, "startOffset": 0, "endNodeId": end_node_id, "endOffset": max(0, end_offset - start_offset)}

            start_run_len = len(start_run.text)
            end_run_len = len(end_run.text)
            if start_offset < 0:
                start_offset = 0
            if start_offset > start_run_len:
                start_offset = start_run_len
            if end_offset < 0:
                end_offset = 0
            if end_offset > end_run_len:
                end_offset = end_run_len

            start_cursor_run = start_run
            end_cursor_run = end_run
            include_end_run = True
            selection_start_offset = start_offset
            selection_end_offset = end_offset

            if start_offset == start_run_len:
                start_cursor_run = start_run
            elif start_offset == 0:
                self._apply_style_to_font(start_run.font, style)
                start_cursor_run = start_run
            else:
                try:
                    start_bm.remove()
                except Exception:
                    pass
                _, mid, _ = self._split_run_keep_mid_id(start_run, start_node_id, start_offset, start_run_len, style)
                start_cursor_run = mid if mid is not None else start_run
                selection_start_offset = 0

            if end_offset == 0:
                include_end_run = False
                end_cursor_run = end_run
            elif end_offset == end_run_len:
                self._apply_style_to_font(end_run.font, style)
                end_cursor_run = end_run
            else:
                try:
                    end_bm.remove()
                except Exception:
                    pass
                _, mid, _ = self._split_run_keep_mid_id(end_run, end_node_id, 0, end_offset, style)
                end_cursor_run = mid if mid is not None else end_run
                selection_end_offset = len(end_cursor_run.text or "") if end_cursor_run is not None else end_offset

            node = start_cursor_run.next_pre_order(self.doc)
            while node and node != end_cursor_run:
                if node.node_type == aw.NodeType.RUN:
                    self._apply_style_to_font(node.as_run().font, style)
                node = node.next_pre_order(self.doc)

            if include_end_run and end_cursor_run and end_cursor_run.node_type == aw.NodeType.RUN:
                self._apply_style_to_font(end_cursor_run.as_run().font, style)

            return {
                "startNodeId": start_node_id,
                "startOffset": selection_start_offset,
                "endNodeId": end_node_id,
                "endOffset": selection_end_offset,
            }

else:
    # --- Mock Implementation ---
    class DocumentState:
        def __init__(self):
            self.doc_id = uuid.uuid4().hex
            self.version = 0
            self.undo_stack: list[dict] = []
            self.redo_stack: list[dict] = []
            self.max_history = 50
            self._applied_ops: dict[str, dict] = {}
            self._applied_ops_order: list[str] = []
            self._max_applied_ops = 5000
            self.load_default()

        def _reset_identity(self):
            self.doc_id = uuid.uuid4().hex
            self.version = 0

        def _op_cache_key(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None):
            p = 0 if page is None else int(page)
            return f"{req_doc_id}:{int(req_base_version)}:{client_op_id}:{p}"

        def get_cached_response(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None):
            key = self._op_cache_key(req_doc_id, req_base_version, client_op_id, page)
            return self._applied_ops.get(key)

        def cache_response(self, req_doc_id: str, req_base_version: int, client_op_id: str, page: int | None, payload: dict):
            key = self._op_cache_key(req_doc_id, req_base_version, client_op_id, page)
            if key in self._applied_ops:
                return
            self._applied_ops[key] = payload
            self._applied_ops_order.append(key)
            overflow = len(self._applied_ops_order) - self._max_applied_ops
            if overflow > 0:
                for _ in range(overflow):
                    oldest = self._applied_ops_order.pop(0)
                    self._applied_ops.pop(oldest, None)

        def validate_version(self, req_doc_id: str, req_base_version: int):
            if req_doc_id != self.doc_id:
                return {"error": "doc_conflict", "docId": self.doc_id, "version": self.version}
            if int(req_base_version) != int(self.version):
                return {"error": "version_conflict", "docId": self.doc_id, "version": self.version}
            return None

        def bump_version(self):
            self.version += 1

        def _snapshot(self):
            return {
                "font_name": self.font_name,
                "font_size": self.font_size,
                "color": self.color,
                "text": list(self.text),
            }

        def _restore(self, snapshot: dict):
            self.font_name = snapshot.get("font_name", "Arial")
            self.font_size = snapshot.get("font_size", 12)
            self.color = snapshot.get("color", "#000000")
            self.text = list(snapshot.get("text", []))

        def _push_undo_snapshot(self, snapshot: dict):
            self.undo_stack.append(snapshot)
            if len(self.undo_stack) > self.max_history:
                self.undo_stack.pop(0)

        def record_change(self):
            self._push_undo_snapshot(self._snapshot())
            self.redo_stack.clear()

        def can_undo(self) -> bool:
            return len(self.undo_stack) > 0

        def can_redo(self) -> bool:
            return len(self.redo_stack) > 0

        def history(self):
            return {
                "canUndo": self.can_undo(),
                "canRedo": self.can_redo(),
                "undoDepth": len(self.undo_stack),
                "redoDepth": len(self.redo_stack),
            }

        def undo(self) -> bool:
            if not self.can_undo():
                return False
            current = self._snapshot()
            snapshot = self.undo_stack.pop()
            self.redo_stack.append(current)
            if len(self.redo_stack) > self.max_history:
                self.redo_stack.pop(0)
            self._restore(snapshot)
            return True

        def redo(self) -> bool:
            if not self.can_redo():
                return False
            current = self._snapshot()
            snapshot = self.redo_stack.pop()
            self._push_undo_snapshot(current)
            self._restore(snapshot)
            return True

        def load_default(self):
            self._reset_identity()
            self.font_name = "Arial"
            self.font_size = 12
            self.color = "#000000"
            self.text = [
                "这是一个原型文档（模拟模式）。",
                "您可以使用右侧的控件更改文本的字体样式。",
                "Aspose.Words 未安装在当前环境中，但逻辑已准备就绪。"
            ]
            self.ids = [] 
            self.undo_stack.clear()
            self.redo_stack.clear()

        def load_from_stream(self, file_stream):
            self._reset_identity()
            self.text = [
                "您已上传新文档（模拟模式）。",
                "由于 Aspose.Words 未安装在当前环境中，我无法渲染真实内容。",
                "但我可以模拟文件已成功接收！"
            ]

        def page_count(self) -> int:
            return 1

        def get_html(self, page: int | None = None):
            style = f"font-family: '{self.font_name}'; font-size: {self.font_size}pt; color: {self.color};"
            html_parts = ["<html><body>"]
            self.ids = []
            for i, line in enumerate(self.text):
                run_id = f"Run_Mock_{i}"
                self.ids.append(run_id)
                html_parts.append(f"<p style=\"{style}\"><a name=\"{run_id}\">{line}</a></p>")
            html_parts.append("</body></html>")
            return "\n".join(html_parts)
        
        def get_document_stream(self):
            content = "\n".join(self.text)
            return io.BytesIO(content.encode("utf-8"))

        def update_style(self, font_name: str = None, font_size: float = None, color: str = None):
            self.record_change()
            if font_name: self.font_name = font_name
            if font_size: self.font_size = font_size
            if color: self.color = color

        def update_node_style(self, node_id: str, start_offset: int, end_offset: int, style: StyleUpdate):
            self.record_change()
            print(f"MOCK: Updating node '{node_id}' [{start_offset}:{end_offset}] to style {style}")
            pass

        def update_range_style(self, start_node_id: str, start_offset: int, end_node_id: str, end_offset: int, style: StyleUpdate):
            self.record_change()
            print(f"MOCK: Updating range '{start_node_id}' [{start_offset}:{end_offset}] -> '{end_node_id}' to style {style}")
            pass

doc_state = DocumentState()


def _conflict_response(conflict: dict, page: int):
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    payload = {
        **conflict,
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
        "html": doc_state.get_html(page),
    }
    return JSONResponse(status_code=409, content=payload)

@app.get("/api/init")
async def init_document(page: int = 1):
    doc_state.load_default()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "html": doc_state.get_html(page),
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
    }


@app.post("/api/insert_text")
async def insert_text_endpoint(data: TextInsert, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    doc_state.insert_text(data.node_id, data.offset, data.text, data.style)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_paragraph_patch_by_run_id"):
        patch = doc_state.make_paragraph_patch_by_run_id(data.node_id)
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "history": doc_state.history(), "pageIndex": page, "pageCount": total}
    if patch:
        payload["patches"] = [patch]
        doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
        return payload
    payload["html"] = doc_state.get_html(page)
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload


@app.post("/api/delete_range")
async def delete_range_endpoint(data: RangeDelete, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    doc_state.delete_range(data.start_node_id, data.start_offset, data.end_node_id, data.end_offset)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_single_paragraph_patch_if_same_paragraph"):
        patch = doc_state.make_single_paragraph_patch_if_same_paragraph(data.start_node_id, data.end_node_id)
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "history": doc_state.history(), "pageIndex": page, "pageCount": total}
    if patch:
        payload["patches"] = [patch]
        doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
        return payload
    payload["html"] = doc_state.get_html(page)
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload

@app.post("/api/delete_backward")
async def delete_backward_endpoint(data: CaretPosition, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    sel = {"node_id": data.node_id, "offset": max(0, int(data.offset))}
    if hasattr(doc_state, "delete_backward"):
        count = max(1, int(getattr(data, "count", 1) or 1))
        for _ in range(count):
            sel = doc_state.delete_backward(sel["node_id"], sel["offset"])
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_paragraph_patch_by_run_id"):
        patch = doc_state.make_paragraph_patch_by_run_id(sel["node_id"])
    payload = {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
        **({"patches": [patch]} if patch else {"html": doc_state.get_html(page)}),
        "selection": {
            "startNodeId": sel["node_id"],
            "startOffset": sel["offset"],
            "endNodeId": sel["node_id"],
            "endOffset": sel["offset"],
        },
    }
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload

@app.post("/api/delete_forward")
async def delete_forward_endpoint(data: CaretPosition, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    sel = {"node_id": data.node_id, "offset": max(0, int(data.offset))}
    if hasattr(doc_state, "delete_forward"):
        count = max(1, int(getattr(data, "count", 1) or 1))
        for _ in range(count):
            sel = doc_state.delete_forward(sel["node_id"], sel["offset"])
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_paragraph_patch_by_run_id"):
        patch = doc_state.make_paragraph_patch_by_run_id(sel["node_id"])
    payload = {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
        **({"patches": [patch]} if patch else {"html": doc_state.get_html(page)}),
        "selection": {
            "startNodeId": sel["node_id"],
            "startOffset": sel["offset"],
            "endNodeId": sel["node_id"],
            "endOffset": sel["offset"],
        },
    }
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload

@app.post("/api/insert_break")
async def insert_break_endpoint(data: CaretPosition, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    sel = {"node_id": data.node_id, "offset": 0}
    if hasattr(doc_state, "insert_break"):
        sel = doc_state.insert_break(data.node_id, data.offset)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    payload = {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "html": doc_state.get_html(page),
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
        "selection": {
            "startNodeId": sel["node_id"],
            "startOffset": sel["offset"],
            "endNodeId": sel["node_id"],
            "endOffset": sel["offset"],
        },
    }
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload


@app.get("/api/render")
async def render_document(page: int = 1):
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "html": doc_state.get_html(page),
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
    }

@app.post("/api/update")
async def update_document(data: UpdateDocumentRequest, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    doc_state.update_style(data.font_name, data.font_size, data.color, data.alignment, data.first_line_indent, data.bold, data.italic)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    payload = {
        "docId": doc_state.doc_id,
        "version": doc_state.version,
        "html": doc_state.get_html(page),
        "history": doc_state.history(),
        "pageIndex": page,
        "pageCount": total,
    }
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload

@app.post("/api/update_node")
async def update_node(update: NodeUpdate, page: int = 1):
    cached = doc_state.get_cached_response(update.doc_id, update.base_version, update.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(update.doc_id, update.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    sel = doc_state.update_node_style(update.node_id, update.start_offset, update.end_offset, update.style)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_paragraph_patch_by_run_id"):
        patch = doc_state.make_paragraph_patch_by_run_id(update.node_id)
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "history": doc_state.history(), "pageIndex": page, "pageCount": total}
    if patch:
        payload["patches"] = [patch]
    else:
        payload["html"] = doc_state.get_html(page)
    if sel:
        payload["selection"] = sel
    doc_state.cache_response(update.doc_id, update.base_version, update.client_op_id, page, payload)
    return payload


@app.post("/api/update_range")
async def update_range(update: RangeUpdate, page: int = 1):
    cached = doc_state.get_cached_response(update.doc_id, update.base_version, update.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(update.doc_id, update.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    sel = doc_state.update_range_style(update.start_node_id, update.start_offset, update.end_node_id, update.end_offset, update.style)
    doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    patch = None
    if not USE_MOCK and hasattr(doc_state, "make_single_paragraph_patch_if_same_paragraph"):
        patch = doc_state.make_single_paragraph_patch_if_same_paragraph(update.start_node_id, update.end_node_id)
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "history": doc_state.history(), "pageIndex": page, "pageCount": total}
    if patch:
        payload["patches"] = [patch]
    else:
        payload["html"] = doc_state.get_html(page)
    if sel:
        payload["selection"] = sel
    doc_state.cache_response(update.doc_id, update.base_version, update.client_op_id, page, payload)
    return payload


@app.post("/api/undo")
async def undo_document(data: HistoryOpRequest, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    did_undo = doc_state.undo()
    if did_undo:
        doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total, "didUndo": did_undo}
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload


@app.post("/api/redo")
async def redo_document(data: HistoryOpRequest, page: int = 1):
    cached = doc_state.get_cached_response(data.doc_id, data.base_version, data.client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(data.doc_id, data.base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    did_redo = doc_state.redo()
    if did_redo:
        doc_state.bump_version()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total, "didRedo": did_redo}
    doc_state.cache_response(data.doc_id, data.base_version, data.client_op_id, page, payload)
    return payload

@app.post("/api/upload")
async def upload_document(
    file: UploadFile = File(...),
    doc_id: str = Form(...),
    base_version: int = Form(...),
    client_op_id: str = Form(...),
    page: int = 1,
):
    cached = doc_state.get_cached_response(doc_id, base_version, client_op_id, page)
    if cached is not None:
        return cached
    conflict = doc_state.validate_version(doc_id, base_version)
    if conflict is not None:
        return _conflict_response(conflict, page)
    content = await file.read()
    file_stream = io.BytesIO(content)
    snapshot = None
    if not USE_MOCK:
        snapshot = doc_state._serialize_doc() if getattr(doc_state, "doc", None) is not None else None
    else:
        snapshot = doc_state._snapshot()

    doc_state.load_from_stream(file_stream)
    if snapshot is not None:
        doc_state._push_undo_snapshot(snapshot)
        doc_state.redo_stack.clear()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = 1
    payload = {"docId": doc_state.doc_id, "version": doc_state.version, "html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}
    doc_state.cache_response(doc_id, base_version, client_op_id, page, payload)
    return payload

@app.get("/api/download")
async def download_document():
    stream = doc_state.get_document_stream()
    filename = "modified_document.docx" if not USE_MOCK else "mock_document.txt"
    media_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document" if not USE_MOCK else "text/plain"
    
    return StreamingResponse(
        stream, 
        media_type=media_type,
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
