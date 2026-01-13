from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import io
import os
import re
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

class NodeUpdate(BaseModel):
    node_id: str
    start_offset: int
    end_offset: int
    style: StyleUpdate


class RangeUpdate(BaseModel):
    start_node_id: str
    start_offset: int
    end_node_id: str
    end_offset: int
    style: StyleUpdate

# Try importing Aspose.Words, fallback to Mock if fails
try:
    import aspose.words as aw
    import aspose.pydrawing as drawing
    USE_MOCK = False
except ImportError:
    USE_MOCK = True
    print("WARNING: Aspose.Words not installed or incompatible. Using MOCK mode.")

if not USE_MOCK:
    # --- Real Aspose Implementation ---
    class DocumentState:
        def __init__(self):
            self.doc = None
            self.undo_stack: list[bytes] = []
            self.redo_stack: list[bytes] = []
            self.max_history = 50
            self.load_default()

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

        def record_change(self):
            if self.doc is None:
                return
            self._push_undo_snapshot(self._serialize_doc())
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
            self.doc = aw.Document()
            builder = aw.DocumentBuilder(self.doc)
            builder.writeln("Hello, this is a prototype document.")
            builder.writeln("You can change the font style of this text using the controls on the right.")
            builder.writeln("Aspose.Words makes document processing easy!")
            self._inject_bookmarks()
            self.undo_stack.clear()
            self.redo_stack.clear()
            
        def load_from_stream(self, file_stream):
            try:
                self.doc = aw.Document(file_stream)
                self._inject_bookmarks()
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

        def update_style(self, font_name: str = None, font_size: float = None, color: str = None):
            self.record_change()
            runs = self.doc.get_child_nodes(aw.NodeType.RUN, True)
            style = StyleUpdate(font_name=font_name, font_size=font_size, color=color)
            for run in runs:
                run = run.as_run()
                self._apply_style_to_font(run.font, style)

        def update_node_style(self, node_id: str, start_offset: int, end_offset: int, style: StyleUpdate):
            self.update_range_style(node_id, start_offset, node_id, end_offset, style)

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
            start_bm = self._find_bookmark(start_node_id)
            end_bm = self._find_bookmark(end_node_id)
            if not start_bm or not end_bm:
                return

            start_run = self._run_in_bookmark(start_bm)
            end_run = self._run_in_bookmark(end_bm)
            if not start_run or not end_run:
                return

            self.record_change()

            if start_node_id == end_node_id and start_run == end_run:
                text_len = len(start_run.text)
                if start_offset < 0:
                    start_offset = 0
                if end_offset > text_len:
                    end_offset = text_len
                if start_offset >= end_offset:
                    return
                if start_offset == 0 and end_offset == text_len:
                    self._apply_style_to_font(start_run.font, style)
                    return
                try:
                    start_bm.remove()
                except Exception:
                    pass
                self._split_run_keep_mid_id(start_run, start_node_id, start_offset, end_offset, style)
                return

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

            node = start_cursor_run.next_pre_order(self.doc)
            while node and node != end_cursor_run:
                if node.node_type == aw.NodeType.RUN:
                    self._apply_style_to_font(node.as_run().font, style)
                node = node.next_pre_order(self.doc)

            if include_end_run and end_cursor_run and end_cursor_run.node_type == aw.NodeType.RUN:
                self._apply_style_to_font(end_cursor_run.as_run().font, style)

else:
    # --- Mock Implementation ---
    class DocumentState:
        def __init__(self):
            self.undo_stack: list[dict] = []
            self.redo_stack: list[dict] = []
            self.max_history = 50
            self.load_default()

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
            self.font_name = "Arial"
            self.font_size = 12
            self.color = "#000000"
            self.text = [
                "Hello, this is a prototype document (MOCK MODE).",
                "You can change the font style of this text using the controls on the right.",
                "Aspose.Words is not installed in this env, but the logic is ready."
            ]
            self.ids = [] 
            self.undo_stack.clear()
            self.redo_stack.clear()

        def load_from_stream(self, file_stream):
            self.text = [
                "You have uploaded a new document (MOCK MODE).",
                "Since Aspose.Words is not available, I cannot render the real content.",
                "But I can simulate that the file was received successfully!"
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

@app.get("/api/init")
async def init_document(page: int = 1):
    doc_state.load_default()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}


@app.get("/api/render")
async def render_document(page: int = 1):
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}

@app.post("/api/update")
async def update_document(style: StyleUpdate, page: int = 1):
    doc_state.update_style(style.font_name, style.font_size, style.color)
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}

@app.post("/api/update_node")
async def update_node(update: NodeUpdate, page: int = 1):
    doc_state.update_node_style(update.node_id, update.start_offset, update.end_offset, update.style)
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}


@app.post("/api/update_range")
async def update_range(update: RangeUpdate, page: int = 1):
    doc_state.update_range_style(update.start_node_id, update.start_offset, update.end_node_id, update.end_offset, update.style)
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}


@app.post("/api/undo")
async def undo_document(page: int = 1):
    did_undo = doc_state.undo()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total, "didUndo": did_undo}


@app.post("/api/redo")
async def redo_document(page: int = 1):
    did_redo = doc_state.redo()
    total = doc_state.page_count() if hasattr(doc_state, "page_count") else 1
    page = max(1, min(int(page), total))
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total, "didRedo": did_redo}

@app.post("/api/upload")
async def upload_document(file: UploadFile = File(...), page: int = 1):
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
    return {"html": doc_state.get_html(page), "history": doc_state.history(), "pageIndex": page, "pageCount": total}

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
