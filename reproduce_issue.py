import aspose.words as aw
import io

def test_indent_export():
    doc = aw.Document()
    builder = aw.DocumentBuilder(doc)
    builder.writeln("Test paragraph for indent.")
    
    para = doc.first_section.body.first_paragraph
    para.paragraph_format.first_line_indent = 0.5 # 0.5 pt
    para.paragraph_format.alignment = aw.ParagraphAlignment.JUSTIFY
    
    print(f"Indent set to: {para.paragraph_format.first_line_indent}, Alignment: {para.paragraph_format.alignment}")
    
    # Test method 1: para.to_string(options)
    try:
        options = aw.saving.HtmlSaveOptions()
        options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
        options.export_images_as_base64 = True # Fix error
        html = para.to_string(options)
        print("Method 1 (para.to_string) result:")
        print(html)
    except Exception as e:
        print(f"Method 1 failed: {e}")

    # Test method 2: frag_doc
    try:
        frag_doc = aw.Document()
        frag_doc.first_section.body.remove_all_children()
        imported = frag_doc.import_node(para, True, aw.ImportFormatMode.KEEP_SOURCE_FORMATTING)
        frag_doc.first_section.body.append_child(imported)
        
        options = aw.saving.HtmlSaveOptions()
        options.css_style_sheet_type = aw.saving.CssStyleSheetType.INLINE
        options.pretty_format = True
        
        out_stream = io.BytesIO()
        frag_doc.save(out_stream, options)
        html = out_stream.getvalue().decode("utf-8")
        print("Method 2 (frag_doc) result:")
        print(html)
    except Exception as e:
        print(f"Method 2 failed: {e}")

if __name__ == "__main__":
    try:
        test_indent_export()
    except Exception as e:
        print(f"Global error: {e}")
