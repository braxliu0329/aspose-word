import com.aspose.words.Document;
import com.aspose.words.DocumentBuilder;

public class TestAspose {
    public static void main(String[] args) throws Exception {
        System.out.println("Aspose.Words verification...");
        Document doc = new Document();
        DocumentBuilder builder = new DocumentBuilder(doc);
        builder.writeln("Hello Aspose.Words!");
        doc.save("Output.docx");
        System.out.println("Document created successfully: Output.docx");
    }
}
