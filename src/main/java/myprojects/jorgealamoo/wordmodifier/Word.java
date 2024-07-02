package myprojects.jorgealamoo.wordmodifier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Word {
    private XWPFDocument document;
    private List<String> paragraphs;

    public Word() {
        this.paragraphs = new ArrayList<>();
    }

    public void setDocument(XWPFDocument document) {
        this.document = document;
    }

    public XWPFDocument getDocument() {
        return document;
    }

    public List<String> getParagraphs(){
        return paragraphs;
    }

    public void setParagraphs(List<String> paragraphs) {
        this.paragraphs = paragraphs;
    }

    public void addParagraph(String newParagraph){
        this.paragraphs.add(newParagraph);
        if (document != null){
            XWPFParagraph paragraph = document.createParagraph();
            paragraph.createRun().setText(newParagraph);
        }
    }

    public void save(String newFilePath) throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(newFilePath)) {
            if (document == null){
                document = new XWPFDocument();
                for (String paragraphText : paragraphs){
                    XWPFParagraph paragraph = document.createParagraph();
                    paragraph.createRun().setText(paragraphText);
                }
            }
            document.write(outputStream);
        }
    }

    public void close() throws IOException {
        if (document != null) {
            this.document.close();
        }
    }

}
