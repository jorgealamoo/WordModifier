package myprojects.jorgealamoo.wordmodifier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class Word {
    private final XWPFDocument document;
    private final String filePath;

    public Word(String filePath) throws IOException {
        this.document = new XWPFDocument(new FileInputStream(filePath));
        this.filePath = filePath;
    }

    public XWPFDocument getDocument() {
        return document;
    }

    public List<XWPFParagraph> getParagraphs(){
        return this.document.getParagraphs();
    }

    public void addParagraph(String newParagraph){
        XWPFParagraph paragraph = this.document.createParagraph();
        paragraph.createRun().setText(newParagraph);
    }

    public void save() throws IOException {
        try (FileOutputStream outputStream = new FileOutputStream(filePath)) {
            this.document.write(outputStream);
        }
    }

    public void close() throws IOException {
        this.document.close();
    }

}
