package myprojects.jorgealamoo.wordmodifier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Word {
    private final XWPFDocument document;
    private List<String> paragraphs;

    public Word(XWPFDocument document) {
        this.paragraphs = new ArrayList<>();
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
            this.document.write(outputStream);
            System.out.println("El documento se ha modificado correctamente.");
        }
    }

    public void close() throws IOException {
        if (document != null) {
            this.document.close();
        }
    }

}
