package myprojects.jorgealamoo.wordmodifier;


import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class WordReader {
    private final XWPFDocument document;

    public WordReader(String filePath) throws IOException {
        this.document = new XWPFDocument(new FileInputStream(filePath));
    }

    public Word read(){
        Word word = new Word(document);
        List<XWPFParagraph> xwpfParagraphs = document.getParagraphs();
        List<String> paragraphs = new ArrayList<>();
        for (XWPFParagraph paragraph: xwpfParagraphs){
            paragraphs.add(paragraph.getText());
        }
        word.setParagraphs(paragraphs);

        return word;
    }

    public void close() throws IOException {
        document.close();
    }
}
