package myprojects.jorgealamoo.wordmodifier;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;

import java.io.FileInputStream;
import java.io.IOException;

public class WordReader {
    public static void main(String[] args) {
        try (FileInputStream fis = new FileInputStream("E:\\Usuario\\Documents\\Otros\\documento.docx");
        XWPFDocument document = new XWPFDocument(fis)) {

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                System.out.println(paragraph.getText());
            }

        } catch (IOException e) {
            System.out.println("Error. Document can´t be read");
        }
    }
}
