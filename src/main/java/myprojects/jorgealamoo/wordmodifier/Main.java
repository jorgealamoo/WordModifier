package myprojects.jorgealamoo.wordmodifier;

import java.io.IOException;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        WordReader reader = new WordReader("E:\\Usuario\\Documents\\Otros\\documento.docx");
        Word word = reader.read();

        List<String> paragraphs = word.getParagraphs();
        for (String paragraph : paragraphs){
            System.out.println(paragraph);
        }

        word.addParagraph("Nuevo parrafo de prueba añadido al final del documento");

        word.save("E:\\Usuario\\Documents\\Otros\\documento_modificado.docx");
        word.close();
    }
}
