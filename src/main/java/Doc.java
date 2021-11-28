import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.Section;
import com.spire.doc.documents.HorizontalAlignment;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.documents.ParagraphStyle;

public class Doc {

    Document document;
    Section section;
    Paragraph[] paragraphs = new Paragraph[5];
    String[] strings = {
        "A DOMÉRAT, ",
        "Le 18 novembre 2021",
        " Mme & Mr CLÉMENT",
        "18 Rue des Chapelettes",
        "03410 DOMÉRAT "
    };

    public Doc(Document document) {
        this.document = document;
        section = document.addSection();
    }

    public void SetTopMargin () {
        section.getPageSetup().getMargins().setTop(56f);
    }

    public void CreateParagraph () {
        for (int i = 0; i < paragraphs.length; i++) {
            paragraphs[i] = section.addParagraph();
            paragraphs[i].appendText(strings[i]);
        }
    }

    public ParagraphStyle CreateStyle (String name, String Fontname, float Fontsize) {
        ParagraphStyle style = new ParagraphStyle(document);
        style.setName(name);
        style.getCharacterFormat().setFontName(Fontname);
        style.getCharacterFormat().setFontSize(Fontsize);
        document.getStyles().add(style);
        for (Paragraph p : paragraphs) {
            p.applyStyle(name);
        }
        return style;
    }

    public void ApplyStyle (ParagraphStyle style) {
        document.getStyles().add(style);
        for (Paragraph p : paragraphs) {
            p.applyStyle(style.getName());
        }
    }

    public void HorizontalAlignment () {
        for (Paragraph p : paragraphs) {
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        }
    }

    //Specific for our paragraphs array

    public void AddRightIndentation () {
        paragraphs[0].getFormat().setRightIndent(117);
        paragraphs[1].getFormat().setRightIndent(85);
        paragraphs[2].getFormat().setRightIndent(74);
        paragraphs[3].getFormat().setRightIndent(76);
        paragraphs[4].getFormat().setRightIndent(85);
    }

    public void AddSpace (int para, float space) {
        paragraphs[para].getFormat().setAfterSpacing(space);
    }

    public void SaveDoc (String output) {
        document.saveToFile(output, FileFormat.Docx);
    }
}
