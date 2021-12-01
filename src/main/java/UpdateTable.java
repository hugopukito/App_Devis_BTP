import com.spire.doc.Document;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.*;
import com.spire.doc.collections.ParagraphCollection;
import com.spire.doc.documents.Paragraph;

public class UpdateTable {

    private Document document;
    private Section section;
    private Table table;

    public UpdateTable(Document document) {
        this.document = document;
        document.loadFromFile("DevisVierge.docx");
        section = document.getSections().get(0);
    }

    public String ReadTable (int row, int column) {
        table = section.getTables().get(0);

        TableCell cell = table.get(row, column);
        String text = getCellText(cell);

        return text;
    }

    public String MultiplyCells (String cell1, String cell2) {
        float c1 = StringToFloat(cell1);
        float c2 = StringToFloat(cell2);
        String c3 = String.format("%.2f", c1*c2);
        return c3 = c3 + " €";
    }

    public void SetCell (int row, int column, String value) {
        TableCell cell = table.get(row, column);

        for (int i = 0; i < cell.getParagraphs().getCount(); i++) {
            cell.getParagraphs().get(i).setText(value);
        }
    }

    public float StringToFloat (String s) {

        if (s.equals("L’ens")){
            return 1.00F;
        }
        else {
            s = s.replace(',', '.');

            String [] s1 = s.split(" ");

            return Float.parseFloat(s1[0]);
        }
    }

    public static String getCellText(TableCell cell) {
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < cell.getParagraphs().getCount(); i++) {
            text.append(cell.getParagraphs().get(i).getText().trim());
        }
        return text.toString();
    }

    public void SaveDoc (String output) {
        document.saveToFile(output, FileFormat.Docx);
    }
}
