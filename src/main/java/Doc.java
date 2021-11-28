import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

public class Doc {

    private Document document;
    private Section section;
    private Paragraph[] paragraphs = new Paragraph[5];
    private String[] texts = {
        "A DOMÉRAT, ",
        "Le 18 novembre 2021",
        " Mme & Mr CLÉMENT",
        "18 Rue des Chapelettes",
        "03410 DOMÉRAT "
    };

    private String[] tableHeader = {"DEVIS N°2021-239", "", "", ""};
    private String[][] tableData = {
            new String[]{"Argentina", "Buenos Aires", "South America", "2777815"},
            new String[]{"Argentina", "Buenos Aires", "South America", "2777815"},
            new String[]{"Argentina", "Buenos Aires", "South America", "2777815"},
            new String[]{"Argentina", "Buenos Aires", "South America", "2777815"},
    };

    public Doc(Document document) {
        this.document = document;
        section = document.addSection();
    }

    public void SetTopMargin () {
        section.getPageSetup().getMargins().setTop(56.6f);
        section.getPageSetup().getMargins().setBottom(70.75f);
        section.getPageSetup().getMargins().setLeft(28.3f);
        section.getPageSetup().getMargins().setRight(28.3f);
    }

    public void CreateParagraph () {
        for (int i = 0; i < paragraphs.length; i++) {
            paragraphs[i] = section.addParagraph();
            paragraphs[i].appendText(texts[i]);
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
        paragraphs[0].getFormat().setRightIndent(139);
        paragraphs[1].getFormat().setRightIndent(107);
        paragraphs[2].getFormat().setRightIndent(96);
        paragraphs[3].getFormat().setRightIndent(98);
        paragraphs[4].getFormat().setRightIndent(106);
    }

    public void AddSpace (int para, float space) {
        paragraphs[para].getFormat().setAfterSpacing(space);
    }

    public void SaveDoc (String output) {
        document.saveToFile(output, FileFormat.Docx);
    }

    public void CreateTable () {

        // Add a table
        Table table = section.addTable(false);
        table.resetCells(tableData.length + 1, tableHeader.length);
        table.getTableFormat().setLeftIndent(5.5f);

        // Set the first row as table header
        TableRow row = table.getRows().get(0);
        row.isHeader(true);
        row.setHeight(14);
        row.setHeightType(TableRowHeightType.Exactly);

        row.getCells().get(0).setCellWidth(57, CellWidthType.Percentage);
        row.getCells().get(1).setCellWidth(9, CellWidthType.Percentage);
        row.getCells().get(2).setCellWidth(9, CellWidthType.Percentage);
        row.getCells().get(3).setCellWidth(25, CellWidthType.Percentage);

        for (int i = 0; i < tableHeader.length; i++) {
            row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
            Paragraph p = row.getCells().get(i).addParagraph();
            p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
            TextRange txtRange = p.appendText(tableHeader[i]);
            txtRange.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);
        }

        //Add data to the rest of rows
        for (int r = 0; r < tableData.length; r++) {
            TableRow dataRow = table.getRows().get(r + 1);
            dataRow.setHeight(14);
            dataRow.setHeightType(TableRowHeightType.Exactly);
            for (int c = 0; c < tableData[r].length; c++) {
                dataRow.getCells().get(c).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
                dataRow.getCells().get(c).addParagraph().appendText(tableData[r][c]);
            }
        }
    }
}
