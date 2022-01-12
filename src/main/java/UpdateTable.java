import com.spire.doc.Document;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.*;
import com.spire.doc.collections.ParagraphCollection;
import com.spire.doc.documents.Paragraph;

import java.text.DecimalFormat;
import java.util.regex.Pattern;

public class UpdateTable {

    private Document document;
    private Section section;
    private Table table;

    public UpdateTable(Document document) {
        this.document = document;
        document.loadFromFile("Devis.docx");
        section = document.getSections().get(0);
        table = section.getTables().get(0);
    }

    public String ReadTable (int row, int column) {

        TableCell cell = table.get(row, column);
        String text = getCellText(cell);

        return text;
    }

    public String MultiplyCells (String cell1, String cell2) {
        float c1 = StringToFloatLeft(cell1);
        float c2 = StringToFloatRight(cell2);
        String c3 = String.format("%.2f", c1*c2);
        return PlacePrice(c3);
    }

    public String AddCells (String cell1, String cell2) {
        float c1 = StringToFloatTTC(cell1);
        float c2 = StringToFloatTTC(cell2);
        String c3 = String.format("%.2f", c1+c2);
        return PlacePrice(c3);
    }

    public String PlacePrice (String s) {

        int i = s.length();

        switch (i) {
            // 5,00 €
            case 4 :
                return s = "            " + s + " €";

            // 50,00 €
            case 5:
                return s = "          " + s + " €";

            // 500,00 €
            case 6:
                return s = "        " + s + " €";

            // 5 000,00 €
            case 7:
                s = spaceForBigNumbers(1, s);
                return s = "     " + s + " €";

            // 50 000,00 €
            case 8:
                s = spaceForBigNumbers(2, s);
                return s = "   " + s + " €";

            // 500 000,00 €
            case 9:
                s = spaceForBigNumbers(3, s);
                return s = " " + s + " €";

            default:
                return s = s + " €";
        }
    }

    public String spaceForBigNumbers (int length, String s) {
        String first = s.substring(0, length);
        String second = s.substring(length);

        return first + " " + second;
    }

    public void SetCell (int row, int column, String value) {
        TableCell cell = table.get(row, column);

        for (int i = 0; i < cell.getParagraphs().getCount(); i++) {
            cell.getParagraphs().get(i).setText(value);
        }
    }

    public float StringToFloatLeft (String s) {

        if (s.contains("ens")){
            return 1.00F;
        }
        else if (s.contains("h")) {
            String [] s1 = s.split("h");
            return Float.parseFloat(s1[0]);
        }
        else {
            s = s.replace(',', '.');

            String [] s1 = s.split(" ");

            return Float.parseFloat(s1[0]);
        }
    }

    public float StringToFloatRight (String s) {
        s = s.replace(',', '.');
        s = s.replaceAll(" ", "");

        return Float.parseFloat(s);
    }

    public float StringToFloatTTC (String s) {
        s = s.replace("€", "");
        s = s.replace(",", ".");
        s = s.replaceAll(" ", "");
        s = s.replaceAll(" ", "");

        return Float.parseFloat(s);
    }

    public static String getCellText(TableCell cell) {
        StringBuilder text = new StringBuilder();
        for (int i = 0; i < cell.getParagraphs().getCount(); i++) {
            text.append(cell.getParagraphs().get(i).getText().trim());
        }
        return text.toString();
    }

    public void PlaceEverything() {

        TableRow lastRow = table.getLastRow();
        int tableLength = lastRow.getRowIndex();

        for (int i=0; i<tableLength-14; i++) {

            String s = ReadTable(i,1);

            int j = s.length();

            switch (j) {

                case 5:
                    s = "     " + s;
                    break;

                case 6:
                    s = "   " + s;
                    break;

                case 7:
                    s = "  " + s;
                    break;

                case 8:
                    break;

                default:
                    break;
            }

            SetCell(i,1,s);
        }

        for (int i=0; i<tableLength-14; i++) {

            String s = ReadTable(i,2);

            int j = s.length();

            switch (j) {

                // 5,00
                case 4 :
                    s = "       " + s;
                    break;

                // 50,00
                case 5:
                    s = "     " + s;
                    break;

                // 500,00
                case 6:
                    s = "   " + s;
                    break;


                // 5 000,00
                case 8:
                    break;

                // 50 000,00
                case 9:
                    break;


                default:
            }

            SetCell(i,2,s);
        }
    }

    public void MainUpdateTable () {
        TableRow lastRow = table.getLastRow();
        int tableLength = lastRow.getRowIndex();
        float total = 0;
        boolean deduc = false;
        String soldeHt = "";
        String ht = "1";
        String tva = "";

        for (int i=0; i<tableLength-14; i++) {

            String s = ReadTable(i,1);
            String s2 = ReadTable(i,2);

            if (!(s == "") || !(s2 == "")) {

                String s3 = MultiplyCells(s,s2);
                SetCell(i,3,s3);

                String totalString = s3.replace("€", "");
                totalString = totalString.replace(',', '.');
                totalString = totalString.replaceAll(" ", "");
                float totalStringInFloat = Float.parseFloat(totalString);

                total += totalStringInFloat;
            }
        }

        for (int i=0; i<tableLength; i++) {

            String s = ReadTable(i,0);

            if (s.equals("MONTANT H.T.")) {
                ht = PlacePrice(String.valueOf(String.format("%.2f", total)));
                SetCell(i,3,ht);
            }

            if (s.contains("Déduction Facture")) {
                deduc = true;
                String deduction = ReadTable(i, 3);
                deduction = deduction.replace("€", "");
                deduction = deduction.replace(",", ".");
                deduction = deduction.replaceAll(" ", "");
                deduction = deduction.replaceAll(" ", "");
                float deductionFloat = Float.parseFloat(deduction);
                total += deductionFloat;
            }

            if (s.contains("SOLDE H.T.")) {
                soldeHt = PlacePrice(String.format("%.2f", total));
                SetCell(i,3,soldeHt);
            }

            if (s.contains("T.V.A")) {
                String [] s1 = s.split(" ");
                String s2 = s1[1];
                s2 = s2.replace(",", ".");

                tva = PlacePrice(String.format("%.2f", total*(Float.parseFloat(s2))/100));
                SetCell(i,3,tva);
            }

            if (s.equals("MONTANT T.T.C.")) {
                String ttc;
                if (deduc){
                    ttc = AddCells(soldeHt, tva);
                } else {
                    ttc = AddCells(ht, tva);
                }
                SetCell(i,3,ttc);
            }
        }
        PlaceEverything();
    }

    public void SaveDoc (String output) {
        document.saveToFile(output, FileFormat.Docx);
    }
}
