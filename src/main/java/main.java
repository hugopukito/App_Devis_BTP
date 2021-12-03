import com.spire.doc.Document;
import com.spire.doc.Section;
import com.spire.doc.Table;
import com.spire.doc.TableCell;
import com.spire.doc.documents.UnderlineStyle;

import java.text.SimpleDateFormat;
import java.util.Date;

public class main {
    public static void main(String[] args) throws Exception {

        /*Doc doc = new Doc(new Document());

        doc.SetTopMargin();

        doc.CreateParagraph();

        doc.ApplyStyle(doc.CreateStyle("style1", "Times New Roman", 12f));

        doc.HorizontalAlignment();

        doc.AddRightIndentation();

        doc.AddSpace(1, 40f);
        doc.AddSpace(4, 98f);

        doc.CreateTable();

        doc.SaveDoc("output.docx");*/

        UpdateTable doc = new UpdateTable(new Document());
        String s = doc.ReadTable(6,1);
        String s2 = doc.ReadTable(6,2);
        String s3 = doc.MultiplyCells(s,s2);
        doc.SetCell(6,3,s3);
        doc.SaveDoc("output2.docx");
        float f = 10.5f;
        String.format("%.2f", f);
        System.out.println(f);
    }


}
