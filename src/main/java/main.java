import com.spire.doc.Document;

import java.text.SimpleDateFormat;
import java.util.Date;

public class main {
    public static void main (String[] args){

        Doc doc = new Doc(new Document());

        doc.SetTopMargin();

        doc.CreateParagraph();

        doc.ApplyStyle(doc.CreateStyle("style1", "Times New Roman", 12f));

        doc.HorizontalAlignment();

        doc.AddRightIndentation();

        doc.AddSpace(1, 40f);

        doc.SaveDoc("output.docx");
    }
}
