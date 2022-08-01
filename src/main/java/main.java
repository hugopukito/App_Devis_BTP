import com.spire.doc.Document;

import java.io.FileWriter;

public class main {
    public static void main(String[] args) throws Exception {
        try {
            new Frame();
            UpdateTable doc = new UpdateTable(new Document());
            doc.MainUpdateTable();
            doc.SaveDoc("output2.docx");
            new FileWriter("C:\\Users\\pukit\\Desktop\\test.txt");
        } catch (Exception e) {
            FileWriter file_writer = new FileWriter("error.txt");
            file_writer.write(e.getMessage());
            file_writer.close();
        }
    }
}


