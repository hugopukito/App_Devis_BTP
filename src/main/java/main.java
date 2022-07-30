import com.spire.doc.Document;

import java.io.File;
import java.io.FileWriter;
import java.util.concurrent.ExecutionException;

public class main {
    public static void main(String[] args) throws Exception {
        try {
            UpdateTable doc = new UpdateTable(new Document());
            doc.MainUpdateTable();
            doc.SaveDoc("output2.docx");
        } catch (Exception e) {
            System.out.println(e);
            FileWriter file_writer = new FileWriter("error.txt");
            file_writer.write(e.getMessage());
            file_writer.close();
        }
    }
}
