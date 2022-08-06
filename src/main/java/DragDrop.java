import com.spire.doc.Document;

import javax.swing.*;
import java.awt.*;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.dnd.*;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.List;

class DragDrop implements DropTargetListener {

    JFrame frame;
    JLabel label;
    Color myGreen = new Color(3, 153, 53);
    Color myWhite = new Color(177, 206, 252);

    public DragDrop(Frame _frame) {
        frame = _frame;
        frame.getRootPane().setBorder(BorderFactory.createMatteBorder(4,3,3,3, myWhite));
        label = new JLabel("Déposer fichier", SwingConstants.CENTER);
        label.setForeground(myWhite);
        label.setFont(new Font("Calibri", Font.BOLD, 30));
        frame.add(label);
    }

    @Override
    public void drop(DropTargetDropEvent event) {
        // Accept copy drops
        event.acceptDrop(DnDConstants.ACTION_COPY);
        // Get the transfer which can provide the dropped item data
        Transferable transferable = event.getTransferable();
        // Get the data formats of the dropped item
        DataFlavor[] flavors = transferable.getTransferDataFlavors();

        // Loop through the flavors
        for (DataFlavor flavor : flavors) {
            try {
                // If the drop items are files
                if (flavor.isFlavorJavaFileListType()) {
                    // Get all of the dropped files
                    @SuppressWarnings("rawtypes")
                    List files = (List) transferable.getTransferData(flavor);
                    // Loop them through
                    for (Object item : files) {
                        File file = (File) item;
                        boolean failed = false;
                        try {
                            UpdateTable doc = new UpdateTable(new Document(), file.getPath());
                            doc.MainUpdateTable();

                            JFileChooser f = new JFileChooser(file.getParent());
                            f.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                            f.showSaveDialog(null);

                            doc.SaveDoc(f.getSelectedFile() + "\\new_" + file.getName());
                        } catch (Exception e) {
                            failed = true;
                            frame.getRootPane().setBorder(BorderFactory.createMatteBorder(4,3,3,3, Color.red));
                            frame.setSize(1920,1080/4);
                            frame.setLocationRelativeTo(null);
                            label.setForeground(Color.red);
                            label.setText(e.getMessage());
                        } finally {
                            if (!failed) {
                                label.setForeground(myGreen);
                                label.setText("Finis !");
                            }
                        }
                    }
                }
            } catch (Exception e) {
                // Print out the error stack
                e.printStackTrace();
            }
        }
        // Inform that the drop is complete
        event.dropComplete(true);
    }

    @Override
    public void dragEnter(DropTargetDragEvent event) {
        label.setForeground(myGreen);
        frame.getRootPane().setBorder(BorderFactory.createMatteBorder(4,3,3,3, myGreen));
    }

    @Override
    public void dragExit(DropTargetEvent event) {
        label.setForeground(myWhite);
        frame.getRootPane().setBorder(BorderFactory.createMatteBorder(4,3,3,3, myWhite));
    }

    @Override
    public void dragOver(DropTargetDragEvent event) {
    }

    @Override
    public void dropActionChanged(DropTargetDragEvent event) {
    }

}