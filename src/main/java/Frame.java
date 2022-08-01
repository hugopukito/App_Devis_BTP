import javax.swing.*;
import java.awt.*;
import java.awt.dnd.DropTarget;

public class Frame extends JFrame {

    public Frame() {

        // Set the frame title
        super("App devis");

        // Set the size
        setSize(1920/4, 1080/4);
        getContentPane().setBackground(new Color(33,33,33));

        // Create the drag and drop listener
        DragDrop dragDrop = new DragDrop(this);

        // Connect the frame with a drag and drop listener
        new DropTarget(this, dragDrop);

        setLocationRelativeTo(null);

        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

        // Show the frame
        setVisible(true);

    }
}
