package appli;

import java.awt.FlowLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;

import javax.swing.*;

public class FileChooser extends JPanel implements ActionListener {

	
private static final long serialVersionUID = 1L;

	private JButton openButton;
    private JTextField path;
    private JFileChooser fc;
    private File file; 
    private JLabel label;
    private ExcelReader excelReader;

    FileChooser(String l, ExcelReader excelReader) {
        super(new FlowLayout());
        this.excelReader=excelReader;

        path = new JTextField(15);
        fc = new JFileChooser();

        label = new JLabel(l);
        openButton = new JButton("Parcourir");
        openButton.addActionListener(this);
     
        this.add(label);
        this.add(path);
        this.add(openButton);
    }
    FileChooser(String l) {
        super(new FlowLayout());

        path = new JTextField(15);
        fc = new JFileChooser();

        label = new JLabel(l);
        openButton = new JButton("Parcourir");
        openButton.addActionListener(this);

        this.add(label);
        this.add(path);
        this.add(openButton);
    }

    public JFileChooser getJFileChooser(){
        return fc;
    }
    public void actionPerformed(ActionEvent e) {

        if (e.getSource() == openButton) {

            if (fc.showOpenDialog(FileChooser.this) == JFileChooser.APPROVE_OPTION) {
                file = fc.getSelectedFile();
                path.setText(file.getPath());
                excelReader.addRadioButtons();
            }
            path.setCaretPosition(path.getDocument().getLength());
        }
    }

    File getFile() {
    	return file;
    }
    
}
