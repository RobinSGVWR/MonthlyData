//
// Source code recreated from a .class file by IntelliJ IDEA
// (powered by Fernflower decompiler)
//

package appli;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Enumeration;
import java.util.Iterator;
import javax.swing.AbstractButton;
import javax.swing.BoxLayout;
import javax.swing.ButtonGroup;
import javax.swing.JButton;
import javax.swing.JEditorPane;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JProgressBar;
import javax.swing.JRadioButton;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader extends JPanel implements ActionListener {
    private FileChooser c;
    private JButton startButton;
    private JFrame container;
    private ButtonGroup group = new ButtonGroup();
    private JEditorPane label = new JEditorPane("text/html", "");
    private JProgressBar bar;
    private Thread t;
    private JTextField txt;
    private JTextArea console;

    ExcelReader(JFrame container) {
        this.container = container;
        this.c = new FileChooser("Choisir un fichier Excel compatible", this);
        this.label.setText(" <b> Entrez le titre de la nouvelle feuille à creer </b>");
        this.txt = new JTextField("new data sheet");
        this.startButton = new JButton("valider");
        this.startButton.addActionListener(this);
        this.console = new JTextArea("Console");
        this.setLayout(new BoxLayout(this, BoxLayout.Y_AXIS));
        this.add(this.c);
        this.add(this.label);
        this.add(this.txt);
        this.add(this.console);

        this.bar = new JProgressBar();
        this.bar.setMaximum(100);
        this.bar.setMinimum(0);
        this.bar.setStringPainted(true);
        this.add(this.bar, "Center");

    }

    public void actionPerformed(ActionEvent e) {
        try {
            if (e.getSource() == this.startButton) {
                FileInputStream file = new FileInputStream(this.c.getFile());
                String sheetName;
                ArrayList sheetL;
                Iterator rowIt;
                ArrayList rowL;
                Row row;
                Iterator cellIt;
                if (FilenameUtils.getExtension(this.c.getFile().getPath().toUpperCase()).compareToIgnoreCase("XLS") == 0) {
                    HSSFWorkbook wb = new HSSFWorkbook(file);
                    this.addToConsole("event BUTTON triggered");
                    System.out.println("event BUTTON triggered");
                    sheetName = this.getSelectedButtonLabel();
                    if (sheetName != null) {
                        HSSFSheet sheet = this.getSheetByName(wb, sheetName);
                        sheetL = new ArrayList();
                        if (sheet != null) {
                            rowIt = sheet.iterator();

                            while(rowIt.hasNext()) {
                                rowL = new ArrayList();
                                row = (Row)rowIt.next();
                                cellIt = row.cellIterator();

                                while(cellIt.hasNext()) {
                                    HSSFCell cell = (HSSFCell)cellIt.next();
                                    rowL.add(this.cellToString(cell, cell.getCellTypeEnum()));
                                }

                                sheetL.add(rowL);
                            }


                            System.out.println(this.txt.getText());
                        }

                        ExcelWriter excelW = new ExcelWriter(this.txt.getText(), wb, sheetL, this);
                        excelW.generateExcel();
                        bar.setValue(100);
                    }
                } else if (FilenameUtils.getExtension(this.c.getFile().getPath().toUpperCase()).compareToIgnoreCase("XLSX") == 0) {
                    XSSFWorkbook wb = new XSSFWorkbook(file);
                    System.out.println("event BUTTON triggered");
                    this.addToConsole("event BUTTON triggered");
                    sheetName = this.getSelectedButtonLabel();
                    if (sheetName != null) {
                        XSSFSheet sheet = this.getSheetByName(wb, sheetName);
                        sheetL = new ArrayList();
                        if (sheet != null) {
                            rowIt = sheet.iterator();

                            while(rowIt.hasNext()) {
                                rowL = new ArrayList();
                                row = (Row)rowIt.next();
                                cellIt = row.cellIterator();

                                while(cellIt.hasNext()) {
                                    XSSFCell cell = (XSSFCell)cellIt.next();
                                    rowL.add(this.cellToString(cell, cell.getCellTypeEnum()));
                                }

                                sheetL.add(rowL);
                            }
                            bar.setValue(20);

                            System.out.println(this.txt.getText());
                            this.addToConsole(this.txt.getText());
                        }

                        ExcelWriterXSSF excelW = new ExcelWriterXSSF(this.txt.getText(), wb, sheetL, this);
                        excelW.generateExcel();
                        bar.setValue(100);
                    }
                }
            }

            this.deleteRadioButtons();
        } catch (IOException var13) {
            var13.printStackTrace();
            this.addToConsole(var13.toString());
        }

    }

    private String getSelectedButtonLabel() {
        Enumeration buttons = this.group.getElements();

        while(buttons.hasMoreElements()) {
            AbstractButton button = (AbstractButton)buttons.nextElement();
            if (button.isSelected()) {
                return button.getText();
            }
        }

        return null;
    }

    public void addRadioButtons() {
        System.out.println("event File triggered");
        this.addToConsole("event File triggered");

        try {
            FileInputStream file = new FileInputStream(this.c.getFile());
            System.out.println(this.c.getFile().getPath().toUpperCase());
            System.out.println(FilenameUtils.getExtension(this.c.getFile().getPath().toUpperCase()));
            this.addToConsole(this.c.getFile().getPath().toUpperCase());
            Iterator var3;
            Object aWb;
            JRadioButton btn;
            if (FilenameUtils.getExtension(this.c.getFile().getPath().toUpperCase()).compareToIgnoreCase("XLS") == 0) {
                HSSFWorkbook wb = new HSSFWorkbook(file);
                var3 = wb.iterator();

                while(var3.hasNext()) {
                    aWb = var3.next();
                    HSSFSheet sheet = (HSSFSheet)aWb;
                    btn = new JRadioButton(sheet.getSheetName());
                    this.add(btn);
                    this.group.add(btn);
                }

                this.group.setSelected(this.group.getElements().nextElement().getModel(), true);
                this.add(this.startButton);
                this.add(this.txt);
                this.container.pack();
                this.enableStart();
            } else if (FilenameUtils.getExtension(this.c.getFile().getPath().toUpperCase()).compareToIgnoreCase("XLSX") == 0) {
                XSSFWorkbook wb = new XSSFWorkbook(file);
                System.out.println("Woorbook created");
                this.addToConsole("Woorbook created");
                var3 = wb.iterator();

                while(var3.hasNext()) {
                    aWb = var3.next();
                    XSSFSheet sheet = (XSSFSheet)aWb;
                    btn = new JRadioButton(sheet.getSheetName());
                    this.add(btn);
                    this.group.add(btn);
                }

                this.group.setSelected(this.group.getElements().nextElement().getModel(), true);
                this.add(this.startButton);
                this.add(this.txt);
                this.container.pack();
                this.enableStart();
            }
        } catch (FileNotFoundException var7) {
//            System.out.println(var7);
            this.addToConsole(var7.toString());
        } catch (IOException var8) {
            var8.printStackTrace();
            this.addToConsole(var8.toString());
        }

    }

    private void deleteRadioButtons() {
    }

    private void enableStart() {
        this.startButton.setEnabled(true);
    }

    private HSSFSheet getSheetByName(HSSFWorkbook workbook, String sheetName) {

        for (Object aWorkbook : workbook) {
            HSSFSheet sheet = (HSSFSheet) aWorkbook;
            if (sheet.getSheetName().equals(sheetName)) {
                return sheet;
            }
        }

        return null;
    }

    private XSSFSheet getSheetByName(XSSFWorkbook workbook, String sheetName) {

        for (Object aWorkbook : workbook) {
            XSSFSheet sheet = (XSSFSheet) aWorkbook;
            if (sheet.getSheetName().equals(sheetName)) {
                return sheet;
            }
        }

        return null;
    }

    private String cellToString(HSSFCell cell, CellType type) {
        switch(type) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
        case ERROR:
            return String.valueOf(cell.getErrorCellValue());
        case BLANK:
            return "[x]";
        case FORMULA:
            return this.cellToString(cell, cell.getCachedFormulaResultTypeEnum());
        default:
            return "----------------------------------------- " + type.toString();
        }
    }

    private String cellToString(XSSFCell cell, CellType type) {
        switch(type) {
        case STRING:
            return cell.getStringCellValue();
        case NUMERIC:
            return String.valueOf(cell.getNumericCellValue());
        case ERROR:
            return String.valueOf(cell.getErrorCellValue());
        case BLANK:
            return "[x]";
        case FORMULA:
            return this.cellToString(cell, cell.getCachedFormulaResultTypeEnum());
        default:
            return "----------------------------------------- " + type.toString();
        }
    }

    void addToConsole(String txt) {
        this.console.append(" - " + txt + "\r\n");
        this.container.pack();
        this.repaint();
    }

    public FileChooser getC() {
        return c;
    }

    public void setC(FileChooser c) {
        this.c = c;
    }

    public JButton getStartButton() {
        return startButton;
    }

    public void setStartButton(JButton startButton) {
        this.startButton = startButton;
    }

    public JFrame getContainer() {
        return container;
    }

    public void setContainer(JFrame container) {
        this.container = container;
    }

    public ButtonGroup getGroup() {
        return group;
    }

    public void setGroup(ButtonGroup group) {
        this.group = group;
    }

    public JEditorPane getLabel() {
        return label;
    }

    public void setLabel(JEditorPane label) {
        this.label = label;
    }

    JProgressBar getBar() {
        return bar;
    }

    public void setBar(JProgressBar bar) {
        this.bar = bar;
    }

    public Thread getT() {
        return t;
    }

    public void setT(Thread t) {
        this.t = t;
    }

    public JTextField getTxt() {
        return txt;
    }

    public void setTxt(JTextField txt) {
        this.txt = txt;
    }

    public JTextArea getConsole() {
        return console;
    }

    public void setConsole(JTextArea console) {
        this.console = console;
    }


}
