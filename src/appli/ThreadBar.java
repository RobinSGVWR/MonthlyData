package appli;

public class ThreadBar implements Runnable {
    private ExcelReader excelReader;
    private int value;

    public ThreadBar(ExcelReader excelReader){
        this.excelReader=excelReader;
    }
    public void run(int Value) {


            excelReader.getBar().setValue(value);
            excelReader.repaint();
            excelReader.revalidate();
        try {
            Thread.sleep(value);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }


    }

    @Override
    public void run() {

    }
}
