import DocConvert_API.*;

public class Test02 {
    public static void main(String args[]) throws InterruptedException {
        new Thread(new Runnable() {
            @Override
            public void run() {
                try {
                    new Thread(new Runnable() {
                        @Override
                        public void run() {
                            try {
                                //new DocConvert().DocConvert_Start("C:\\Users\\o0_0o\\Desktop\\Document", "C:\\Users\\o0_0o\\Desktop\\Document", "excel.xlsx", 1);
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }).start();
                    //Thread.sleep(3000);
                    new DocConvert().DocConvert_Start("C:\\Users\\o0_0o\\Desktop\\Document", "C:\\Users\\o0_0o\\Desktop\\Document", "word.docx", 1);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        }).start();
    }
}
