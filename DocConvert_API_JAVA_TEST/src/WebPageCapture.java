import java.io.File;
import java.util.Scanner;

import DocConvert_API.*;

public class WebPageCapture {
    public static void main(String args[]) throws Exception {
        Scanner scan = new Scanner(System.in);
        System.out.print("URL: ");
        String url = scan.nextLine();
        System.out.print("DownloadPath: ");
        String downloadPath = scan.nextLine();
        String isSuccess = new WebCapture().WebCapture_Start(url, downloadPath);
        System.out.println(isSuccess);
    }
}
