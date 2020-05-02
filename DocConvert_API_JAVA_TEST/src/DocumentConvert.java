import java.io.File;
import java.util.Scanner;

import DocConvert_API.*;

public class DocumentConvert {
    public static void main(String args[]) throws Exception {
        Scanner scan = new Scanner(System.in);
        System.out.println("DocConvert(1), WebCapture(2)");
        int getMethod = Integer.parseInt(scan.nextLine());
        if(getMethod == 1) {
            System.out.print("파일 절대경로 입력: ");
            File fileAbsolutePath = new File(scan.nextLine());
            System.out.print("이미지 변환 (0:안함) (1:JPG) (2:PNG) (3:BMP): ");
            String toImg = scan.nextLine();
            if (toImg.equals(""))
                toImg = "0";
            String isSuccess = new DocConvert().DocConvert_Start(fileAbsolutePath.getParent(), fileAbsolutePath.getParent(), fileAbsolutePath.getName(), Integer.parseInt(toImg));
            System.out.println(isSuccess);
        }else if(getMethod == 2){
            System.out.print("URL: ");
            String url = scan.nextLine();
            System.out.print("DownloadPath: ");
            String downloadPath = scan.nextLine();
            String isSuccess = new WebCapture().WebCapture_Start(url, downloadPath);
            System.out.println(isSuccess);
        }
    }
}