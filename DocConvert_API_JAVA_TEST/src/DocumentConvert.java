import java.io.File;
import java.util.Scanner;

import DocConvert_API.*;

public class DocumentConvert {
    public static void main(String args[]) throws Exception {
        Scanner scan = new Scanner(System.in);
        System.out.print("파일 절대경로 입력: ");
        File fileAbsolutePath = new File(scan.nextLine());
        System.out.print("이미지 변환 (0:안함) (1:JPG) (2:PNG) (3:BMP): ");
        String toImg = scan.nextLine();
        if(toImg.equals(""))
            toImg = "0";
        String isSuccess = new DocConvert().DocConvert_Start(fileAbsolutePath.getParent(), fileAbsolutePath.getParent(), fileAbsolutePath.getName(), Integer.parseInt(toImg));
        System.out.println(isSuccess);
    }
}