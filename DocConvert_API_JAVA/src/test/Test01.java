package test;

import docconvert.DocConvert;
import docconvert.getProperties;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Properties;
import java.util.Scanner;

public class Test01 {
    public static void main(String args[]) throws Exception {
        Scanner scan = new Scanner(System.in);
        System.out.print("파일 절대경로 입력: ");
        File fileAbsolutePath = new File(scan.nextLine());
        System.out.print("이미지 변환 (0:안함) (1:JPG) (2:PNG) (3:BMP): ");
        String toImg = scan.nextLine();
        String isSuccess = new DocConvert().DocConvert_Start(fileAbsolutePath.getParent(), fileAbsolutePath.getParent(), fileAbsolutePath.getName(), Integer.parseInt(toImg));
        System.out.println(isSuccess);
    }
}