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

public class Test01 {
    public static void main(String args[]) throws Exception {
        new DocConvert("C:\\Users\\JMKIM\\Desktop", "C:\\Users\\JMKIM\\Desktop", "test.pptx", 0);
    }
}