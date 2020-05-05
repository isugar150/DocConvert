package org.namejm.dc;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Properties;

public class dcProperties {
    private String workspacePath;

    public String workspacePath() { return workspacePath; }

    public boolean readProperties() {
        ClassLoader cl;
        cl = Thread.currentThread().getContextClassLoader();

        if (cl == null) {
            cl = ClassLoader.getSystemClassLoader();
        }
        URL url = cl.getResource("dc.properties");

        File propFile = null;
        if(new File("./dc.properties").exists()){
            propFile = new File("./dc.properties");
        } else{
            propFile = new File(url.getPath());
        }
        FileInputStream is;

        try {
            is = new FileInputStream(propFile);
            Properties props = new Properties();
            props.load(is);

            this.workspacePath = props.getProperty("workspacePath").trim();
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }
}
