package DocConvert_API;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Properties;

public class getProperties {
    private String serverIP;
    private int serverPORT;
    private int ftpPORT;
    private String ftpUSER;
    private String ftpPASS;
    private boolean OnlyImgDownload;

    public String getServerIP() {
        return serverIP;
    }

    public int getServerPORT() {
        return serverPORT;
    }

    public int getFtpPORT() { return ftpPORT; }

    public String getFtpUSER() { return ftpUSER; }

    public String getFtpPASS() { return ftpPASS; }

    public boolean getOnlyImgDownload() { return OnlyImgDownload; }

    public boolean readProperties() {
        ClassLoader cl;
        cl = Thread.currentThread().getContextClassLoader();

        if (cl == null) {
            cl = ClassLoader.getSystemClassLoader();
        }
        URL url = cl.getResource("setting.properties");

        File propFile = null;
        if(new File("./setting.properties").exists()){
            propFile = new File("./setting.properties");
        } else{
            propFile = new File(url.getPath());
        }
        FileInputStream is;

        try {
            is = new FileInputStream(propFile);
            Properties props = new Properties();
            props.load(is);

            this.serverIP = props.getProperty("serverIP").trim();
            this.serverPORT = Integer.parseInt(props.getProperty("serverPORT").trim());
            this.ftpPORT = Integer.parseInt(props.getProperty("ftpPORT").trim());
            this.ftpUSER = props.getProperty("ftpUSER").trim();
            this.ftpPASS = props.getProperty("ftpPASS").trim();
            this.OnlyImgDownload = Boolean.getBoolean("OnlyImgDownload");
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }
}
