package docconvert;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;

import org.apache.commons.net.PrintCommandListener;
import org.apache.commons.net.ftp.FTP;
import org.apache.commons.net.ftp.FTPClient;
import org.apache.commons.net.ftp.FTPReply;

public class FTPUploader {

    FTPClient ftp = null;

    //param( host server ip, username, password )
    public FTPUploader(String host, int port, String user, String pwd) throws Exception {
        ftp = new FTPClient();
        ftp.addProtocolCommandListener(new PrintCommandListener(new PrintWriter(System.out)));
        int reply;
        ftp.connect(host, port);//호스트 연결
        reply = ftp.getReplyCode();
        if (!FTPReply.isPositiveCompletion(reply)) {
            ftp.disconnect();
            throw new Exception("Exception in connecting to FTP Server");
        }
        ftp.login(user, pwd);//로그인
        ftp.setFileType(FTP.BINARY_FILE_TYPE);
        ftp.enterLocalPassiveMode();
    }

    //param( 보낼파일경로+파일명, 호스트에서 받을 파일 이름, 호스트 디렉토리 )
    public void uploadFile(String localFileFullName, String fileName, String hostDir) throws Exception {
        System.out.println("[DocConvert] 파일 업로드: " + localFileFullName);
        try (InputStream input = new FileInputStream(new File(localFileFullName))) {
            this.ftp.storeFile(hostDir + fileName, input);
            System.out.println("[DocConvert] 파일 업로드 성공!");
            //storeFile() 메소드가 전송하는 메소드
        }
    }

    public void downloadFiles() {
        /*string root = "d:\\ftptest\\download";
        // FTP에서 파일 리스트와 디렉토리 정보를 취득한다.
        if (getFileList(client, File.separator, files, directories)) {
            // 디렉토리 구조대로 로컬 디렉토리 생성
            for (String directory : directories) {
                File file = new File(root + directory);
                file.mkdir();
            }
            for (String file : files) {
                // 파일의 OutputStream을 가져온다.
                try (FileOutputStream fo = new FileOutputStream(root + File.separator + file)) {
                    // FTPClient의 retrieveFile함수로 보내면 다운로드가 이루어 진다.
                    if (client.retrieveFile(file, fo)) {
                        System.out.println("Download - " + file);
                    }
                }
            }*/
        }

        public void disconnect () {
            if (this.ftp.isConnected()) {
                try {
                    this.ftp.logout();
                    this.ftp.disconnect();
                } catch (IOException f) {
                    f.printStackTrace();
                }
            }
        }
    }