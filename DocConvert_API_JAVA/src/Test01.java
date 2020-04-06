import docconvert.DocConvert;

public class Test01 {
    public static void main(String args[]) throws Exception {
        try {
            new DocConvert("C:\\Users\\JMKIM\\Desktop\\", "C:\\Users\\JMKIM\\Desktop\\", "word.docx", 0);
        } catch(Exception e){
            System.out.println("[Error] " + e.getMessage());
        }
    }
}
