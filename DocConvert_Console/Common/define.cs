using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Console
{
    public static class define
    {
        public static int OK = 0; // 정상

        public static int UNDEFINE_ERROR = 1000; // 알수없는 오류
        public static int PARSING_INI_ERROR = 1001; // INI파일 파싱 오류
        public static int INVALID_METHOD_ERROR = 1002; // 잘못된 변환 메소드
        public static int PDF_TO_PDF_ERROR = 1003; // PDF파일을 PDF로 변환 요청했을때
        public static int INVALID_IMAGE_CONVERT_REQUEST_ERROR = 1004; // PDF파일을 PDF로 변환 요청했을때
    }
}
