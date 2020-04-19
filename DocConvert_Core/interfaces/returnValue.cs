using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocConvert_Core.interfaces
{
    interface returnValue
    {
        bool isSuccess { get; set; }
        string Message { get; set; }
        int PageCount { get; set; }
    }
    public class ReturnValue : returnValue
    {
        private bool _isSuccess;
        private string _Message;
        private int _PageCount = -1;
        public bool isSuccess { get { return _isSuccess; } set { _isSuccess = value; } }
        public string Message { get { return _Message; } set { _Message = value; } }
        public int PageCount { get { return _PageCount; } set { _PageCount = value; } }
    }
}
