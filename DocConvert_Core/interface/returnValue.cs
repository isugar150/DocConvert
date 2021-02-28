namespace DocConvert_Core.interfaces
{
    internal interface returnValue
    {
        int resultCode { get; set; }
        bool isSuccess { get; set; }
        string Message { get; set; }
        int PageCount { get; set; }
    }
    public class ReturnValue : returnValue
    {
        private int _resultCode = 1000;
        private bool _isSuccess;
        private string _Message;
        private int _PageCount = -1;
        public int resultCode { get { return _resultCode; } set { _resultCode = value; } }
        public bool isSuccess { get { return _isSuccess; } set { _isSuccess = value; } }
        public string Message { get { return _Message; } set { _Message = value; } }
        public int PageCount { get { return _PageCount; } set { _PageCount = value; } }
    }
}
