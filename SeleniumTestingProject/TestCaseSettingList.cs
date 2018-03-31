
namespace Setup
{
    public class TestCaseSettingList
    {
        private string elementType;

        private string childLink;
        private string errorId;
        private string assertValue;
        private string moduleName;
        private string btnvalue;
        private string testCaseId;
        private string imageURL;
        private string errorText;
        private int errorValueIE = 0;
        private int errorValueChrome = 0;
        private int errorValueMozilla = 0;
        private int errorValueSafari = 0;
        private bool isTestCasePassed = true;
        private string sheetName;
        public string SheetName
        {
            get
            {
                return sheetName;
            }
            set
            {
                sheetName = value;
            }
        }
        public bool IsTestCasePassed
        {
            get
            {
                return isTestCasePassed;
            }
            set
            {
                isTestCasePassed = value;
            }
        }
        public int ErrorValueSafari
        {
            get
            {
                return errorValueSafari;
            }
            set
            {
                errorValueSafari = value;
            }
        }
        public int ErrorValueMozilla
        {
            get
            {
                return errorValueMozilla;
            }
            set
            {
                errorValueMozilla = value;
            }
        }
        public int ErrorValueChrome
        {
            get
            {
                return errorValueChrome;
            }
            set
            {
                errorValueChrome = value;
            }
        }
        public int ErrorValueIE
        {
            get
            {
                return errorValueIE;
            }
            set
            {
                errorValueIE = value;
            }
        }
        public string ErrorText
        {
            get
            {
                return errorText;
            }
            set
            {
                errorText = value;
            }
        }
        public string ImageURL
        {
            get
            {
                return imageURL;
            }
            set
            {
                imageURL = value;
            }
        }
        public string TestCaseId
        {
            get
            {
                return testCaseId;
            }
            set
            {
                testCaseId = value;
            }
        }
        public string BtnValue
        {
            get
            {
                return  btnvalue;
            }
            set 
            {
                btnvalue = value;
            }
        }

        public string ElementType
        {
            get
            {
                return elementType;
            }
            set
            {
                elementType = value;
            }
        }
       
        public string ChildLink
        {
            get
            {
                return childLink;
            }
            set
            {
                childLink = value;
            }

        }
        public string ErrorId
        {
            get
            {
                return errorId;
            }
            set
            {
                errorId = value;
            }

        }
        public string AssertValue
        {
            get
            {
                return assertValue;
            }
            set
            {
                assertValue = value;
            }

        }
        public string ModuleName
        {
            get
            {
                return moduleName;
            }
            set
            {
                moduleName = value;
            }

        }
    }
}
