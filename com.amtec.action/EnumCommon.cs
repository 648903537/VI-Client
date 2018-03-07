using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace com.amtec.action
{
    class EnumCommon
    {
        public enum EnumResultState
        {
            PASS = 0,
            FAIL = 1,
            SCRAP = 2,
            ERROR = -1
        }
        public enum EnumOutputMode
        {
            COMOutput = 0,
            KeyboardOutput = 1,
            NoOutput = 2
        }
        public enum EnumInputDataMode
        {
            COMInput = 0,
            SocketInput = 1,
            KeyobardInput = 2
        }
        public enum EnumYesOrNo
        {
            Yes = 1,
            No = 0
        }
        public enum EnumProcessLayer
        {
            T = 0,
            B = 1,
            //t=0,
            //b=1,
            Independent = 2
        }
    }
}
