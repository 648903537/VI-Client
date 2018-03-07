using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class GetSerialNumberInfo
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetSerialNumberInfo(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public string[] GetSNInfo(string serialNumber)
        {
            int error = 0;
            string errorMsg = "";
            string[] serialNumberResultKeys = new string[] { "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER", "ATTRIBUTE_1" };
            string[] serialNumberResultValues = new string[] { };
            LogHelper.Info("begin api trGetSerialNumberInfo (serial number =" + serialNumber + ")");
            error = imsapi.trGetSerialNumberInfo(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", serialNumberResultKeys, out serialNumberResultValues);
            LogHelper.Info("end api trGetSerialNumberInfo (result code = " + error + ")");
            //imsapi.imsapiGetErrorText(sessionContext, error, out errorMsg);
            errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0 || error == -203 || error == 209)
            {
                LogHelper.Info(init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error);
                error = 0;
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberInfo " + error + "," + errorMsg, "");
            }
            return serialNumberResultValues;
        }

        public SerialNumberData[] GetSerialNumberBySNRef(string serialNumberRef)
        {
            int error = 0;
            string errorMsg = "";
            SerialNumberData[] serialNumberArray = new SerialNumberData[] { };
            error = imsapi.trGetSerialNumberBySerialNumberRef(sessionContext, init.configHandler.StationNumber, serialNumberRef, "-1", out serialNumberArray);
            //imsapi.imsapiGetErrorText(sessionContext, error, out errorMsg);
            errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("api trGetSerialNumberBySerialNumberRef (serial number ref = " + serialNumberRef + ",result code = " + error);
            if (error == 0 || error == -203 || error == 209)
            {
                LogHelper.Info(init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberBySerialNumberRef " + error);
                error = 0;
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberBySerialNumberRef " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trGetSerialNumberBySerialNumberRef " + error + "," + errorMsg, "");
            }
            return serialNumberArray;
        }

    }
}
