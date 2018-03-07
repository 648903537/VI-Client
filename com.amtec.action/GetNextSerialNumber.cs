using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class GetNextSerialNumber
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetNextSerialNumber(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public SerialNumberData[] GetSerialNumber(string Temp_PartNo, int numberRecords)
        {
            SerialNumberData[] serialNumberArray = new SerialNumberData[] { };
            int error = imsapi.trGetNextSerialNumber(sessionContext, init.configHandler.StationNumber, "-1", Temp_PartNo, numberRecords, out serialNumberArray);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("API trGetNextSerialNumber:partnumber" + Temp_PartNo + ",ERROR" + error);
            return serialNumberArray;
        }
    }
}
