using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class CheckUserSkill
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public CheckUserSkill(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int CheckUserSkillForWS(string userName, string workorder, int processLayer)
        {
            int errorCode = 0;
            KeyValue[] checkUserSkillFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder) };
            LogHelper.Info("begin api trCheckUserSkill (Station number:" + init.configHandler.StationNumber + ",workorder:" + workorder + ",process layer:" + processLayer + ",user name:" + userName + ")");
            errorCode = imsapi.trCheckUserSkill(sessionContext, init.configHandler.StationNumber, processLayer, userName, checkUserSkillFilter);
            string errorString = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api trCheckUserSkill (result code = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trCheckUserSkill " + errorCode+"("+errorString+")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckUserSkill " + errorCode, "");
            }
            return errorCode;
        }

        public int CheckUserSkillForSN(string userName, string serialNumber, int processLayer)
        {
            int errorCode = 0;
            KeyValue[] checkUserSkillFilter = new KeyValue[] { new KeyValue("SERIAL_NUMBER", serialNumber) };
            LogHelper.Info("begin api trCheckUserSkill (Station number:" + init.configHandler.StationNumber + ",serial number:" + serialNumber + ",process layer:" + processLayer + ",user name:" + userName + ")");
            errorCode = imsapi.trCheckUserSkill(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, userName, checkUserSkillFilter);
            string errorString = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api trCheckUserSkill (result code = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trCheckUserSkill " + errorCode+"("+errorString+")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckUserSkill " + errorCode, "");
            }
            return errorCode;
        }
    }
}
