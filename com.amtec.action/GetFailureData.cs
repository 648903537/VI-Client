using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace com.amtec.action
{
    public class GetFailureData
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private MainView view;
        private InitModel init;

        public GetFailureData(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public string[] MdataGetFailureDataforStation()
        {
            var failuredataResultKeys = new string[] { "FAILURE_TYPE_CODE", "FAILURE_TYPE_DESC" };
            var failureDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetFailureDataForStation (Station:" + init.configHandler.StationNumber);
            int error = imsapi.mdataGetFailureDataForStation(sessionContext, init.configHandler.StationNumber, 0, failuredataResultKeys, out failureDataResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mdataGetFailureDataForStation (ResultCode = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetFailureDataForStation " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdataGetFailureDataForStation " + error+"("+errorMsg+")", "");
            }
            return failureDataResultValues;
        }
    }
}
