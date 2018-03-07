using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using com.amtec.action;
using com.amtec.model;
using com.amtec.forms;

namespace com.amtec.action
{
    public class MergeManger
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public MergeManger(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int VerifyMerge(string serialNumberSlave, string serialNumberMaster)
        {
            SerialNumberStateData[] serialNumberStateDataArray = new SerialNumberStateData[] { };
            int error = imsapi.trVerifyMerge(sessionContext, init.configHandler.StationNumber, serialNumberSlave, "-1", serialNumberMaster, "-1", 0, out serialNumberStateDataArray);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("API trVerifyMerge (serial number master = " + serialNumberMaster + ", serial number slave = " + serialNumberSlave + ",error =" + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, " trVerifyMerge " + error, "");
            }
            else
            {
                view.errorHandler(2, " trVerifyMerge " + error+"("+errorMsg+")", "");
            }
            return error;
        }

        //trMergeParts
        public int MergeSerialNumber(string serialNumberMaster, string serialNumberSalve, int processLayer)
        {
            int error = 0;
            error = imsapi.trMergeParts(sessionContext, init.configHandler.StationNumber, processLayer, 0, serialNumberMaster, "-1", serialNumberSalve, "-1");
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("API trMergeParts (serial number master = " + serialNumberMaster + ", serial number slave = " + serialNumberSalve + ",error =" + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, " trMergeParts " + error, "");
            }
            else
            {
                view.errorHandler(2, " trMergeParts " + error+"("+errorMsg+")", "");
            }
            return error;
        }

        public int MergeSerialNumber(string serialNumber, int processLayer)
        {
            SwitchSerialNumberData snData = new SwitchSerialNumberData(0, serialNumber, "1", serialNumber + "A", 0);
            SwitchSerialNumberData[] serialNumberArray = new SwitchSerialNumberData[] { snData };
            int error = imsapi.trAssignSerialNumberMergeAndUploadState(sessionContext, init.configHandler.StationNumber, processLayer, serialNumber + "A", "1", new SerialNumberData[] { }, serialNumber, 0, -1, 0);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0)
            {
                view.errorHandler(0, "API trAssignSerialNumberMergeAndUploadState " + error, "");
                //switch serial number
                int error1 = imsapi.trSwitchSerialNumber(sessionContext, init.configHandler.StationNumber, "-1", "-1", ref serialNumberArray);
                string errorMsg1 = UtilityFunction.GetZHSErrorString(error1, init, sessionContext);
                if (error1 == 0)
                {
                    view.errorHandler(0, " trSwitchSerialNumber " + error1, "");
                }
                else
                {
                    view.errorHandler(2, " trSwitchSerialNumber " + error1+"("+errorMsg1+")", "");
                    return error1;
                }
            }
            else
            {
                view.errorHandler(2, " trAssignSerialNumberMergeAndUploadState " + error+"("+errorMsg+")", "");
            }
            return error;
        }

        public int SwitchSerialNumber(string serialNumber, int processLayer)
        {
            SwitchSerialNumberData snData = new SwitchSerialNumberData(0, serialNumber + "_1", "1", serialNumber, 0);
            SwitchSerialNumberData[] serialNumberArray = new SwitchSerialNumberData[] { snData };

            int error = imsapi.trSwitchSerialNumber(sessionContext, init.configHandler.StationNumber, "-1", "-1", ref serialNumberArray);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0)
            {
                view.errorHandler(0, " trSwitchSerialNumber " + error, "");
            }
            else
            {
                view.errorHandler(2, " trSwitchSerialNumber " + error+"("+errorMsg+")", "");
            }
            return error;
        }
    }
}
