using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;

namespace com.amtec.action
{
    public class UploadProcessResult
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private int error;
        private MainView view;

        public UploadProcessResult(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int UploadProcessResultCall(String[] serialNumberArray, int processLayer)
        {
            String[] serialNumberUploadKey = new String[] { "ERROR_CODE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberUploadValues = new String[] { };
            String[] serialNumberResultValues = new String[] { };
            serialNumberUploadValues = serialNumberArray;
            error = imsapi.trUploadState(sessionContext, init.configHandler.StationNumber, 2, "-1", "-1", 0, 1, -1, 0, serialNumberUploadKey, serialNumberUploadValues, out serialNumberResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if ((error != 0) && (error != 210))
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error+"("+errorMsg+")", "");
                return error;
            }
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error, "");
            return error;
        }

        public int UploadProcessResultCall(string serialNumber)
        {
            String[] serialNumberUploadKey = new String[] { "ERROR_CODE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberUploadValues = new String[] { };
            String[] serialNumberResultValues = new String[] { };
            error = imsapi.trUploadState(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, serialNumber, "1", 0, 1, -1, 0, serialNumberUploadKey, serialNumberUploadValues, out serialNumberResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("Api trUploadState: serial number =" + serialNumber + ", result code =" + error);
            if ((error != 0) && (error != 210))
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error+"("+errorMsg+")", "");
                return error;
            }
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error, "");
            return error;
        }

        public int UploadProcessResultCall(string serialNumber, int processLayer, int snstate)
        {
            String[] serialNumberUploadKey = new String[] { "ERROR_CODE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberUploadValues = new String[] { };
            String[] serialNumberResultValues = new String[] { };
            error = imsapi.trUploadState(sessionContext, init.configHandler.StationNumber, processLayer, serialNumber, "-1", snstate, 1, -1, 0, serialNumberUploadKey, serialNumberUploadValues, out serialNumberResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("Api trUploadState: serial number =" + serialNumber + ", result code =" + error);
            if ((error != 0) && (error != 210))
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error+"("+errorMsg+")", "");
                return error;
            }
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadState " + error, "");
            return error;
        }

        public int UploadFailureAndResultData(string serialNumber, string serialNumberPos, int processLayer, int serialNumberState, int duplicateSerialNumber
            , string[] measureValues, string[] failureValues)
        {
            string[] measureKeys = new string[] { "ERROR_CODE", "MEASURE_FAIL_CODE", "MEASURE_NAME", "MEASURE_VALUE" };
            string[] failureKeys = new string[] { "ERROR_CODE", "FAILURE_TYPE_CODE" };
            string[] failureSlipKeys = new string[] { "ERROR_CODE", "TEST_STEP_NAME" };
            string[] failureSlipValues = new string[] { };
            string[] measureResultValues = new string[] { };
            string[] failureResultValues = new string[] { };
            string[] failureSlipResultValues = new string[] { };

            error = imsapi.trUploadFailureAndResultData(sessionContext, init.configHandler.StationNumber, processLayer, serialNumber, serialNumberPos,
                serialNumberState, duplicateSerialNumber, 0, -1, measureKeys, measureValues, out measureResultValues, failureKeys, failureValues, out failureResultValues,
                failureSlipKeys, failureSlipValues, out failureSlipResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("Api trUploadFailureAndResultData (serial number:" + serialNumber + ",error code:" + error);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadFailureAndResultData " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadFailureAndResultData " + error+"("+errorMsg+")", "");
            }
            return error;
        }

    }
}
