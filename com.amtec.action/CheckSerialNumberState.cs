﻿using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;

namespace com.amtec.action
{
    public class CheckSerialNumberState
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public CheckSerialNumberState(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public List<SerialNumberStateEntity> GetSerialNumberData(string serialNumber, int processLayer)
        {
            List<SerialNumberStateEntity> snList = new List<SerialNumberStateEntity>();
            String[] serialNumberStateResultKeys = new String[] { "LOCK_STATE", "SERIAL_NUMBER", "SERIAL_NUMBER_POS", "SERIAL_NUMBER_STATE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            LogHelper.Info("end api trCheckSerialNumberState (result code = " + error + ")");
            if ((error != 0) && (error != 5) && (error != 6) && (error != 204) && (error != 207) && (error != 212))
            {
                string errorString = "";
                //imsapi.imsapiGetErrorText(sessionContext, error, out errorString);
                errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "," + errorString, "");
            }
            else
            {
                int counter = serialNumberStateResultValues.Length;
                int loop = serialNumberStateResultKeys.Length;
                for (int i = 0; i < counter; i += loop)
                {
                    snList.Add(new SerialNumberStateEntity(serialNumberStateResultValues[i], serialNumberStateResultValues[i + 1], serialNumberStateResultValues[i + 2], serialNumberStateResultValues[i + 3]));
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
            }
            return snList;
        }

        public bool CheckSNState(string serialNumber)
        {
            String[] serialNumberStateResultKeys = new String[] { "ERROR_CODE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trCheckSerialNumberState (errorcode = " + error + ")");
            if ((error != 0) && (error != 5) && (error != 6) && (error != 204) && (error != 207) && (error != 212))
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "(" + errorString + ")", "");
                return false;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
                if (error == 5)//202 Serial no. is invalid for this station; it was not seen by the previous station
                {
                    foreach (var item in serialNumberStateResultValues)
                    {
                        if (item == "0")
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
            return true;
        }

        #region add by qy
        public int trCheckSNState(string serialNumber, int processLayer)
        {
            String[] serialNumberStateResultKeys = new String[] { "ERROR_CODE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            string errorString = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trCheckSerialNumberState (errorcode = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "(" + errorString + ")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
            }
            return error;
        }

        public int trCheckSNStateNextStep(string serialNumber, int processLayer)
        {
            String[] serialNumberStateResultKeys = new String[] { "ERROR_CODE" };
            String[] serialNumberStateResultValues = new String[] { };
            LogHelper.Info("begin api trCheckSerialNumberState (Serial number:" + serialNumber + ")");
            int error = imsapi.trCheckSerialNumberState(sessionContext, init.configHandler.StationNumber, processLayer, 1, serialNumber, "-1", serialNumberStateResultKeys, out serialNumberStateResultValues);
            //string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trCheckSerialNumberState (errorcode = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
            }
            else if (error == 5 || error == 6)
            {
                string errorMsg = "";
                int looplength = serialNumberStateResultKeys.Length;
                int alllength = serialNumberStateResultValues.Length;
                for (int i = 0; i < alllength; i += looplength)
                {
                    int errorcode = Convert.ToInt32(serialNumberStateResultValues[i]);
                    if (errorcode != 0)
                    {
                        if (errorcode == 202 || errorcode == 203)
                        {
                            string workstepdesc = GetNextProductionStep(serialNumber);
                            string errorString = UtilityFunction.GetZHSErrorString(errorcode, init, sessionContext);
                            errorMsg = errorcode + ";" + errorString + "(" + workstepdesc + ")";
                        }
                        else if (errorcode == -201 || errorcode == 204 || errorcode == 207 || errorcode == 212)//scrap
                        {
                        }
                        else
                        {
                            string errorString = UtilityFunction.GetZHSErrorString(errorcode, init, sessionContext);
                            errorMsg = errorcode + ";" + errorString;
                        }
                    }
                }
                if (errorMsg != "")
                {
                    view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error + "(" + errorMsg + ")", "");
                }
                else
                {
                    error = 0;
                    view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState " + error, "");
                }
            }
            else
            {
                string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trCheckSerialNumberState (" + errorMsg + ")", "");
            }

            return error;
        }

        public string GetNextProductionStep(string serialNumber)
        {
            string workStepDesc = "";
            String[] productionStepResultKeys = new String[] { "WORKSTEP_DESC" };
            String[] productionStepResultValues = new String[] { };
            int error = imsapi.trGetNextProductionStep(sessionContext, init.configHandler.StationNumber, serialNumber, "-1", 0, 1, 1, productionStepResultKeys, out productionStepResultValues);
            if (error != 0)
            {
                string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trGetNextProductionStep " + error + "(" + errorMsg + ")", "");
            }
            else
            {
                if (productionStepResultValues.Length > 0)
                    workStepDesc = productionStepResultValues[0];
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetNextProductionStep " + error, "");
            }
            return workStepDesc;
        }
        #endregion
    }
}
