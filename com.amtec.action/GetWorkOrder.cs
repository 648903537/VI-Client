using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Data;

namespace com.amtec.action
{
    public class GetWorkOrder
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetWorkOrder(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public DataTable GetAllWorkorders()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("PART_NUMBER", typeof(string)));
            dt.Columns.Add(new DataColumn("WorkorderNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("Status", typeof(string)));

            KeyValue[] workorderFilter = new KeyValue[] { new KeyValue("WORKORDER_STATE", "S,F") };//F = opened (released);S = started
            string[] workorderResultKeys = new string[] { "PART_NUMBER", "WORKORDER_NUMBER", "WORKORDER_STATE" };
            string[] workorderResultValues = new string[] { };
            LogHelper.Info("begin api trGetWorkOrderForStation (Station number:" + init.configHandler.StationNumber + ")");
            int error = imsapi.trGetWorkOrderForStation(sessionContext, init.configHandler.StationNumber, workorderFilter, workorderResultKeys, out workorderResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trGetWorkOrderForStation (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error + "(" + errorMsg + ")", "");
            }
            else
            {
                if (workorderResultValues.Length > 0)
                {
                    DataRow rowEmpty = dt.NewRow();
                    rowEmpty["PART_NUMBER"] = "";
                    rowEmpty["WorkorderNumber"] = "";
                    rowEmpty["Status"] = "";
                    dt.Rows.Add(rowEmpty);

                    int loop = workorderResultKeys.Length;
                    int count = workorderResultValues.Length;
                    for (int i = 0; i < count; i += loop)
                    {
                        string strPartNumber = workorderResultValues[i].ToString();
                        string strWorkorder = workorderResultValues[i + 1].ToString();
                        DataRow row = dt.NewRow();
                        row["PART_NUMBER"] = strPartNumber;
                        row["WorkorderNumber"] = strWorkorder;
                        row["Status"] = GetStatusText(workorderResultValues[i + 2].ToString());
                        dt.Rows.Add(row);
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error, "");
            }
            return dt;
        }

        public DataTable GetWorkordersByFilter(string woFilter)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("Quantity", typeof(string)));
            dt.Columns.Add(new DataColumn("WorkorderNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("Status", typeof(string)));

            KeyValue[] workorderFilter = new KeyValue[] { new KeyValue("WORKORDER_STATE", "S,F"), new KeyValue("WORKORDER_NUMBER", woFilter) };//F = opened (released);S = started
            string[] workorderResultKeys = new string[] { "QUANTITY", "WORKORDER_NUMBER", "WORKORDER_STATE" };
            string[] workorderResultValues = new string[] { };
            LogHelper.Info("begin api trGetWorkOrderForStation (Station number:" + init.configHandler.StationNumber + ")");
            int error = imsapi.trGetWorkOrderForStation(sessionContext, init.configHandler.StationNumber, workorderFilter, workorderResultKeys, out workorderResultValues);
            LogHelper.Info("end api trGetWorkOrderForStation (result code = " + error + ")");
            string errorMsg = "";
            //imsapi.imsapiGetErrorText(sessionContext, error, out errorMsg);
            errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error + ",error message:" + errorMsg, "");
            }
            else
            {
                if (workorderResultValues.Length > 0)
                {
                    int loop = workorderResultKeys.Length;
                    int count = workorderResultValues.Length;
                    for (int i = 0; i < count; i += loop)
                    {
                        DataRow row = dt.NewRow();
                        row["Quantity"] = workorderResultValues[i].ToString();
                        row["WorkorderNumber"] = workorderResultValues[i + 1].ToString();
                        row["Status"] = GetStatusText(workorderResultValues[i + 2].ToString());
                        dt.Rows.Add(row);
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error, "");
            }
            return dt;
        }

        public string[] GetWorkorderByFilter(string woFilter)
        {
            KeyValue[] workorderFilter = new KeyValue[] { new KeyValue("WORKORDER_STATE", "S,F"), new KeyValue("WORKORDER_NUMBER", woFilter) };//F = opened (released);S = started
            string[] workorderResultKeys = new string[] { "PART_NUMBER", "QUANTITY" };
            string[] workorderResultValues = new string[] { };
            LogHelper.Info("begin api trGetWorkOrderForStation (Station number:" + init.configHandler.StationNumber + ",workorder filter:" + woFilter + ")");
            int error = imsapi.trGetWorkOrderForStation(sessionContext, init.configHandler.StationNumber, workorderFilter, workorderResultKeys, out workorderResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trGetWorkOrderForStation (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error + "(" + errorMsg + ")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error, "");
            }
            return workorderResultValues;
        }

        public DataTable GetBomMaterialData(string workorder, string erpGroup)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("PartNo", typeof(string)));
            dt.Columns.Add(new DataColumn("MaterialBinNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("Qty", typeof(string)));
            dt.Columns.Add(new DataColumn("CompName", typeof(string)));
            dt.Columns.Add(new DataColumn("ErpGroup", typeof(string)));
            KeyValue[] bomDataFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder), new KeyValue("PROCESS_BASED", "1") };
            string[] bomDataResultKeys = new string[] { "PART_NUMBER", "QUANTITY", "MACHINE_GROUP_NUMBER" };
            string[] bomDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetBomData (Work Order:" + workorder + ")");
            int error = imsapi.mdataGetBomData(sessionContext, init.configHandler.StationNumber, bomDataFilter, bomDataResultKeys, out bomDataResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mdataGetBomData (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error + "(" + errorMsg + ")", "");
                return null;
            }
            else
            {
                int loop = bomDataResultKeys.Length;
                int count = bomDataResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    if (erpGroup == bomDataResultValues[i + 2].ToString())
                    {
                        DataRow row = dt.NewRow();
                        row["PartNo"] = bomDataResultValues[i].ToString();
                        row["CompName"] = bomDataResultValues[i + 1].ToString();
                        row["ErpGroup"] = bomDataResultValues[i + 2].ToString();
                        row["MaterialBinNumber"] = "";
                        row["Qty"] = "";
                        dt.Rows.Add(row);
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error, "");
            }
            return dt;
        }

        private string GetStatusText(string value)
        {
            string returnValue = "";
            switch (value)
            {
                case "S":
                    returnValue = "started";
                    break;
                case "F":
                    returnValue = "created";
                    break;
                default:
                    returnValue = "started";
                    break;
            }
            return returnValue;
        }

        private string ConvertDateTime(string value)
        {
            long numer = long.Parse(value);
            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(numer).ToLocalTime();
            return date.ToString("yyyy/MM/dd HH:mm:ss");
        }

        public DataTable GetAllWorkordersExt()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("RunID", typeof(string)));
            dt.Columns.Add(new DataColumn("WONumber", typeof(string)));
            dt.Columns.Add(new DataColumn("WODesc", typeof(string)));
            dt.Columns.Add(new DataColumn("Info", typeof(string)));
            dt.Columns.Add(new DataColumn("PartNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("PartDesc", typeof(string)));
            dt.Columns.Add(new DataColumn("MOQty", typeof(string)));
            dt.Columns.Add(new DataColumn("ActualQty", typeof(string)));
            dt.Columns.Add(new DataColumn("Status", typeof(string)));
            //dt.Columns.Add(new DataColumn("Activated", typeof(string)));

            KeyValue[] workorderFilter = new KeyValue[] { new KeyValue("WORKORDER_STATE", "S,F") };//F = opened (released);S = started
            string[] workorderResultKeys = new string[] { "WORKORDER_NUMBER", "WORKORDER_DESC", "QUANTITY", "PartNumber", "PART_DESC", "PLANNED_START_DATE", "WORKORDER_STATE" };
            string[] workorderResultValues = new string[] { };
            LogHelper.Info("begin api trGetWorkOrderForStation (Station number:" + init.configHandler.StationNumber + ")");
            int error = imsapi.trGetWorkOrderForStation(sessionContext, init.configHandler.StationNumber, workorderFilter, workorderResultKeys, out workorderResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api trGetWorkOrderForStation (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error + "(" + errorMsg + ")", "");
            }
            else
            {
                if (workorderResultValues.Length > 0)
                {
                    int loop = workorderResultKeys.Length;
                    int count = workorderResultValues.Length;
                    int n = 0;
                    for (int i = 0; i < count; i += loop)
                    {
                        n++;
                        DataRow row = dt.NewRow();
                        row["RunID"] = n;
                        row["WONumber"] = workorderResultValues[i].ToString();
                        row["WODesc"] = workorderResultValues[i + 1].ToString();
                        row["ActualQty"] = "";
                        row["MOQty"] = workorderResultValues[i + 2].ToString();
                        row["PartNumber"] = workorderResultValues[i + 3].ToString();
                        row["PartDesc"] = workorderResultValues[i + 4].ToString();
                        row["Info"] = ConvertDateTime(workorderResultValues[i + 5].ToString());//WORKORDER_START_DATE
                        row["Status"] = GetStatusText(workorderResultValues[i + 6].ToString());
                        //row["Activated"] = "";
                        dt.Rows.Add(row);
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trGetWorkOrderForStation " + error, "");
            }
            return dt;
        }

        public string[] GetBomData(string serialnumber, string station)
        {
            KeyValue[] bomDataFilter;
            if (init.configHandler.IsGetAllCompByBom == "Y")
            {
                bomDataFilter = new KeyValue[] { new KeyValue("SERIAL_NUMBER", serialnumber), new KeyValue("BOM_ALTERNATIVE", "0") };
            }
            else
            {
                bomDataFilter = new KeyValue[] { new KeyValue("SERIAL_NUMBER", serialnumber), new KeyValue("BOM_ALTERNATIVE", "0"), new KeyValue("PROCESS_BASED", "1") };

            }

            string[] bomDataResultKeys = new string[] { "COMP_NAME", "PART_DESC", "SETUP_FLAG", "PROCESS_LAYER" };
            string[] bomDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetBomData (serialnumber:" + serialnumber + ")");
            int error = imsapi.mdataGetBomData(sessionContext, station, bomDataFilter, bomDataResultKeys, out bomDataResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mdataGetBomData (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(3, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error + "(" + errorMsg + ")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error, "");
            }
            return bomDataResultValues;
        }

        public string[] GetLineNumber(string stationNumber)
        {
            KeyValue[] machineAssetStructureFilter = new KeyValue[] { new KeyValue("DISSOLVING_LEVEL", "1"), new KeyValue("FUNC_MODE", "1"), new KeyValue("STATION_NUMBER", stationNumber) };
            string[] machineAssetStructureResultKeys = new string[] { "LINE_NUMBER", "LINE_DESC" };
            string[] machineAssetStructureValues = new string[] { };
            int error = imsapi.mdataGetMachineAssetStructure(sessionContext, stationNumber, machineAssetStructureFilter, machineAssetStructureResultKeys, out machineAssetStructureValues);
            LogHelper.Info("api mdataGetMachineAssetStructure (station number = " + stationNumber + "), error code =" + error);
            return machineAssetStructureValues;
        }
    }
}
