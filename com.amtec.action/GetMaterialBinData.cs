using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System.Data;

namespace com.amtec.action
{
    public class GetMaterialBinData
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public GetMaterialBinData(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public string[] GetMaterialBinDataDetails(string materialBinNo)
        {
            KeyValue[] materialBinFilters = new KeyValue[] { new KeyValue("MATERIAL_BIN_NUMBER", materialBinNo) };
            AttributeInfo[] attributes = new AttributeInfo[] { };
            string[] materialBinResultKeys = new string[] { "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "MATERIAL_BIN_QTY_TOTAL", "PART_DESC", "MSL_FLOOR_LIFETIME_REMAIN", "EXPIRATION_DATE", "LOCK_STATE" };
            string[] materialBinResultValues = new string[] { };
            LogHelper.Info("begin api mlGetMaterialBinData (Material bin number:" + materialBinNo + ")");
            int error = imsapi.mlGetMaterialBinData(sessionContext, init.configHandler.StationNumber, materialBinFilters, attributes, materialBinResultKeys, out materialBinResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mlGetMaterialBinData (result code = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlGetMaterialBinData " + error + "(" + errorMsg + ")", "");
            }
            return materialBinResultValues;
        }

        public string GetNextMaterialBinNumber(string partNumber)
        {
            string materialBinNo = "";
            LogHelper.Info("begin api mlGetNextMaterialBinNumber (part number:" + partNumber + ")");
            int error = imsapi.mlGetNextMaterialBinNumber(sessionContext, init.configHandler.StationNumber, partNumber, out materialBinNo);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mlGetNextMaterialBinNumber (result code = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlGetNextMaterialBinNumber " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlGetNextMaterialBinNumber " + error + "(" + errorMsg + ")", "");
            }
            return materialBinNo;
        }

        public int CreateMaterialBinNumber(string materialBinNo, string partNumber)
        {
            int error = 0;
            string[] materialBinUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL" };
            string[] materialBinUploadValues = new string[] { "0", materialBinNo, partNumber, "0" };
            string[] materialBinResultValues = new string[] { };
            LogHelper.Info("begin api mlCreateNewMaterialBin (part number:" + partNumber + ",material bin number:" + materialBinNo + ")");
            error = imsapi.mlCreateNewMaterialBin(sessionContext, init.configHandler.StationNumber, materialBinUploadKeys, materialBinUploadValues, out materialBinResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mlCreateNewMaterialBin (result code = " + error + ")");
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlCreateNewMaterialBin " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlCreateNewMaterialBin " + error + "(" + errorMsg + ")", "");
            }
            return error;
        }

        public int ChangeMaterialBinData(string materialBinNo, long expirationDate)
        {
            KeyValue[] materialBinDataUploadValues = new KeyValue[] { new KeyValue("EXPIRATION_DATE", expirationDate.ToString()) };
            int error = imsapi.mlChangeMaterialBinData(sessionContext, init.configHandler.StationNumber, materialBinNo, materialBinDataUploadValues);
            LogHelper.Info("api mlChangeMaterialBinData (material bin number = " + materialBinNo + "result code = " + error + ")");
            return error;
        }

        public DataTable GetBomMaterialData(string workorder)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("ErpGroup", typeof(string)));
            dt.Columns.Add(new DataColumn("PartNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("PartDesc", typeof(string)));
            dt.Columns.Add(new DataColumn("Quantity", typeof(string)));
            dt.Columns.Add(new DataColumn("CompName", typeof(string)));
            dt.Columns.Add(new DataColumn("ProcessLayer", typeof(string)));
            dt.Columns.Add(new DataColumn("ProductFlag", typeof(string)));

            KeyValue[] bomDataFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder), new KeyValue("BOM_ALTERNATIVE", "0") };
            string[] bomDataResultKeys = new string[] { "MACHINE_GROUP_NUMBER", "PART_NUMBER", "PART_DESC", "SETUP_FLAG", "QUANTITY", "COMP_NAME", "PROCESS_LAYER", "PRODUCT_FLAG" };
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
                string strErpGroup = GetErpGroupNumber(init.configHandler.StationNumber, workorder);
                int loop = bomDataResultKeys.Length;
                int count = bomDataResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    if (bomDataResultValues[i + 3] == "1")
                    {
                        //S08SMDXX-02000
                        string strMachineGroup = bomDataResultValues[i].ToString();
                        //S08SMD04-02000-01
                        string strStationNo = init.configHandler.StationNumber;
                        if (strErpGroup == strMachineGroup)
                        {
                            DataRow row = dt.NewRow();
                            row["ErpGroup"] = bomDataResultValues[i].ToString();
                            row["PartNumber"] = bomDataResultValues[i + 1].ToString();
                            row["PartDesc"] = bomDataResultValues[i + 2].ToString();
                            row["Quantity"] = bomDataResultValues[i + 4].ToString();
                            row["CompName"] = bomDataResultValues[i + 5].ToString();
                            row["ProcessLayer"] = bomDataResultValues[i + 6].ToString();
                            row["ProductFlag"] = bomDataResultValues[i + 7].ToString();
                            dt.Rows.Add(row);
                        }
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error, "");
            }
            return dt;
        }

        private string GetErpGroupNumber(string stationNumber, string workorder)
        {
            string erpGroupNo = "";
            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder), new KeyValue("WORKSTEP_FLAG", "1") };
            string[] workplanDataResultKeys = new string[] { "ERP_GROUP_NUMBER" };
            string[] workplanDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetWorkplanData (Work Order:" + workorder + ")");
            int error = imsapi.mdataGetWorkplanData(sessionContext, stationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            LogHelper.Info("end api mdataGetWorkplanData (result code = " + error + ")");
            if (error == 0)
            {
                erpGroupNo = workplanDataResultValues[0];
            }
            return erpGroupNo;
        }

        private int GetProcessLayer(string stationNumber, string workorder)
        {
            int processLayer = 2;
            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder), new KeyValue("WORKSTEP_FLAG", "1") };
            string[] workplanDataResultKeys = new string[] { "PROCESS_LAYER" };
            string[] workplanDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetWorkplanData (Work Order:" + init.currentSettings.workorderNumber + ")");
            int error = imsapi.mdataGetWorkplanData(sessionContext, stationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            LogHelper.Info("end api mdataGetWorkplanData (result code = " + error + ")");
            if (error == 0)
            {
                processLayer = int.Parse(workplanDataResultValues[0]);
            }
            return processLayer;
        }

        #region add by qy

        private string GetErpGroupNumberBySN(string stationNumber, string serialnumber)
        {
            string erpGroupNo = "";
            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("SERIAL_NUMBER", serialnumber), new KeyValue("WORKSTEP_FLAG", "1") };
            string[] workplanDataResultKeys = new string[] { "ERP_GROUP_NUMBER" };
            string[] workplanDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetWorkplanData (Serial Number:" + serialnumber + ")");
            int error = imsapi.mdataGetWorkplanData(sessionContext, stationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            LogHelper.Info("end api mdataGetWorkplanData (result code = " + error + ")");
            if (error == 0)
            {
                erpGroupNo = workplanDataResultValues[0];
            }
            return erpGroupNo;
        }

        public DataTable GetBomMaterialDataBySN(string serialnumber)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add(new DataColumn("ErpGroup", typeof(string)));
            dt.Columns.Add(new DataColumn("PartNumber", typeof(string)));
            dt.Columns.Add(new DataColumn("PartDesc", typeof(string)));
            dt.Columns.Add(new DataColumn("Quantity", typeof(string)));
            dt.Columns.Add(new DataColumn("CompName", typeof(string)));
            dt.Columns.Add(new DataColumn("ProcessLayer", typeof(string)));

            KeyValue[] bomDataFilter = new KeyValue[] { new KeyValue("SERIAL_NUMBER", serialnumber), new KeyValue("BOM_ALTERNATIVE", "0") };
            string[] bomDataResultKeys = new string[] { "MACHINE_GROUP_NUMBER", "PART_NUMBER", "PART_DESC", "SETUP_FLAG", "QUANTITY", "COMP_NAME", "PROCESS_LAYER", "PRODUCT_FLAG" };
            string[] bomDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetBomData (Serial Number:" + serialnumber + ")");
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
                //string strErpGroup = GetErpGroupNumberBySN(init.configHandler.StationNumber, serialnumber);
                int loop = bomDataResultKeys.Length;
                int count = bomDataResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    if (bomDataResultValues[i + 3] == "1")
                    {
                        //S08SMDXX-02000
                        string strMachineGroup = bomDataResultValues[i].ToString();
                        //S08SMD04-02000-01
                        string strStationNo = init.configHandler.StationNumber;
                        //if (strErpGroup == strMachineGroup)
                        //{
                        DataRow row = dt.NewRow();
                        row["ErpGroup"] = bomDataResultValues[i].ToString();
                        row["PartNumber"] = bomDataResultValues[i + 1].ToString();
                        row["PartDesc"] = bomDataResultValues[i + 2].ToString();
                        row["Quantity"] = bomDataResultValues[i + 4].ToString();
                        row["CompName"] = bomDataResultValues[i + 5].ToString();
                        row["ProcessLayer"] = bomDataResultValues[i + 6].ToString();
                        row["ProductFlag"] = bomDataResultValues[i + 7].ToString();
                        dt.Rows.Add(row);
                        //}
                    }
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetBomData " + error, "");
            }
            return dt;
        }

        public int GetProcessLayer(string workorder)
        {
            int processLayer = 0;
            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder), new KeyValue("WORKSTEP_FLAG", "1") };
            string[] workplanDataResultKeys = new string[] { "PROCESS_LAYER" };
            string[] workplanDataResultValues = new string[] { };
            LogHelper.Info("begin api mdataGetWorkplanData (Work Order:" + workorder + ")");
            int error = imsapi.mdataGetWorkplanData(sessionContext, init.configHandler.StationNumber, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            LogHelper.Info("end api mdataGetWorkplanData (result code = " + error + ")");
            if (error == 0)
            {
                processLayer = int.Parse(workplanDataResultValues[0]);
            }
            return processLayer;
        }
        #endregion
    }
}
