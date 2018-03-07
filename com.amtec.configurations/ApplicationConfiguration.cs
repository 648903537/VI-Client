using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Linq;

namespace com.amtec.configurations
{
    public class ApplicationConfiguration
    {
        public String StationNumber { get; set; }

        public String Client { get; set; }

        public String RegistrationType { get; set; }

        public String SerialPort { get; set; }

        public String BaudRate { get; set; }

        public String Parity { get; set; }

        public String StopBits { get; set; }

        public String DataBits { get; set; }

        public String NewLineSymbol { get; set; }

        public String DLExtractPattern { get; set; }

        public String MBNExtractPattern { get; set; }

        public String MDAPath { get; set; }

        public String EquipmentExtractPattern { get; set; }

        public String OpacityValue { get; set; }

        public String LocationXY { get; set; }

        public String IPAddress { get; set; }

        public String Port { get; set; }

        public String OutputEnter { get; set; }

        #region add by qy
        public String LogFileFolder { get; set; }

        public String LogTransOK { get; set; }

        public String LogTransError { get; set; }

        public String ChangeFileName { get; set; }

        public String LineRegular { get; set; }

        public String ResultData { get; set; }

        public String SNICT { get; set; }

        public String SNState { get; set; }

        public String LowerLimit { get; set; }

        public String UpperLimit { get; set; }

        public String MeasureFailCode { get; set; }

        public String MeasureName { get; set; }

        public String MeasureValue { get; set; }

        public String UNIT { get; set; }

        public String MNRegular { get; set; }

        public String kValue { get; set; }

        public String MEGValue { get; set; }

        public String MValue { get; set; }

        public String NValue { get; set; }

        public String UValue { get; set; }

        public String PValue { get; set; }

        public String Language { get; set; }

        public String FileType { get; set; }

        public String IsListenerFolder { get; set; }

        public String OutSerialPort { get; set; }

        public String OutBaudRate { get; set; }

        public String OutParity { get; set; }

        public String OutStopBits { get; set; }

        public String OutDataBits { get; set; }

        public String DataOutputInterface { get; set; }

        public String TrayExtractPattern { get; set; }

        public String LogInRegular { get; set; }

        public String LogInType { get; set; }

        public String CheckListFolder { get; set; }

        public String IsSelectWO { get; set; }

        public String IsCheckList { get; set; }

        public String IsMaterialSetup { get; set; }

        public String IsEquipSetup { get; set; }

        public String SDLExtractPattern { get; set; }

        public String ScanSNType { get; set; }

        public String GetCompStation { get; set; }

        public String IsCheckLayer { get; set; }

        public String IsNeedCompColumn { get; set; }

        public String IsGetAllCompByBom { get; set; }

        public String IsNeedInfoField { get; set; }

        public String FileNamePattern { get; set; }

        public String FilterByFileName { get; set; }

        public String OK_CHANNEL_Open { get; set; }

        public String OK_CHANNEL_CLOSE { get; set; }

        public String NG_CHANNEL_OPEN { get; set; }

        public String NG_CHANGE_CLOSE { get; set; }

        public String CommandSleepTime { get; set; }

        public String UserTeam { get; set; }

        public String Upload_BARCODE { get; set; }

        public String RefreshWO { get; set; }

        public String PrintType { get; set; }

        public String IsPrint { get; set; }

        public String MaxiumCount { get; set; }

        public String PrintSerialPort { get; set; }

        public String TemplateFolder { get; set; }

        public String PrinterTypeName { get; set; }

        public String Temp_PartNo { get; set; }

        public String IPI_STATUS_CHECK { get; set; }

        public String Production_Inspection_CHECK { get; set; }

        public String ACTIVE_WORKORDER_LINE { get; set; }

        public String LAYER_DISPLAY { get; set; }

        #region checklist
        public String CHECKLIST_IPAddress { get; set; }
        public String CHECKLIST_Port { get; set; }
        public String CHECKLIST_SOURCE { get; set; }
        public String AUTH_CHECKLIST_APP_TEAM { get; set; }
        public String CHECKLIST_FREQ { get; set; }
        public String SHIFT_CHANGE_TIME { get; set; }
        public String RESTORE_TIME { get; set; }
        public String RESTORE_TREAD_TIMER { get; set; }
        #endregion
        #endregion

        Dictionary<string, string> dicConfig = null;
        public ApplicationConfiguration(IMSApiSessionContextStruct sessionContext, MainView view)
        {

            try
            {
                CommonModel commonModel = ReadIhasFileData.getInstance();
                StationNumber = commonModel.Station;
                Client = commonModel.Client;
                RegistrationType = commonModel.RegisterType;
                if(commonModel.UpdateConfig=="L")
                {
                    XDocument config = XDocument.Load("ApplicationConfig.xml");
                    SerialPort = GetDescendants("SerialPort", config);//config.Descendants("SerialPort").First().Value;
                    BaudRate = GetDescendants("BaudRate", config);//config.Descendants("BaudRate").First().Value;
                    Parity = GetDescendants("Parity", config);//config.Descendants("Parity").First().Value;
                    StopBits = GetDescendants("StopBits", config);// config.Descendants("StopBits").First().Value;
                    DataBits = GetDescendants("DataBits", config);//config.Descendants("DataBits").First().Value;
                    NewLineSymbol = GetDescendants("NewLineSymbol", config);// config.Descendants("NewLineSymbol").First().Value;
                    DLExtractPattern = GetDescendants("DLExtractPattern", config);//config.Descendants("DLExtractPattern").First().Value;
                    MBNExtractPattern = GetDescendants("MBNExtractPattern", config);// config.Descendants("MBNExtractPattern").First().Value;
                    EquipmentExtractPattern = GetDescendants("EquipmentExtractPattern", config);//config.Descendants("EquipmentExtractPattern").First().Value;
                    OpacityValue = GetDescendants("OpacityValue", config);//config.Descendants("OpacityValue").First().Value;
                    LocationXY = GetDescendants("LocationXY", config);//config.Descendants("LocationXY").First().Value;
                    MDAPath = GetDescendants("MDAPath", config);// config.Descendants("MDAPath").First().Value;
                    OutputEnter = GetDescendants("OutputEnter", config);//config.Descendants("OutputEnter").First().Value;

                    #region add by qy
                    Language = GetDescendants("Language", config);//config.Descendants("Language").First().Value;
                    OutSerialPort = GetDescendants("OutSerialPort", config);//config.Descendants("OutSerialPort").First().Value;
                    OutBaudRate = GetDescendants("OutBaudRate", config);//config.Descendants("OutBaudRate").First().Value;
                    OutParity = GetDescendants("OutParity", config);// config.Descendants("OutParity").First().Value;
                    OutStopBits = GetDescendants("OutStopBits", config);//config.Descendants("OutStopBits").First().Value;
                    OutDataBits = GetDescendants("OutDataBits", config);//config.Descendants("OutDataBits").First().Value;
                    DataOutputInterface = GetDescendants("DataOutputInterface", config);//config.Descendants("DataOutputInterface").First().Value;
                    TrayExtractPattern = GetDescendants("TrayExtractPattern", config);//config.Descendants("TrayExtractPattern").First().Value;
                    LogInRegular = GetDescendants("LogInRegular", config);//config.Descendants("LogInRegular").First().Value;
                    LogInType = GetDescendants("LogInType", config);//config.Descendants("LogInType").First().Value;
                    CheckListFolder = GetDescendants("CheckListFolder", config);//config.Descendants("CheckListFolder").First().Value;
                    IsSelectWO = GetDescendants("IsSelectWO", config);// config.Descendants("IsSelectWO").First().Value;
                    IsCheckList = GetDescendants("IsCheckList", config);//config.Descendants("IsCheckList").First().Value;
                    IsMaterialSetup = GetDescendants("IsMaterialSetup", config);// config.Descendants("IsMaterialSetup").First().Value;
                    IsEquipSetup = GetDescendants("IsEquipSetup", config);// config.Descendants("IsEquipSetup").First().Value;
                    SDLExtractPattern = GetDescendants("SDLExtractPattern", config);//config.Descendants("SDLExtractPattern").First().Value;
                    IPAddress = GetDescendants("IPAddress", config);// config.Descendants("IPAddress").First().Value;
                    Port = GetDescendants("Port", config);// config.Descendants("Port").First().Value;
                    ScanSNType = GetDescendants("ScanSNType", config);// config.Descendants("ScanSNType").First().Value;
                    GetCompStation = GetDescendants("GetCompStation", config);//config.Descendants("GetCompStation").First().Value;
                    IsCheckLayer = GetDescendants("IsCheckLayer", config);//config.Descendants("IsCheckLayer").First().Value;
                    IsNeedCompColumn = GetDescendants("IsNeedCompColumn", config);//config.Descendants("IsNeedCompColumn").First().Value;
                    IsGetAllCompByBom = GetDescendants("IsGetAllCompByBom", config);//config.Descendants("IsGetAllCompByBom").First().Value;
                    IsNeedInfoField = GetDescendants("IsNeedInfoField", config);//config.Descendants("IsNeedInfoField").First().Value;

                    FilterByFileName = GetDescendants("FilterByFileName", config);//config.Descendants("FilterByFileName").First().Value;
                    FileNamePattern = GetDescendants("FileNamePattern", config);//config.Descendants("FileNamePattern").First().Value;
                    OK_CHANNEL_Open = GetDescendants("OK_CHANNEL_Open", config);
                    OK_CHANNEL_CLOSE = GetDescendants("OK_CHANNEL_CLOSE", config);
                    NG_CHANNEL_OPEN = GetDescendants("NG_CHANNEL_OPEN", config);
                    NG_CHANGE_CLOSE = GetDescendants("NG_CHANGE_CLOSE", config);
                    CommandSleepTime = GetDescendants("PRESS_TIMER", config);
                    UserTeam = GetDescendants("AUTH_TEAM", config);//config.Descendants("UserTeam").First().Value;
                    Upload_BARCODE = GetDescendants("Upload_BARCODE", config);
                    RefreshWO = GetDescendants("RefreshWO", config);
                    PrintType = GetDescendants("PrintType", config);
                    IsPrint = GetDescendants("Traveler_Slip", config);
                    MaxiumCount = GetDescendants("PCB_Magazine_Rack_Qty", config);
                    PrintSerialPort = GetDescendants("PrintSerialPort", config);
                    TemplateFolder = GetDescendants("TemplateFolder", config);
                    PrinterTypeName = GetDescendants("PrinterTypeName", config);
                    Temp_PartNo = GetDescendants("Temp_PartNo", config);
                    IPI_STATUS_CHECK = GetDescendants("IPI_STATUS_CHECK", config);
                    Production_Inspection_CHECK = GetDescendants("Production_Inspection_CHECK", config);

                    ACTIVE_WORKORDER_LINE = GetDescendants("ACTIVE_WORKORDER_LINE", config);
                    LAYER_DISPLAY = GetDescendants("LAYER_DISPLAY", config);
                    #endregion

                    #region checklist
                    CHECKLIST_IPAddress = GetDescendants("CHECKLIST_IPAddress", config);
                    CHECKLIST_Port = GetDescendants("CHECKLIST_Port", config);
                    CHECKLIST_SOURCE = GetDescendants("CHECKLIST_SOURCE", config);
                    AUTH_CHECKLIST_APP_TEAM = GetDescendants("AUTH_CHECKLIST_APP_TEAM", config);
                    CHECKLIST_FREQ = GetDescendants("CHECKLIST_FREQ", config);
                    SHIFT_CHANGE_TIME = GetDescendants("SHIFT_CHANGE_TIME", config);
                    RESTORE_TIME = GetDescendants("RESTORE_TIME", config);
                    RESTORE_TREAD_TIMER = GetDescendants("RESTORE_TREAD_TIMER", config);
                    #endregion
                }
                else
                {
                    dicConfig = new Dictionary<string, string>();
                    ConfigManage configHandler = new ConfigManage(sessionContext, view);
                    if (commonModel.UpdateConfig == "Y")
                    {
                        //int error = configHandler.DeleteConfigParameters(commonModel.APPTYPE);
                        //if (error == 0 || error == -3303 || error == -3302)
                        //{
                        //    WriteParameterToiTac(configHandler);
                        //}
                        string[] parametersValue = configHandler.GetParametersForScope(commonModel.APPTYPE);
                        if (parametersValue != null && parametersValue.Length > 0)
                        {
                            foreach (var parameterID in parametersValue)
                            {
                                configHandler.DeleteConfigParametersExt(parameterID);
                            }
                        }
                        WriteParameterToiTac(configHandler);
                    }
                    List<ConfigEntity> getvalues = configHandler.GetConfigData(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station);
                    if (getvalues != null)
                    {
                        foreach (var item in getvalues)
                        {
                            if (item != null)
                            {
                                string[] strs = item.PARAMETER_NAME.Split(new char[] { '.' });
                                dicConfig.Add(strs[strs.Length - 1], item.CONFIG_VALUE);
                            }
                        }
                    }
                    SerialPort = GetParameterValue("SerialPort");
                    BaudRate = GetParameterValue("BaudRate");
                    Parity = GetParameterValue("Parity");
                    StopBits = GetParameterValue("StopBits");
                    DataBits = GetParameterValue("DataBits");
                    NewLineSymbol = GetParameterValue("NewLineSymbol");
                    DLExtractPattern = GetParameterValue("DLExtractPattern");
                    MBNExtractPattern = GetParameterValue("MBNExtractPattern");
                    EquipmentExtractPattern = GetParameterValue("EquipmentExtractPattern");
                    OpacityValue = GetParameterValue("OpacityValue");
                    LocationXY = GetParameterValue("LocationXY");
                    MDAPath = GetParameterValue("MDAPath");
                    OutputEnter = GetParameterValue("OutputEnter");

                    #region add by qy
                    Language = GetParameterValue("Language");
                    OutSerialPort = GetParameterValue("OutSerialPort");
                    OutBaudRate = GetParameterValue("OutBaudRate");
                    OutParity = GetParameterValue("OutParity");
                    OutStopBits = GetParameterValue("OutStopBits");
                    OutDataBits = GetParameterValue("OutDataBits");
                    DataOutputInterface = GetParameterValue("DataOutputInterface");
                    TrayExtractPattern = GetParameterValue("TrayExtractPattern");
                    LogInRegular = GetParameterValue("LogInRegular");
                    LogInType = GetParameterValue("LogInType");
                    CheckListFolder = GetParameterValue("CheckListFolder");
                    IsSelectWO = GetParameterValue("IsSelectWO");
                    IsCheckList = GetParameterValue("IsCheckList");
                    IsMaterialSetup = GetParameterValue("IsMaterialSetup");
                    IsEquipSetup = GetParameterValue("IsEquipSetup");
                    SDLExtractPattern = GetParameterValue("SDLExtractPattern");
                    IPAddress = GetParameterValue("IPAddress");
                    Port = GetParameterValue("Port");
                    ScanSNType = GetParameterValue("ScanSNType");
                    GetCompStation = GetParameterValue("GetCompStation");
                    IsCheckLayer = GetParameterValue("IsCheckLayer");
                    IsNeedCompColumn = GetParameterValue("IsNeedCompColumn");
                    IsGetAllCompByBom = GetParameterValue("IsGetAllCompByBom");
                    IsNeedInfoField = GetParameterValue("IsNeedInfoField");
                    OK_CHANNEL_Open = GetParameterValue("OK_CHANNEL_Open");
                    OK_CHANNEL_CLOSE = GetParameterValue("OK_CHANNEL_CLOSE");
                    NG_CHANNEL_OPEN = GetParameterValue("NG_CHANNEL_OPEN");
                    NG_CHANGE_CLOSE = GetParameterValue("NG_CHANGE_CLOSE");
                    #endregion
                }
                
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private string GetParameterValue(string parameterName)
        {
            if (dicConfig.ContainsKey(parameterName))
            {
                return dicConfig[parameterName];
            }
            else
            {
                return "";
            }
        }

        private void WriteParameterToiTac(ConfigManage configHandler)
        {
            GetApplicationDatas getData = new GetApplicationDatas();
            List<ParameterEntity> entityList = getData.GetApplicationEntity();
            string[] strs = GetParameterString(entityList);
            string[] strvalues = GetValueString(entityList);
            if (strs != null && strs.Length > 0)
            {
                int errorCode = configHandler.CreateConfigParameter(strs);
                if (errorCode == 0 || errorCode == 5)
                {
                    CommonModel commonModel = ReadIhasFileData.getInstance();
                    int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
                }
            }

            //if (entityList.Count > 0)
            //{
            //    List<ParameterEntity> entitySubList = null;
            //    CommonModel commonModel = ReadIhasFileData.getInstance();
            //    foreach (var entity in entityList)
            //    {
            //        entitySubList = new List<ParameterEntity>();
            //        entitySubList.Add(entity);
            //        string[] strs = GetParameterString(entitySubList);
            //        string[] strvalues = GetValueString(entitySubList);
            //        if (strs != null && strs.Length > 0)
            //        {
            //            int errorCode = configHandler.CreateConfigParameter(strs);
            //            if (errorCode == 0 || errorCode == 5)
            //            {                           
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //            else if (errorCode == -3301)//Parameter already exists
            //            {
            //                int re = configHandler.UpdateParameterValues(commonModel.APPID, commonModel.APPTYPE, commonModel.Cluster, commonModel.Station, strvalues);
            //            }
            //        }
            //    }
            //}
        }

        private string[] GetParameterString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                strList.Add(entity.PARAMETER_DESCRIPTION);
                strList.Add(entity.PARAMETER_DIMPATH);
                strList.Add(entity.PARAMETER_DISPLAYNAME);
                strList.Add(entity.PARAMETER_NAME);
                strList.Add(entity.PARAMETER_PARENT_NAME);
                strList.Add(entity.PARAMETER_SCOPE);
                strList.Add(entity.PARAMETER_TYPE_NAME);
            }
            return strList.ToArray();
        }

        private string[] GetValueString(List<ParameterEntity> entityList)
        {
            List<string> strList = new List<string>();
            foreach (var entity in entityList)
            {
                if (entity.PARAMETER_VALUE == "")
                    continue;
                strList.Add(entity.PARAMETER_VALUE);
                strList.Add(entity.PARAMETER_NAME);

            }
            return strList.ToArray();
        }

        private string GetDescendants(string parameter, XDocument _config)
        {
            try
            {
                string value = _config.Descendants(parameter).First().Value;

                return value;
            }
            catch
            {
                LogHelper.Info("Parameter is not exist." + parameter);
                return "";
            }
        }
    }
}
