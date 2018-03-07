using com.amtec.configurations;
using com.amtec.device;
using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace com.amtec.action
{
    public class InitializeMainGUI
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private ApplicationConfiguration config;
        private InitModel initModel;
        private MainView view;
        private LanguageResources lang;
        public bool isInitializeSucces = true;
        FileSystemWatcher watcher = new FileSystemWatcher();

        public InitializeMainGUI(IMSApiSessionContextStruct sessionContext, ApplicationConfiguration config, MainView view, LanguageResources lang)
        {
            this.sessionContext = sessionContext;
            this.config = config;
            this.view = view;
            this.lang = lang;
        }

        public InitModel Initialize()
        {
            initModel = new InitModel();

            try
            {
                initModel.configHandler = config;
                initModel.lang = lang;
            }
            catch
            {
                view.errorHandler(2, "Config error.", "");
                isInitializeSucces = false;
            }

            try
            {
                
                if (config.ScanSNType != "PFC")//config.IsMaterialSetup=="Y"||config.IsEquipSetup=="Y"
                {
                    initModel.scannerHandler = new ScannerHeandler(initModel, view);
                    initModel.scannerHandler.handler().DataReceived += new SerialDataReceivedEventHandler(view.DataRecivedHeandler);
                    initModel.scannerHandler.handler().Open();
                    view.errorHandler(0, lang.ERROR_SCANNER_PORT_OPEN, "");
                }
                if (config.DataOutputInterface == "COM" && config.OutSerialPort != "" && config.OutSerialPort != null)
                {
                    initModel.scannerHandler.OutputCOM().Open();
                    view.errorHandler(0, lang.ERROR_SCANNER_PORT_OPEN, "");
                }
            }
            catch (Exception e)
            {
                view.errorHandler(2, lang.ERROR_SCANNER_PORT_CLOSE, "");
                isInitializeSucces = false;
                LogHelper.Error(e.Message);
            }

            if (isInitializeSucces)
            {
                try
                {
                    //GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, view);
                    //initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
                }
                catch (Exception ex)
                {
                    view.errorHandler(2, "current setting error", "current setting error");
                    isInitializeSucces = false;
                    LogHelper.Error(ex.Message);
                }
            }

            if (initModel.currentSettings != null)
            {
                try
                {
                    GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, view);
                    List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(initModel.currentSettings.partNumber);
                    if (listData != null && listData.Count > 0)
                    {
                        MdataGetPartData mData = listData[0];
                        initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                    }
                }
                catch (Exception ex)
                {
                    view.errorHandler(2, "current setting error", "current setting error");
                    isInitializeSucces = false;
                    LogHelper.Error(ex.Message);
                }

                try
                {
                    switch (initModel.currentSettings.getError)
                    {
                        case 0:
                            view.Invoke(new MethodInvoker(delegate
                            {
                                view.getFieldPartNumber().Text = initModel.currentSettings.partNumber;
                                view.getFieldWorkorder().Text = initModel.currentSettings.workorderNumber;
                            }));
                            break;

                        default:
                            isInitializeSucces = false;
                            return initModel;
                    }
                }
                catch
                {
                    view.errorHandler(2, "Station Setting Error.", "");
                    isInitializeSucces = false;
                    return initModel;
                }
            }
            //read error code ZHS from excel
            try
            {
                string[] LineList = File.ReadAllLines(@"ErrorCodeZH.csv", Encoding.Default);
                Dictionary<int, string> dicErrorCodeMapping = new Dictionary<int, string>();
                if (LineList == null || LineList.Length == 0)
                { }
                else
                {
                    for (int i = 1; i < LineList.Length; i++)
                    {
                        string linecontent = LineList[i].Trim();
                        string[] linegroups = linecontent.Split(new char[] { ',' });
                        int iErrorCode = Convert.ToInt32(linegroups[0].Trim());
                        string strECDesc = linegroups[2].ToString();
                        if (!dicErrorCodeMapping.ContainsKey(iErrorCode))
                            dicErrorCodeMapping[iErrorCode] = strECDesc;
                    }
                }
                initModel.ErrorCodeZHS = dicErrorCodeMapping;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                isInitializeSucces = false;
            }
            if (isInitializeSucces)
            {
                view.errorHandler(0, view.message(initModel.lang.ERROR_INITIALIZE_SUCCESS), initModel.lang.ERROR_INITIALIZE_SUCCESS);
                view.SetStatusLabelText(view.message(initModel.lang.ERROR_INITIALIZE_SUCCESS));
            }
            else
            {
                view.errorHandler(3,view.message(initModel.lang.ERROR_INITIALIZE_ERROR), initModel.lang.ERROR_INITIALIZE_ERROR);
                view.SetStatusLabelText(view.message(initModel.lang.ERROR_INITIALIZE_ERROR));
            }

            #region add by qy
            switch (config.Language)
            {
                case "US":
                    SystemVariable.CurrentLangaugeCode = LanguageType.English;
                    break;
                case "ZHS":
                    SystemVariable.CurrentLangaugeCode = LanguageType.SimplifiedChinese;
                    break;
                case "ZHT":
                    SystemVariable.CurrentLangaugeCode = LanguageType.TraditionalChinese;
                    break;
                default:
                    break;
            }
            if (config.IsListenerFolder == "1")
            {
                ListenerFolder(config.LogFileFolder);
            }
            #endregion
            return initModel;
        }

        public void ListenerFolder(string path)
        {
            view.errorHandler(0, "ListenerFolder Started. " + path, "");
            //process exist files
            //string[] fileNames = Directory.GetFiles(path);
            //if (fileNames != null && fileNames.Length > 0)
            //{
            //    foreach (var item in fileNames)
            //    {
            //        view.ListenFile(item);
            //    }              
            //}
            watcher.Path = path;
            watcher.NotifyFilter = NotifyFilters.FileName; //| NotifyFilters.| NotifyFilters.FileName
            watcher.Filter = "*.*"; //设定监听的文件类型           
            //watcher.Changed += new FileSystemEventHandler(OnUpdated); //暂时不处理
            watcher.Created += new FileSystemEventHandler(OnCreated);
            watcher.EnableRaisingEvents = true;
        }

        private void OnCreated(object source, FileSystemEventArgs e)
        {
            try
            {
                string filename = e.Name;
                string filepath = e.FullPath;
                view.errorHandler(0, "Listener filename:" + filename + " start!", "");
                //string senconds = config.timeout;
                Thread.Sleep(Convert.ToInt32(5000));//等待10s
                try
                {
                }
                catch (Exception ex)
                {
                    LogHelper.Error(ex.Message, ex);
                }

            }
            catch (Exception ex)
            {
                view.errorHandler(3, "Read error." + ex.Message, "");
            }
        }
    }
}
