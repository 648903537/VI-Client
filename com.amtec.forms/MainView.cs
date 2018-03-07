using com.amtec.action;
using com.amtec.configurations;
using com.amtec.model;
using com.itac.mes.imsapi.domain.container;
using com.itac.oem.common.container.imsapi.utils;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Net.NetworkInformation;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Linq;
using System.Timers;
using Compal.Onlineprg.Printing.Port;
using System.Drawing.Printing;
using ScreenPrinter.com.amtec.forms;
using com.amtec.device;
using System.Collections.Concurrent;

namespace com.amtec.forms
{
    public partial class MainView : Form
    {
        public ApplicationConfiguration config;
        IMSApiSessionContextStruct sessionContext;
        IMSApiSessionContextStruct sessionContext1;
        public bool isScanProcessEnabled = false;
        private InitModel initModel;
        private LanguageResources res;
        public string UserName = "";
        private string indata = "";
        private DateTime PFCStartTime = DateTime.Now;
        List<SerialNumberData> serialNumberArray = new List<SerialNumberData>();
        public delegate void HandleInterfaceUpdateTopMostDelegate(string sn, string message);
        public HandleInterfaceUpdateTopMostDelegate topmostHandle;
        public TopMostForm topmostform = null;
        CommonModel commonModel = null;
        public string CaptionName = "";
        private int iProcessLayer = 2;
        private string TrayNumber = "";
        private int Pscount = 0;

        private System.Timers.Timer RestoreMaterialTimer = null;
        string Supervisor_OPTION = "1";
        string IPQC_OPTION = "1";
        private SocketClientHandler2 checklist_cSocket = null;
        bool isStartLineCheck = true;//开线点检已经获取=true. 过程点检=false

        #region Init
        public MainView(string userName, DateTime dTime, IMSApiSessionContextStruct _sessionContext, IMSApiSessionContextStruct _sessionContext1)
        {
            InitializeComponent();
            sessionContext = _sessionContext;
            sessionContext1 = _sessionContext1;
            UserName = userName;
            commonModel = ReadIhasFileData.getInstance();
            this.lblLoginTime.Text = dTime.ToString("yyyy/MM/dd HH:mm:ss");
            this.lblUser.Text = userName == "" ? commonModel.Station : userName;
            this.lblStationNO.Text = commonModel.Station;
            //ReadLogfile(@"C:\Users\atl_liuxue\Desktop\GreateWall\ViT AOI\output\OCR output\VIT08_AOI_611398700000000129_2009-08-05_15-53-08.xml");
        }

        private void MainView_Shown(object sender, EventArgs e)
        {
            BackgroundWorker bgWorker = new BackgroundWorker();
            bgWorker.DoWork += new DoWorkEventHandler(bgWorkerInit);
            bgWorker.RunWorkerAsync();
        }

        #region add by qy
        private void InitCintrolLanguage()
        {
            MutiLanguages lang = new MutiLanguages();
            foreach (Control ctl in this.Controls)
            {
                lang.InitLangauge(ctl);
                if (ctl is TabControl)
                {
                    lang.InitLangaugeForTabControl((TabControl)ctl);
                }
            }

            //Controls不包含ContextMenuStrip，可用以下方法获得
            System.Reflection.FieldInfo[] fieldInfo = this.GetType().GetFields(System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);
            for (int i = 0; i < fieldInfo.Length; i++)
            {
                switch (fieldInfo[i].FieldType.Name)
                {
                    case "ContextMenuStrip":
                        ContextMenuStrip contextMenuStrip = (ContextMenuStrip)fieldInfo[i].GetValue(this);
                        lang.InitLangauge(contextMenuStrip);
                        break;
                }
            }
        }
        #endregion
        bool isOK = true;
        private void bgWorkerInit(object sender, DoWorkEventArgs args)
        {
            errorHandler(0, "Application start...", "");
            errorHandler(0, "Version :" + Assembly.GetExecutingAssembly().GetName().Version.ToString(), "");
            res = new LanguageResources();
            config = new ApplicationConfiguration(sessionContext, this);
            InitializeMainGUI init = new InitializeMainGUI(sessionContext, config, this, res);
            initModel = init.Initialize();
            this.InvokeEx(x =>
            {
                this.Text = res.MAIN_TITLE + " (" + Assembly.GetExecutingAssembly().GetName().Version.ToString() + ")";
                CaptionName = res.MAIN_TITLE + System.Environment.NewLine + config.StationNumber;
                GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, this);
                initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
                if (initModel.currentSettings != null)
                {
                    this.txbCDAMONumber.Text = initModel.currentSettings.workorderNumber;
                    this.txbCDAPartNumber.Text = initModel.currentSettings.partNumber;
                    iProcessLayer = initModel.currentSettings.processLayer;
                    this.txtLayer.Text = ConvertProcessLayerToString2(iProcessLayer.ToString());
                }

                LoadYield();
                if (config.LAYER_DISPLAY == "")
                {
                    this.btnShowPCB.Visible = false;
                }
                if (config.IsCheckList != "Y")
                {
                    this.tabCheckList.Parent = null;
                    this.tabCheckListTable.Parent = null;
                }
                else
                {
                    if (config.AUTH_CHECKLIST_APP_TEAM != "" && config.AUTH_CHECKLIST_APP_TEAM != null)
                    {
                        string[] teams = config.AUTH_CHECKLIST_APP_TEAM.Split(';');
                        string[] items = teams[0].Split(',');
                        string Super = items[0];
                        Supervisor_OPTION = items[1];
                        string[] IPQCitems = teams[1].Split(',');
                        string IP = IPQCitems[0];
                        IPQC_OPTION = IPQCitems[1];
                    }
                    if (config.RESTORE_TIME != "" && config.RESTORE_TREAD_TIMER != "")
                    {
                        GetRestoreTimerStart();
                    }
                    if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")//20161208 edit by qy
                    {
                        InitShift2(txbCDAMONumber.Text);
                        InitWorkOrderType();
                        this.tabCheckList.Parent = null;
                        checklist_cSocket = new SocketClientHandler2(this);
                        isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                        if (isOK)
                        {
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                if (!ReadCheckListFile())//20161214 edit by qy
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                        }
                    }
                    else
                    {
                        InitTaskData();
                        this.tabCheckListTable.Parent = null;
                    }
                }
                InitCheckResultMapping();
                if (config.IsSelectWO != "Y")
                {
                    this.tabWo.Parent = null;

                    //CheckIPIStatus();
                }
                else
                {
                    if (config.ACTIVE_WORKORDER_LINE.ToUpper() == "DISABLE")
                    {
                        this.panel15.Visible = false;
                        this.panel14.Location = new Point(0, 3);
                    }

                    InitWorkOrder();
                }
                if (config.IsMaterialSetup != "Y")
                {
                    this.tabSetup.Parent = null;
                }
                if (config.IsEquipSetup != "Y")
                {
                    this.tabEquipment.Parent = null;
                    this.tabMachine.Parent = null;
                }
                if (config.RefreshWO != "Y")
                {
                    this.btnRefreshWo.Visible = false;
                }
                if (config.IsPrint != "Enable")
                {
                    this.dgvSNInfo.Size = new Size(335, 382);
                    plRack.Visible = false;
                }
                else
                {
                    InitPrinterList();
                    lblTextRackNo.Text = GetNextSN();
                    lblPCBQty.Text = "0/" + config.MaxiumCount;
                    CheckLocalFile();
                    Pscount = Convert.ToInt32(this.lblPCBQty.Text.Split('/')[0]);
                }
                #region add by qy
                //InitFailureMapTable();
                InitCintrolLanguage();
                #endregion

                this.cmbLayer.Text = "T";
                ShowTopWindow();
                InitDocumentGrid();
                SetTipMessage(MessageType.OK, message("Initialize Success"));
                this.txbCDADataInput.Focus();
                SetAlarmStatusText("");

                if (config.LogInType == "COM")
                {
                    this.txbCDADataInput.ReadOnly = true;
                }
                if (config.ScanSNType == "PFC")
                {
                    Application.DoEvents();
                    cSocket = new SocketClientHandler(this);
                    bool isOK = cSocket.connect(config.IPAddress, config.Port);
                    if (isOK)
                    {
                        GetTimerStart();
                    }
                }
                else
                {
                    this.tabConnection.Parent = null;
                }

                if (config.IsNeedCompColumn == "N")
                {
                    dgvCompName.Visible = false;
                    cmbComp.Visible = false;
                    this.dgvSN.Size = new Size(361, 415);
                    this.cmbDefct.Size = new Size(217, 25);
                    if (config.IsNeedInfoField == "N")
                    {
                        this.dgvDefect.Size = new Size(217, 382);//217, 415
                    }
                    else
                    {
                        this.dgvDefect.Size = new Size(217, 351);//217, 382
                    }
                    if (config.IsPrint == "Enable")
                    {
                        this.dgvSNInfo.Size = new Size(406, 250);
                        this.plRack.Size = new Size(406, 127);
                    }
                    else
                    {
                        this.dgvSNInfo.Size = new Size(406, 382);
                    }
                    this.pnlInfo.Location = new Point(372, 393);
                    this.dgvDefect.Location = new Point(372, 36);//372, 5
                    this.cmbDefct.Location = new Point(372, 5);
                    this.dgvSNInfo.Location = new Point(594, 5);//448, 5
                    this.plRack.Location = new Point(594, 260);//448, 5
                    this.dgvSNInfo.Columns[2].Visible = false;
                    this.dgvSNInfo.Columns[3].Visible = false;
                }
                if (config.IsNeedInfoField == "N")
                {
                    pnlInfo.Visible = false;
                    if (config.IsNeedCompColumn != "N")
                    {
                        this.dgvCompName.Size = new Size(213, 382);//415
                        this.dgvDefect.Size = new Size(140, 382);//415
                    }
                    this.dgvSNInfo.Columns[7].Visible = false;
                }
            });
            process = new Thread(new ThreadStart(ProcessRecipeCommand));
            process.Start();
        }
        public string message(string info)
        {
            string msg = "";
            msg = MutiLanguages.ParserString("$msg_" + info);
            return msg;
        }
        #endregion

        #region delegate
        public delegate void errorHandlerDel(int typeOfError, String logMessage, String labelMessage);
        public void errorHandler(int typeOfError, String logMessage, String labelMessage)
        {
            if (txtConsole.InvokeRequired)
            {
                errorHandlerDel errorDel = new errorHandlerDel(errorHandler);
                Invoke(errorDel, new object[] { typeOfError, logMessage, labelMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        SetAlarmStatusText("");
                        LogHelper.Info(logMessage);
                        break;
                    case 1:
                        isSucces = "SUCCESS";
                        txtConsole.SelectionColor = Color.Blue;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.OK, logMessage);
                        SetAlarmStatusText("");
                        LogHelper.Info(logMessage);
                        break;
                    case 2:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        SetAlarmStatusText("失败");
                        LogHelper.Error(logMessage);
                        break;
                    case 3:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        SetTipMessage(MessageType.Error, logMessage);
                        SetAlarmStatusText("失败");
                        LogHelper.Error(logMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConsole.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + logMessage + "\n";
                        LogHelper.Error(logMessage);
                        break;
                }
                
                SetStatusLabelText(logMessage);
                txtConsole.AppendText(errorBuilder);
                txtConsole.ScrollToCaret();
            }
        }

        public void SetStatusLabelText(string strText)
        {
            this.InvokeEx(x => this.lblStatus.Text = strText);
        }

        public string GetWorkOrderValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAMONumber.Text);
            return str;
        }

        public string GetPartNumberValue()
        {
            string str = "";
            this.InvokeEx(x => str = this.txbCDAPartNumber.Text);
            return str;
        }

        public TextBox getFieldPartNumber()
        {
            return this.txbCDAPartNumber;
        }

        public TextBox getFieldWorkorder()
        {
            return this.txbCDAMONumber;
        }

        public Label getFieldLabelUser()
        {
            return lblUser;
        }

        public Label getFieldLabelTime()
        {
            return lblLoginTime;
        }

        private delegate void EditDataToTraySNHandle(string serialnumberref, string result);
        private void EditDataToTraySNGrid(string serialnumberref, string result)
        {
            try
            {
                if (this.dgvSN.InvokeRequired)
                {
                    EditDataToTraySNHandle EditDataDel = new EditDataToTraySNHandle(EditDataToTraySNGrid);
                    Invoke(EditDataDel, new object[] { serialnumberref, result });
                }
                else
                {
                    foreach (DataGridViewRow row in dgvSN.Rows)
                    {
                        if (serialnumberref == row.Cells["Column3"].Value.ToString())
                        {
                            row.Cells["Column4"].Value = result;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        public delegate void SetTipMessageDel(MessageType strType, string strMessage);
        private void SetTipMessage(MessageType strType, string strMessage)
        {
            if (this.messageControl1.InvokeRequired)
            {
                SetTipMessageDel messageDel = new SetTipMessageDel(SetTipMessage);
                Invoke(messageDel, new object[] { strType, strMessage });
            }
            else
            {
                switch (strType)
                {
                    case MessageType.OK:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Error:
                        this.messageControl1.BackColor = Color.Red;
                        this.messageControl1.PicType = @"pic\Close.png";
                        this.messageControl1.Title = "Error Message";
                        this.messageControl1.Content = strMessage;
                        break;
                    case MessageType.Instruction:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\Instruction.png";
                        this.messageControl1.Title = "Instruction";
                        this.messageControl1.Content = strMessage;
                        break;
                    default:
                        this.messageControl1.BackColor = Color.FromArgb(184, 255, 160);
                        this.messageControl1.PicType = @"pic\ok.png";
                        this.messageControl1.Title = "OK";
                        this.messageControl1.Content = strMessage;
                        break;
                }
            }
        }
        #endregion

        #region Data process function
        string tempserialnumber = "";
        public void DataRecivedHeandler(object sender, SerialDataReceivedEventArgs e)
        {
            SerialPort sp = (SerialPort)sender;
            try
            {
                if (isFormOutPoump)
                {
                    return;
                }
                Thread.Sleep(200);
                Byte[] bt = new Byte[sp.BytesToRead];
                sp.Read(bt, 0, sp.BytesToRead);
                indata = System.Text.Encoding.ASCII.GetString(bt).Trim();
                //indata = sp.ReadLine();
                if (indata.Length == 2)
                {
                    initModel.scannerHandler.handler().DiscardInBuffer();
                    return;
                }
                if (config.IsCheckList == "Y")
                {
                    if (!VerifyCheckList())
                    {
                        return;
                    }
                }

                if ((txbCDAMONumber.Text == "" || txbCDAPartNumber.Text == "") && config.IsSelectWO == "Y")
                {
                    errorHandler(2, message("Please select a workorder"), "");
                    return;
                }
                LogHelper.Info("Scan number(original): " + indata);
                this.Invoke(new MethodInvoker(delegate
                {
                    this.txbCDADataInput.Text = indata;
                }));
                SetAlarmStatusText("");
                Match match = null;

                #region material bin number
                //match material bin number
                match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, message("Scan material bin number"));
                    }));
                    ProcessMaterialBinNo(match.ToString());
                    return;
                }
                #endregion

                #region equipment
                //match equipment
                match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, message("Scan equipment number"));
                    }));
                    ProcessEquipmentData(match.ToString());
                    return;
                }
                #endregion

                #region serialnumber
                //match serial number
                //match = Regex.Match(indata, config.SDLExtractPattern);
                //if (match.Success)
                //{
                //    if (this.dgvEquipment.RowCount > 0)
                //    {
                //        if (!CheckEquipmentSetup())
                //        {
                //            return;
                //        }
                //    }
                //    this.Invoke(new MethodInvoker(delegate
                //    {
                //        this.tabControl1.SelectedTab = this.tabDetail;
                //        SetTipMessage(MessageType.OK, message("Scan serial number"));
                //    }));
                //    tempserialnumber = match.ToString();
                //    ProcessSerialNumberSingle(match.ToString());
                //    return;
                //}

                match = Regex.Match(indata, config.DLExtractPattern);
                if (match.Success)
                {
                    if (this.dgvEquipment.RowCount > 0)
                    {
                        if (!CheckEquipmentSetup())
                        {
                            return;
                        }
                    }
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabDetail;
                        SetTipMessage(MessageType.OK, message("Scan serial number"));
                    }));
                    tempserialnumber = match.ToString();
                    ProcessSerialNumber(match.ToString());
                    return;
                }

                #endregion

                #region upload barcode
                //match upload barcode
                match = Regex.Match(indata, config.Upload_BARCODE);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabDetail;
                        btnUpload_Click(null, null);
                    }));

                    return;
                }
                #endregion

                errorHandler(3, message("wrong barcode"), "");
            }
            catch (Exception ex)
            {
                errorHandler(2, ex.Message, "");
                LogHelper.Error(ex.Message);
            }
            finally
            {
                initModel.scannerHandler.handler().DiscardInBuffer();
            }
        }

        #endregion

        #region Event
        private void btnRefreshWo_Click(object sender, EventArgs e)
        {
            //if (config.IsCheckList == "Y")
            //{
            //    if (!VerifyCheckList())
            //    {
            //        return;
            //    }
            //}
            string processlayerPre = "";
            string WorkorderPre = this.txbCDAMONumber.Text;
            if (initModel.currentSettings != null)
                processlayerPre = initModel.currentSettings.processLayer.ToString();
            GetCurrentWorkorder currentWorkorder = new GetCurrentWorkorder(sessionContext, initModel, this);
            initModel.currentSettings = currentWorkorder.GetCurrentWorkorderResultCall();
            if (initModel.currentSettings == null)
                return;
            this.txbCDAMONumber.Text = initModel.currentSettings.workorderNumber;
            this.txbCDAPartNumber.Text = initModel.currentSettings.partNumber;
            iProcessLayer = initModel.currentSettings.processLayer;
            this.txtLayer.Text = ConvertProcessLayerToString2(iProcessLayer.ToString());
            if (initModel.currentSettings.workorderNumber != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer.ToString())
            {
                strShiftChecklist = "";
                InitWorkOrderType();
                InitShift2(WorkorderPre);//20161215 add by qy
                if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                {
                    if (!CheckShiftChange2())
                    {
                        InitTaskData_SOCKET("开线点检;设备点检");
                        isStartLineCheck = true;
                    }
                    else
                    {
                        InitTaskData_SOCKET("开线点检");
                        isStartLineCheck = true;
                    }
                }
            }

            this.Invoke(new MethodInvoker(delegate
            {
                LoadYield();
            }));
            InitDocumentGrid();
            //CheckIPIStatus();
        }
        private void MainView_Load(object sender, EventArgs e)
        {
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath) + @"\pic\Chart_Column_Silver.png";
            NetworkChange.NetworkAvailabilityChanged += AvailabilityChanged;
        }

        private void MainView_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult dr = MessageBox.Show(message("Do you want to close the application"), message("Quit Application"), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.OK)
            {
                if (this.txbCDAMONumber.Text != "")
                {
                    EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                    SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                    string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
                    foreach (var wo in wos)
                    {
                        setupHandler.SetupStateChange(wo, iProcessLayer, 1);
                    }
                    foreach (DataGridViewRow row in dgvEquipment.Rows)
                    {
                        string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                        if (string.IsNullOrEmpty(equipmentNo))
                            continue;
                        int errorCode = eqManager.UpdateEquipmentData(wos[0], iProcessLayer, equipmentNo, 1);
                    }
                    foreach (DataGridViewRow rowMAC in this.gridMachine.Rows)
                    {
                        string equipmentNo = rowMAC.Cells["EquipNoMac"].Value.ToString();
                        if (string.IsNullOrEmpty(equipmentNo))
                            continue;
                        int errorCode = eqManager.UpdateEquipmentData(wos[0], iProcessLayer, equipmentNo, 1);
                    }

                    //if (File.Exists(path))
                    //{
                    //    File.Delete(path);
                    //}
                }
                LogHelper.Info("Application end...");
                System.Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }
        }

        private void txbCDADataInput_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                if (config.IsCheckList == "Y")
                {
                    if (!VerifyCheckList())
                    {
                        return;
                    }
                }
                if ((txbCDAMONumber.Text == "" || txbCDAPartNumber.Text == "") && config.IsSelectWO == "Y")
                {
                    errorHandler(2, message("Please select a workorder"), "");
                    return;
                }
                indata = this.txbCDADataInput.Text.Trim();
                SetAlarmStatusText("");
                Match match = null;

                //match material bin number
                match = Regex.Match(indata, config.MBNExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabSetup;
                        SetTipMessage(MessageType.OK, message("Scan material bin number"));
                    }));
                    ProcessMaterialBinNo(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }

                //match equipment
                match = Regex.Match(indata, config.EquipmentExtractPattern);
                if (match.Success)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabEquipment;
                        SetTipMessage(MessageType.OK, message("Scan equipment number"));
                    }));
                    ProcessEquipmentData(match.ToString());
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                //match serial number
                if (config.ScanSNType != "PFC")
                {
                    //edit by qy 20170412 由原来的扫描小板只带出小板改为带出所有板子，并将报废置顶（第一次扫描带出条码，第二次扫描为确定选择的小板条码）
                    //match = Regex.Match(indata, config.SDLExtractPattern);
                    //if (match.Success)
                    //{
                    //    if (this.dgvEquipment.RowCount > 0)
                    //    {
                    //        if (!CheckEquipmentSetup())
                    //        {
                    //            return;
                    //        }
                    //    }
                    //    this.Invoke(new MethodInvoker(delegate
                    //    {
                    //        this.tabControl1.SelectedTab = this.tabDetail;
                    //        SetTipMessage(MessageType.OK, message("Scan serial number"));
                    //    }));
                    //    tempserialnumber = match.ToString();
                    //    ProcessSerialNumberSingle(match.ToString());
                    //    this.txbCDADataInput.SelectAll();
                    //    this.txbCDADataInput.Focus();
                    //    return;
                    //}
                    match = Regex.Match(indata, config.DLExtractPattern);
                    if (match.Success)
                    {
                        if (this.dgvEquipment.RowCount > 0)
                        {
                            if (!CheckEquipmentSetup())
                            {
                                return;
                            }
                        }
                        this.Invoke(new MethodInvoker(delegate
                        {
                            this.tabControl1.SelectedTab = this.tabDetail;
                            SetTipMessage(MessageType.OK, message("Scan serial number"));
                        }));
                        tempserialnumber = match.ToString();
                        ProcessSerialNumber(match.ToString());
                        this.txbCDADataInput.SelectAll();
                        this.txbCDADataInput.Focus();
                        return;
                    }
                }
                #region upload barcode
                //match upload barcode
                match = Regex.Match(indata, config.Upload_BARCODE);
                if (match.Success)
                {
                    LogHelper.Info("TEST");
                    this.Invoke(new MethodInvoker(delegate
                    {
                        this.tabControl1.SelectedTab = this.tabDetail;
                    }));

                    btnUpload_Click(null, null);
                    this.txbCDADataInput.SelectAll();
                    this.txbCDADataInput.Focus();
                    return;
                }
                #endregion

                this.txbCDADataInput.SelectAll();
                this.txbCDADataInput.Focus();
                errorHandler(3, message("wrong barcode"), "");
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.tabControl1.SelectedTab.Name == "tabCheckList")
            {
                this.gridSetup.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabEquipment")
            {
                this.dgvEquipment.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabDocument")
            {
                this.gridDocument.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabMachine")
            {
                this.gridMachine.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabSetup")
            {
                this.gridCheckList.ClearSelection();
            }
            else if (this.tabControl1.SelectedTab.Name == "tabWo")
            {
                this.gridWorkorder.ClearSelection();
                SetWorkorderGridStatus();
            }
        }

        private void gridDocument_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            long documentID = Convert.ToInt64(gridDocument.Rows[e.RowIndex].Cells[0].Value.ToString());
            string fileName = gridDocument.Rows[e.RowIndex].Cells[1].Value.ToString();
            SetDocumentControlForDoc(documentID, fileName);
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            if (this.dgvEquipment.SelectedRows.Count > 0)
            {
                DataGridViewRow row = this.dgvEquipment.SelectedRows[0];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["Status"].Value = ICTClient.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
                foreach (var wo in wos)
                {
                    int errorCode = eqManager.UpdateEquipmentData(wo, 2, equipmentNo, 1);
                }
                this.dgvEquipment.ClearSelection();
            }
        }

        int iIndexItem = -1;
        private void dgvEquipment_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvEquipment.Rows.Count == 0)
                    return;
                this.dgvEquipment.ContextMenuStrip = contextMenuStrip1;
                iIndexItem = ((DataGridView)sender).CurrentRow.Index;
            }
        }

        private void removeEquipment_Click(object sender, EventArgs e)
        {
            if (iIndexItem > -1)
            {
                DataGridViewRow row = this.dgvEquipment.Rows[iIndexItem];
                string equipmentNo = row.Cells["EquipNo"].Value.ToString();
                row.Cells["NextMaintenance"].Value = "";
                row.Cells["ScanTime"].Value = "";
                row.Cells["UsCount"].Value = "";
                row.Cells["EquipNo"].Value = "";
                row.Cells["Status"].Value = ICTClient.Properties.Resources.Close;
                row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);

                //Strip down equipment
                if (string.IsNullOrEmpty(equipmentNo))
                    return;
                EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
                foreach (var wo in wos)
                {
                    int errorCode = eqManager.UpdateEquipmentData(wo, 2, equipmentNo, 1);
                }
                this.dgvEquipment.ClearSelection();
            }
        }

        int PanelQty = 0;
        private void cmbWO_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbWO.SelectedIndex == 0)
                return;
            if (cmbWO.SelectedValue.ToString() != "")
            {
                this.txbCDAMONumber.Text = cmbWO.Text;
                this.txbCDAPartNumber.Text = cmbWO.SelectedValue.ToString();
                PanelQty = GetPCBPanelQty();
                LoadYield();
                InitSetupGrid();
                InitEquipmentGrid();
                InitMachineGrid();
            }
        }

        private void gridWorkorder_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (config.IsCheckList == "Y")
            {
                if (!CheckCheckList())
                {
                    return;
                }
            }
            if (e.RowIndex == -1)
                return;
            DialogResult dr = MessageBox.Show(message("select_workorder_confirm"), message("Information"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {

                int rowWO = gridWorkorder.CurrentRow.Index;//获得选种行的索引

                string selectWO = "";
                string selectPartNumber = "";
                string matchWO = "";
                string matchPartNumner = "";

                if (rowWO > -1 && rowWO < gridWorkorder.Rows.Count)
                {
                    selectPartNumber = gridWorkorder.Rows[rowWO].Cells[5].Value.ToString().Trim();
                    selectWO = gridWorkorder.Rows[rowWO].Cells[2].Value.ToString().Trim();
                    partdesc = gridWorkorder.Rows[rowWO].Cells[6].Value.ToString().Trim();
                }

                if (selectWO.EndsWith("L"))
                {
                    string prefixWO = selectWO.Substring(0, selectWO.Length - 1);
                    matchWO = prefixWO + "R";

                    foreach (DataGridViewRow row in gridWorkorder.Rows)
                    {
                        string wo = row.Cells[2].Value.ToString().Trim();
                        if (matchWO == wo)
                        {
                            matchPartNumner = row.Cells[5].Value.ToString().Trim();
                            row.Selected = true;
                        }
                    }

                    if (gridWorkorder.SelectedRows.Count == 2)
                    {
                        selectPartNumber = selectPartNumber + "," + matchPartNumner;
                        selectWO = selectWO + "," + matchWO;
                    }
                }
                else if (selectWO.EndsWith("R"))
                {
                    string prefixWO = selectWO.Substring(0, selectWO.Length - 1);
                    matchWO = prefixWO + "L";

                    foreach (DataGridViewRow row in gridWorkorder.Rows)
                    {
                        string wo = row.Cells[2].Value.ToString().Trim();
                        if (matchWO == wo)
                        {
                            matchPartNumner = row.Cells[5].Value.ToString().Trim();
                            row.Selected = true;
                        }
                    }

                    if (gridWorkorder.SelectedRows.Count == 2)
                    {
                        selectPartNumber = selectPartNumber + "," + matchPartNumner;
                        selectWO = selectWO + "," + matchWO;
                    }
                }
                //if (!CheckUserSkillByWO(selectWO.Split(new char[] { ',' })[0]))
                //    return;
                this.txbCDAMONumber.Text = selectWO;
                this.txbCDAPartNumber.Text = selectPartNumber;

                //首件检查
                //CheckIPIStatus();

                GetMaterialBinData materialHanlder = new GetMaterialBinData(sessionContext, initModel, this);
                iProcessLayer = materialHanlder.GetProcessLayer(this.txbCDAMONumber.Text);

                PanelQty = GetPCBPanelQty();
                LoadYield();
                if (config.IsMaterialSetup == "Y")
                {
                    InitSetupGrid();
                }
                if (config.IsEquipSetup == "Y")
                {
                    InitEquipmentGrid();
                    SetupMachineAuto();
                }

                InitTab();
                InitCintrolLanguage();
                InitDocumentGrid();
                SetAlarmStatusText("");
                Pscount = 0;
            }
            else
            {
                this.gridWorkorder.ClearSelection();
            }
        }

        #endregion

        #region Other functions
        int workorderTotal = 0;
        private void LoadYield()
        {
            GetProductQuantity getProductHandler = new GetProductQuantity(sessionContext, initModel, this);
            if (!string.IsNullOrEmpty(this.txbCDAMONumber.Text))
            {
                string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
                int totalQty = 0;
                int passQty = 0;
                int failQty = 0;
                int scrapQty = 0;
                foreach (var wo in wos)
                {
                    ProductEntity entity = getProductHandler.GetProductQty(iProcessLayer, wo);
                    if (entity != null)
                    {
                        workorderTotal = Convert.ToInt32(entity.QUANTITY_WORKORDER_TOTAL);
                        passQty += Convert.ToInt32(entity.QUANTITY_PASS);
                        failQty += Convert.ToInt32(entity.QUANTITY_FAIL);
                        scrapQty += Convert.ToInt32(entity.QUANTITY_SCRAP);
                        totalQty = (passQty + failQty + scrapQty);
                        this.lblPass.Text = passQty.ToString();
                        this.lblFail.Text = failQty.ToString();
                        this.lblScrap.Text = scrapQty.ToString();
                        this.lblYield.Text = "0%";
                        if (totalQty > 0)
                        {
                            this.lblYield.Text = Math.Round(Convert.ToDecimal(lblPass.Text) / Convert.ToDecimal(totalQty) * 100, 2) + "%";
                        }
                    }
                }
            }
        }

        private int GetPCBPanelQty()
        {
            int iQty = 0;
            string[] partNos = this.txbCDAPartNumber.Text.Split(new char[] { ',' });
            GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
            List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(partNos[0]);
            if (listData != null && listData.Count > 0)
            {
                MdataGetPartData mData = listData[0];
                iQty = mData.quantityMultipleBoard;
            }
            return iQty;
        }

        private void ShowTopWindow()
        {
            if (topmostform == null)
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
            }
        }

        private void SetTopWindowMessage(string text, string errorMsg)
        {
            if (topmostform != null)
            {
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
            else
            {
                topmostform = new TopMostForm(this);
                topmostHandle = new HandleInterfaceUpdateTopMostDelegate(topmostform.UpdateData);
                topmostform.Show();
                this.Invoke(topmostHandle, new string[] { text, errorMsg });
            }
        }

        private void InitWorkOrder()
        {
            GetWorkOrder getWOHandler = new GetWorkOrder(sessionContext, initModel, this);
            DataTable dt = getWOHandler.GetAllWorkordersExt();
            DataView dv = dt.DefaultView;
            dv.Sort = "Info desc";
            dt = dv.Table;
            if (dt != null)
            {
                this.gridWorkorder.DataSource = dt;
                this.gridWorkorder.ClearSelection();
            }
            for (int i = 0; i < gridWorkorder.Rows.Count; i++)
            {
                gridWorkorder.Rows[i].Cells["columnRunId"].Value = i + 1 + "";
            }

            //DataTable dtBindWO = new DataTable();
            //dtBindWO.Columns.Add("value");
            //dtBindWO.Columns.Add("name");
            //Dictionary<string, string> dicWO = new Dictionary<string, string>();
            //GetWorkOrder getWOHandler = new GetWorkOrder(sessionContext, initModel, this);
            //DataTable dt = getWOHandler.GetAllWorkorders();
            //if (dt != null && dt.Rows.Count > 0)
            //{
            //    foreach (DataRow item in dt.Rows)
            //    {
            //        string strPartNumber = item["PART_NUMBER"].ToString();
            //        string strWorkOrder = item["WorkorderNumber"].ToString();
            //        if (strWorkOrder.EndsWith("L"))
            //        {
            //            string prefixWO = strWorkOrder.Substring(0, strWorkOrder.Length - 1);
            //            string rightWO = prefixWO + "R";
            //            DataRow[] values = dt.Select(string.Format("WorkorderNumber ='{0}' ", rightWO));
            //            if (values.Length > 0)
            //            {
            //                string partNumberR = values[0]["PART_NUMBER"].ToString();
            //                string workOrderR = values[0]["WorkorderNumber"].ToString();
            //                if (!dicWO.ContainsKey(strWorkOrder + "," + rightWO))
            //                {
            //                    dicWO[strWorkOrder + "," + rightWO] = strPartNumber + "," + partNumberR;
            //                }
            //            }
            //        }
            //        else if (!strWorkOrder.EndsWith("L") && !strWorkOrder.EndsWith("R"))
            //        {
            //            if (!dicWO.ContainsKey(strWorkOrder))
            //            {
            //                dicWO[strWorkOrder] = strPartNumber;
            //            }
            //        }
            //    }

            //    foreach (var keyTemp in dicWO.Keys)
            //    {
            //        DataRow row = dtBindWO.NewRow();
            //        row["name"] = keyTemp;
            //        row["value"] = dicWO[keyTemp];
            //        dtBindWO.Rows.Add(row);
            //    }

            //    this.cmbWO.DataSource = dtBindWO;
            //    this.cmbWO.DisplayMember = "name";
            //    this.cmbWO.ValueMember = "value";
            //}
        }

        private void SetupMaterialbySN(string serialNumber, out string workorder)
        {
            //get workorder from serial number
            GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
            string[] snINfo = getSNHandler.GetSNInfo(serialNumber);// "PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER"
            workorder = snINfo[2];
            //deactivated other workorder in the same line
            if (workorder.EndsWith("L"))
            {
                string otherWO = workorder.Substring(0, workorder.Length - 1) + "R";
            }
            else if (workorder.EndsWith("R"))
            {
                string otherWO = workorder.Substring(0, workorder.Length - 1) + "L";
            }
            else
            {
                return;
            }
            setupHandler.SetupStateChange(workorder, iProcessLayer, 2);
            //setup material by wo
            int iCountTemp = 0;
            foreach (var item in dicSetupPPNAndQty.Keys)
            {
                string[] values = item.Split(new char[] { ',' });
                string strWO = values[0];
                string strPPN = values[1];
                if (strWO == workorder)
                {
                    foreach (DataGridViewRow row in gridSetup.Rows)
                    {
                        if (row.Cells["PartNumber"].Value.ToString() == strPPN && (row.Cells["MaterialBinNo"].Value != null && row.Cells["MaterialBinNo"].Value.ToString() != ""))
                        {
                            string materialBinNumber = row.Cells["MaterialBinNo"].Value.ToString();
                            string strActualQty = row.Cells["Qty"].Value.ToString();
                            string strCompName = row.Cells["Comp"].Value.ToString();
                            setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, workorder, materialBinNumber, strActualQty, strPPN, config.StationNumber + iCountTemp, strCompName);
                            setupHandler.SetupStateChange(workorder, iProcessLayer, 0);
                            iCountTemp++;
                        }
                    }
                }
            }
        }

        private void SetupMaterialByWO(string workorder)
        {
            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
            //deactivated other workorder in the same line
            if (workorder.EndsWith("L"))
            {
                string otherWO = workorder.Substring(0, workorder.Length - 1) + "R";
            }
            else if (workorder.EndsWith("R"))
            {
                string otherWO = workorder.Substring(0, workorder.Length - 1) + "L";
            }
            else
            {
                return;
            }
            setupHandler.SetupStateChange(workorder, iProcessLayer, 2);
            //setup material by wo
            int iCountTemp = 0;
            foreach (var item in dicSetupPPNAndQty.Keys)
            {
                string[] values = item.Split(new char[] { ',' });
                string strWO = values[0];
                string strPPN = values[1];
                if (strWO == workorder)
                {
                    foreach (DataGridViewRow row in gridSetup.Rows)
                    {
                        if (row.Cells["PartNumber"].Value.ToString() == strPPN && (row.Cells["MaterialBinNo"].Value != null && row.Cells["MaterialBinNo"].Value.ToString() != ""))
                        {
                            string materialBinNumber = row.Cells["MaterialBinNo"].Value.ToString();
                            string strActualQty = row.Cells["Qty"].Value.ToString();
                            string strCompName = row.Cells["Comp"].Value.ToString();
                            setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, workorder, materialBinNumber, strActualQty, strPPN, config.StationNumber + iCountTemp, strCompName);
                            setupHandler.SetupStateChange(workorder, iProcessLayer, 0);
                            iCountTemp++;
                        }
                    }
                }
            }
        }

        private void ProcessMaterialBinNo(string materialBinNo)
        {
            GetMaterialBinData getMaterialHandler = new GetMaterialBinData(sessionContext, initModel, this);
            string[] values = getMaterialHandler.GetMaterialBinDataDetails(materialBinNo);
            //"MATERIAL_BIN_NUMBER", "MATERIAL_BIN_PART_NUMBER", "MATERIAL_BIN_QTY_ACTUAL", "MATERIAL_BIN_QTY_TOTAL", "PART_DESC", "MSL_FLOOR_LIFETIME_REMAIN" ,"EXPIRATION_DATE"
            if (values != null && values.Length > 0)
            {
                string strPartNumber = values[1];
                string strActualQty = values[2];
                string strLifeTime = values[5];
                string lockState = values[7];
                DateTime dtExpiry = Convert.ToDateTime(ConvertDateFromStamp(values[6]));

                if (lockState == "-1")
                {
                    errorHandler(2, message("TheContainerIsLocked"), "");
                    return;
                }
                //if (!VerifyActivatedWO())
                //    return;
                if (dtExpiry < DateTime.Now)
                {
                    errorHandler(2, message("msg_The solder paste has expiry."), "");
                    return;
                }
                int iCountTemp = 0;
                foreach (DataGridViewRow row in gridSetup.Rows)
                {
                    if (row.Cells["PartNumber"].Value.ToString() == strPartNumber)
                    {
                        if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString() == "")
                        {
                            //setup material
                            string compName = row.Cells["Comp"].Value.ToString();
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
                            if (wos.Length == 1)
                            {
                                setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, this.txbCDAMONumber.Text, materialBinNo, strActualQty, strPartNumber, config.StationNumber + iCountTemp, compName + iCountTemp);
                                setupHandler.SetupStateChange(this.txbCDAMONumber.Text, iProcessLayer, 0);
                            }
                            row.Cells["MaterialBinNo"].Value = materialBinNo;
                            row.Cells["Qty"].Value = Convert.ToDouble(strActualQty);
                            row.Cells["MScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            row.Cells["ExpiryTime"].Value = dtExpiry.ToString("yyyy/MM/dd HH:mm:ss");
                            row.Cells["MSL"].Value = ConverToHourAndMin(Convert.ToInt32("100"));
                            row.Cells["Checked"].Value = ICTClient.Properties.Resources.ok;
                            row.Cells["MaterialBinNo"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            SetTipMessage(MessageType.OK, message("Process material bin number") + materialBinNo + message("SUCCESS"));
                            iCountTemp++;
                            break;
                        }
                        else
                        {
                            if (CheckMaterialBinHasSetup(materialBinNo))
                                return;
                            this.Invoke(new MethodInvoker(delegate
                            {
                                gridSetup.Rows.InsertCopy(row.Index, row.Index + 1);
                                DataGridViewRow newRow = gridSetup.Rows[row.Index + 1];
                                newRow.Cells["Checked"].Value = ICTClient.Properties.Resources.ok;
                                newRow.Cells["MaterialBinNo"].Value = materialBinNo;
                                newRow.Cells["PartNumber"].Value = row.Cells["PartNumber"].Value;
                                newRow.Cells["PartDesc"].Value = row.Cells["PartDesc"].Value;
                                newRow.Cells["Qty"].Value = Convert.ToDouble(strActualQty);
                                newRow.Cells["Comp"].Value = row.Cells["Comp"].Value;
                                newRow.Cells["MSL"].Value = ConverToHourAndMin(Convert.ToInt32("100"));
                                newRow.Cells["MaterialBinNo"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                newRow.Cells["MScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                                newRow.Cells["ExpiryTime"].Value = dtExpiry.ToString("yyyy/MM/dd HH:mm:ss");
                            }));
                            break;
                        }
                    }
                }
            }
        }

        private void ProcessEquipmentData(string equipmentNo)
        {
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            if (!CheckEquipmentValid(values))
            {
                errorHandler(3, message("The equipment is invalid"), "");
                return;
            }
            string eqPartNumber = values[2];
            //check equipment number  whether need to setup?
            if (!CheckEquipmentIsExist(eqPartNumber, equipmentNo))
                return;

            //find the equipment should setup on which wo
            string eqWorkOrder = CheckEquipmentSetupWO(eqPartNumber);
            if (!string.IsNullOrEmpty(eqWorkOrder))
            {
                int errorCode = eqManager.UpdateEquipmentData(eqWorkOrder, iProcessLayer, equipmentNo, 0);//todo
                if (errorCode == 0)//1301 Equipment is already set up
                {
                    EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                    if (entityExt != null)
                    {
                        entityExt.PART_NUMBER = eqPartNumber;
                        entityExt.EQ_WORKORDER = eqWorkOrder;
                        SetEquipmentGridData(entityExt);
                        SetTipMessage(MessageType.OK, message("Process equipment number") + equipmentNo + message("SUCCESS"));
                    }
                }
            }
            else
            {
                SetTipMessage(MessageType.OK, message("The equipment is invalid"));
            }
        }
        private bool ProcessSerialNumberSingle(string serialNumber)
        {
            //verify material&equipment is ok
            if (this.dgvEquipment.RowCount > 0)
            {
                if (!VerifyEquipment())
                {
                    return false;
                }
            }
            if (config.IsSelectWO == "Y" || this.txbCDAMONumber.Text != "")
            {
                if (!VerifySerialNumberByWO(serialNumber))
                {
                    return false;
                }
            }
            else
            {
                InitProcessLayer(serialNumber);
            }

            //check serial state
            bool b = true;
            CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
            int isOK = checkHandler.trCheckSNStateNextStep(serialNumber, iProcessLayer);
            if (isOK == 0)
            {
                SetTopWindowMessage(serialNumber, "");
                CheckIPIStatus();
                this.Invoke(new MethodInvoker(delegate
                {
                    b = LoadVICSingle(serialNumber);

                }));
                if (!b)
                {
                    return false;
                }
            }
            else
            {
                SetTopWindowMessage(serialNumber, message("Check Serial Number State Error"));
                return false;
            }
            return true;
        }
        private bool ProcessSerialNumber(string serialNumber)
        {
            try
            {
                if (dgvSN.RowCount > 0)
                {
                    int columnIndex = 0;
                    int rowIndex = -1;
                    dgvSN.ClearSelection();
                    foreach (DataGridViewRow row in this.dgvSN.Rows)
                    {
                        if (row.Cells[1].Value.ToString() == serialNumber)
                        {
                            rowIndex = row.Index;
                            row.Selected = true;
                            break;
                        }
                    }
                    if (rowIndex != -1)
                    {
                        DataGridViewCellEventArgs e = new DataGridViewCellEventArgs(columnIndex, rowIndex);
                        this.Invoke(new MethodInvoker(delegate
                        {
                            dgvSN_CellDoubleClick(null, e);
                        }));
                        return true;
                    }
                    else
                    {
                        //verify material&equipment is ok
                        if (this.dgvEquipment.RowCount > 0)
                        {
                            if (!VerifyEquipment())
                            {
                                return false;
                            }
                        }
                        if (config.IsSelectWO == "Y" || this.txbCDAMONumber.Text != "")
                        {
                            if (!VerifySerialNumberByWO(serialNumber))
                            {
                                return false;
                            }
                        }
                        else
                        {
                            InitProcessLayer(serialNumber);
                        }

                        //check serial state
                        bool b = true;
                        CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
                        int isOK = checkHandler.trCheckSNStateNextStep(serialNumber, iProcessLayer);
                        if (isOK == 0)
                        {
                            SetTopWindowMessage(serialNumber, "");
                            CheckIPIStatus();
                            this.Invoke(new MethodInvoker(delegate
                            {
                                b = LoadVIC(serialNumber);
                            }));
                            if (!b)
                            {
                                return false;
                            }
                        }
                        else
                        {
                            SetTopWindowMessage(serialNumber, message("Check Serial Number State Error"));
                            return false;
                        }
                        return true;
                    }
                }
                else
                {
                    //verify material&equipment is ok
                    if (this.dgvEquipment.RowCount > 0)
                    {
                        if (!VerifyEquipment())
                        {
                            return false;
                        }
                    }
                    if (config.IsSelectWO == "Y" || this.txbCDAMONumber.Text != "")
                    {
                        if (!VerifySerialNumberByWO(serialNumber))
                        {
                            return false;
                        }
                    }
                    else
                    {
                        InitProcessLayer(serialNumber);
                    }

                    //check serial state
                    bool b = true;
                    CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
                    int isOK = checkHandler.trCheckSNStateNextStep(serialNumber, iProcessLayer);
                    if (isOK == 0)
                    {
                        SetTopWindowMessage(serialNumber, "");
                        CheckIPIStatus();
                        this.Invoke(new MethodInvoker(delegate
                        {
                            b = LoadVIC(serialNumber);
                        }));
                        if (!b)
                        {
                            return false;
                        }
                    }
                    else
                    {
                        SetTopWindowMessage(serialNumber, message("Check Serial Number State Error"));
                        return false;
                    }
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }

        }

        Dictionary<int, string> dicSNAndPos = new Dictionary<int, string>();
        private void GetSerialNumberAndPos(string serialNumberRef)
        {
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
            {
                errorHandler(2, message("Please select a workorder"), "");
                return;
            }

            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            if (wos.Length > 1)
            {
                string strSNRefL = serialNumberRef + "L";
                string strSNRefR = serialNumberRef + "R";
                GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                SerialNumberData[] snArrayL = getSNHandler.GetSerialNumberBySNRef(strSNRefL);
                SerialNumberData[] snArrayR = getSNHandler.GetSerialNumberBySNRef(strSNRefR);
                dicSNAndPos.Clear();
                foreach (var item in snArrayL)
                {
                    string snTemp = item.serialNumber;
                    int lindex = snTemp.LastIndexOf("L");
                    string pos = snTemp.Substring(lindex - 2, 2);
                    if (!dicSNAndPos.ContainsKey(Convert.ToInt32(pos)))
                        dicSNAndPos[Convert.ToInt32(pos)] = snTemp;
                }
                foreach (var item in snArrayR)
                {
                    string snTemp = item.serialNumber;
                    int lindex = snTemp.LastIndexOf("R");
                    string pos = snTemp.Substring(lindex - 2, 2);
                    if (!dicSNAndPos.ContainsKey(Convert.ToInt32(pos)))
                        dicSNAndPos[Convert.ToInt32(pos)] = snTemp;
                }
            }
            else
            {
                GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                SerialNumberData[] snArray = getSNHandler.GetSerialNumberBySNRef(serialNumberRef);
                foreach (var item in snArray)
                {
                    string snTemp = item.serialNumber;
                    string pos = item.serialNumberPos;
                    if (!dicSNAndPos.ContainsKey(Convert.ToInt32(pos)))
                        dicSNAndPos[Convert.ToInt32(pos)] = snTemp;
                }
            }
        }

        private bool AssignSNToWO(string serialNumber)
        {
            //assign serial number to work order
            //150719000101L1, 150719000103L2, 150719000105L3  
            //150719000102R1, 150719000104R2, 150719000106R3 
            List<SerialNumberData> snListL = new List<SerialNumberData>();
            List<SerialNumberData> snListR = new List<SerialNumberData>();
            for (int i = 1; i <= PanelQty; i++)
            {
                string serialNumberL = serialNumber + (2 * i - 1).ToString().PadLeft(2, '0') + "L" + i;
                SerialNumberData snDataL = new SerialNumberData();
                snDataL.serialNumber = serialNumberL;
                snDataL.serialNumberPos = i + "";
                snListL.Add(snDataL);
                string serialNumberR = serialNumber + (2 * i).ToString().PadLeft(2, '0') + "R" + i;
                SerialNumberData snDataR = new SerialNumberData();
                snDataR.serialNumber = serialNumberR;
                snDataR.serialNumberPos = i + "";
                snListR.Add(snDataR);
            }
            int error1 = 0;
            int error2 = 0;
            int error3 = 0;
            AssignSerialNumber assignSNHandler = new AssignSerialNumber(sessionContext, initModel, this);
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            foreach (var wo in wos)
            {
                if (wo.EndsWith("L"))
                {
                    error1 = assignSNHandler.AssignSerialNumberResultCall(serialNumber + "L", snListL.ToArray(), wo, iProcessLayer);
                }
                else if (wo.EndsWith("R"))
                {
                    error2 = assignSNHandler.AssignSerialNumberResultCall(serialNumber + "R", snListR.ToArray(), wo, iProcessLayer);
                }
                else
                {
                    error3 = assignSNHandler.AssignSerialNumberResultCall(serialNumber, new SerialNumberData[] { }, wo, iProcessLayer);
                }
            }
            if (error1 == 0 && error2 == 0 && error3 == 0)
            {
                return true;
            }
            else
            {
                errorHandler(2, message("Assign serial number error"), "");
                return false;
            }

        }

        #region add by qy
        private bool AssignSNToWOSingle(string serialNumber)
        {
            //assign serial number to work order

            int error1 = 0;
            AssignSerialNumber assignSNHandler = new AssignSerialNumber(sessionContext, initModel, this);
            string wo = this.txbCDAMONumber.Text;
            error1 = assignSNHandler.AssignSerialNumberResultCall(serialNumber, new SerialNumberData[] { }, wo, iProcessLayer);

            if (error1 == 0)
            {
                return true;
            }
            else
            {
                errorHandler(2, message("Assign serial number error"), "");
                return false;
            }

        }

        string serialnumberSlave = "";
        string serialnumberMaster = "";
        private bool MergeSN()
        {
            bool iResult = true;

            string productflagMaster = "";
            string productflagSlave = "";
            GetMaterialBinData materialHandler = new GetMaterialBinData(sessionContext, initModel, this);

            DataTable dt = materialHandler.GetBomMaterialDataBySN(serialnumberMaster);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    productflagMaster = row["ProductFlag"].ToString();
                }
            }
            else
            {
                iResult = false;
            }

            DataTable dt2 = materialHandler.GetBomMaterialDataBySN(serialnumberSlave);
            if (dt2 != null && dt2.Rows.Count > 0)
            {
                foreach (DataRow row in dt2.Rows)
                {
                    productflagSlave = row["ProductFlag"].ToString();
                }
            }
            else
            {
                iResult = false;
            }

            if (productflagSlave == "1" && productflagMaster == "1")
            {
                MergeManger mergeHandler = new MergeManger(sessionContext, initModel, this);
                //trVerifyMerge
                int error1 = mergeHandler.VerifyMerge(serialnumberSlave, serialnumberMaster);
                if (error1 != 0)
                {
                    iResult = false;
                }
                else
                {
                    //trMergeParts
                    int error2 = mergeHandler.MergeSerialNumber(serialnumberMaster, serialnumberSlave, iProcessLayer);
                    if (error2 != 0)
                    {
                        iResult = false;
                    }
                }
            }

            return iResult;
        }
        #endregion

        private bool SendSN(string serialNumber)
        {
            try
            {
                //Thread.Sleep(300);
                if (config.DataOutputInterface == "COM")
                {
                    try
                    {
                        initModel.scannerHandler.OutputCOM().Write(strToToHexByte(serialNumber), 0, strToToHexByte(serialNumber).Length);
                        LogHelper.Info("Send command:" + serialNumber);
                        return true;
                    }
                    catch (Exception e)
                    {
                        LogHelper.Error(e);
                        return false;
                    }
                }
                else
                {
                    if (config.OutputEnter == "1")
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber + "\r"); //大写键总是被按起。。。。
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber + "\r");
                        }
                    }
                    else
                    {
                        if (Control.IsKeyLocked(Keys.CapsLock))
                        {
                            SendKeys.SendWait("{CAPSLOCK}" + serialNumber);
                        }
                        else
                        {
                            SendKeys.SendWait(serialNumber);
                        }
                    }
                    SendKeys.Flush();
                    //Thread.Sleep(300);
                    return true;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
                return false;
            }
        }

        private void ConsumptionMaterial()
        {
            foreach (var item in dicSetupPPNAndQty.Keys)
            {
                string[] values = item.Split(new char[] { ',' });
                string strWO = values[0];
                string strPPN = values[1];
                string strQty = dicSetupPPNAndQty[item];
                string materialBinNo = FindMaterialBinNumberExt(strPPN);
                UpdateMaterialGridData(strWO, strPPN, materialBinNo, Convert.ToDouble(strQty) * PanelQty);
            }
        }

        private string FindMaterialBinNumberExt(string partNumber)
        {
            string materialBinNo = "";
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (gridSetup.Rows[i].Cells["PartNumber"].Value.ToString() == partNumber && Convert.ToInt32(gridSetup.Rows[i].Cells["Qty"].Value) > 0)
                {
                    materialBinNo = gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString();
                    break;
                }
            }
            return materialBinNo;
        }

        private void UpdateGridDataAfterUploadState()
        {
            foreach (DataGridViewRow row in dgvEquipment.Rows)
            {
                if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length > 0)
                {
                    int iQty = Convert.ToInt32(row.Cells["UsCount"].Value.ToString());
                    row.Cells["UsCount"].Value = iQty - PanelQty;//
                }
            }
            foreach (DataGridViewRow row in gridMachine.Rows)
            {
                if (row.Cells["UsCountMac"].Value != null && row.Cells["UsCountMac"].Value.ToString().Length > 0)
                {
                    int iQty = Convert.ToInt32(row.Cells["UsCountMac"].Value.ToString());
                    row.Cells["UsCountMac"].Value = iQty - PanelQty;//
                }
            }
            ConsumptionMaterial();
        }

        private void UpdateMaterialGridData(string wo, string ppn, string materialBinNumber, double qty)
        {
            ProcessMaterialBinData materialHandler = new ProcessMaterialBinData(sessionContext, initModel, this);
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString() == materialBinNumber)
                {
                    double iQty = Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value);
                    if (iQty >= qty)
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = iQty - qty;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, -qty);
                        if (iQty == qty)//update 2015/6/24
                        {
                            if (i + 1 < this.gridSetup.Rows.Count)
                            {
                                string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                                string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                                string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                                if (nextPartNumber != ppn)
                                {
                                    errorHandler(3, message("The material not enough"), "");
                                }
                            }
                            else
                            {
                                errorHandler(3, message("The material not enough"), "");
                            }
                        }
                        break;
                    }
                    else
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = 0;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, -iQty);
                        if (i + 1 < this.gridSetup.Rows.Count && gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString() == ppn)
                        {
                            string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                            string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                            string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                            string compName = gridSetup.Rows[i + 1].Cells["Comp"].Value.ToString();
                            //setup material
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, wo, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + compName, compName);
                            setupHandler.SetupStateChange(wo, iProcessLayer, 0);
                            UpdateMaterialGridData(wo, ppn, nextMaterialBinNo, qty - iQty);
                        }
                        else
                        {
                            //warming no material
                            errorHandler(3, message("The material not enough"), "");
                        }
                    }
                }
            }
        }
        private void UpdateMaterialGridData(string materialBinNumber, double qty)
        {
            ProcessMaterialBinData materialHandler = new ProcessMaterialBinData(sessionContext, initModel, this);
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString() == materialBinNumber)
                {
                    double iQty = Convert.ToDouble(gridSetup.Rows[i].Cells["Qty"].Value);
                    if (iQty >= qty)
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = iQty - qty;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, -qty);
                        if (iQty == qty)//update 2015/6/24
                        {
                            if (i + 1 < this.gridSetup.Rows.Count)
                            {
                                string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                                string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                                string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                                //setup material
                                SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                                setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                                setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0, iProcessLayer);
                            }
                            else
                            {
                                errorHandler(3, message("The bare board not enough"), "");
                            }
                        }
                        break;
                    }
                    else
                    {
                        gridSetup.Rows[i].Cells["Qty"].Value = 0;
                        int errorMaterial = materialHandler.UpdateMaterialBinBooking(materialBinNumber, -iQty);
                        if (i + 1 < this.gridSetup.Rows.Count)
                        {
                            string nextMaterialBinNo = gridSetup.Rows[i + 1].Cells["MaterialBinNo"].Value.ToString();
                            string nextPartNumber = gridSetup.Rows[i + 1].Cells["PartNumber"].Value.ToString();
                            string nextQty = gridSetup.Rows[i + 1].Cells["Qty"].Value.ToString();
                            //setup material
                            SetUpManager setupHandler = new SetUpManager(sessionContext, initModel, this);
                            setupHandler.UpdateMaterialSetUpByBin(iProcessLayer, this.txbCDAMONumber.Text, nextMaterialBinNo, nextQty, nextPartNumber, config.StationNumber + "_01", "01");
                            setupHandler.SetupStateChange(this.txbCDAMONumber.Text, 0, iProcessLayer);
                            UpdateMaterialGridData(nextMaterialBinNo, qty - iQty);
                        }
                        else
                        {
                            //warming no material
                            errorHandler(3, message("The bare board not enough"), "");
                        }
                    }
                }
            }
        }

        private string FindMaterialBinNumber()
        {
            string materialBinNo = "";
            for (int i = 0; i < this.gridSetup.Rows.Count; i++)
            {
                if (Convert.ToInt32(gridSetup.Rows[i].Cells["Qty"].Value) > 0)
                {
                    materialBinNo = gridSetup.Rows[i].Cells["MaterialBinNo"].Value.ToString();
                    break;
                }

            }
            return materialBinNo;
        }

        private bool CheckMaterialBinHasSetup(string materailBinNo)
        {
            bool isExist = false;
            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value.ToString() == materailBinNo)
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }

        private bool VerifyEquipment()
        {
            bool isValid = true;
            EquipmentManager equipmentHandler = new EquipmentManager(sessionContext, initModel, this);
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            foreach (var wo in wos)
            {
                int errorCode = equipmentHandler.CheckEquipmentData(wo, iProcessLayer);
                if (errorCode != 0)
                {
                    errorHandler(3, "Check equipment data error :" + errorCode + ",workorder :" + wo, "");
                    return false;
                }
            }

            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                if (Convert.ToInt32(item.Cells["UsCount"].Value) <= 0)
                {
                    isValid = false;
                    item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                    errorHandler(3, message("Stencil usage count cann not less then 0"), "");
                    break;
                }
                else if (Convert.ToDateTime(item.Cells["NextMaintenance"].Value) <= DateTime.Now)
                {
                    isValid = false;
                    item.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(255, 255, 255);
                    errorHandler(3, message("Equipment need to maintenance"), "");
                    break;
                }
            }
            return isValid;
        }

        private bool VerifyActivatedWO()
        {
            bool isValid = true;
            GetCurrentWorkorder getActivatedWOHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
            GetStationSettingModel stationSetting = getActivatedWOHandler.GetCurrentWorkorderResultCall();
            if (stationSetting != null && stationSetting.workorderNumber != null)
            {
                if (stationSetting.workorderNumber == this.txbCDAMONumber.Text)
                {
                    isValid = true;
                }
                else
                {
                    isValid = false;
                    errorHandler(2, message("The current activated work order has changed please refresh"), "");
                }
            }
            return isValid;
        }

        //verify the serial number's part number is equals the current part number
        private bool VerifySerialNumber(string serialNumber)
        {
            bool isValid = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                string snPartNumber = snValues[1];
                if (this.txbCDAPartNumber.Text.Trim().Contains(snPartNumber))
                { }
                else
                {
                    errorHandler(3, message("The serial number part number is not equals the current part number"), "");
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, "The serial number " + serialNumber + " is invalid", "");
                SetTopWindowMessage(serialNumber, message("The serial number is invalid"));
                isValid = false;
            }
            return isValid;
        }

        private bool VerifySerialNumberByWO(string serialNumber)
        {
            bool isValid = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            if (wos.Length > 1)
            {
                serialNumber = serialNumber + "L";
            }
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                string snWorkOrder = snValues[2];
                if (this.txbCDAMONumber.Text.Trim().Contains(snWorkOrder))
                {
                    partdesc = snValues[0];
                }
                else
                {
                    errorHandler(3, message("WO_NotMatch"), "");
                    SetTopWindowMessage(serialNumber, message("WO_NotMatch"));
                    isValid = false;
                }
            }
            else
            {
                errorHandler(3, message("The serial number is invalid") + serialNumber, "");
                SetTopWindowMessage(serialNumber, message("The serial number is invalid"));
                isValid = false;
            }
            return isValid;
        }


        private string ConvertDateFromStamp(string timeStamp)
        {
            double d = Convert.ToDouble(timeStamp);
            DateTime start = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc);
            DateTime date = start.AddMilliseconds(d).ToLocalTime();
            return date.ToString();
        }

        private long ConvertDateToStamp(DateTime dt)
        {
            DateTime dtStart = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            TimeSpan toNow = dt.Subtract(dtStart);
            return Convert.ToInt64(toNow.TotalMilliseconds);
        }

        private string ConverToHourAndMin(int number)
        {
            int iHour = number / 60;
            int iMin = number % 60;
            return iHour + "hr " + iMin + "min";
        }

        private bool CheckMaterialSetUp()
        {
            bool isValid = true;

            foreach (DataGridViewRow row in gridSetup.Rows)
            {
                if (row.Cells["MaterialBinNo"].Value == null || row.Cells["MaterialBinNo"].Value.ToString().Length == 0)
                {
                    errorHandler(3, message("Material setup required"), "");
                    isValid = false;
                    break;
                }
            }
            return isValid;
        }

        public string byteToHexStr(byte[] bytes)
        {
            string returnStr = "";
            if (bytes != null)
            {
                for (int i = 0; i < bytes.Length; i++)
                {
                    returnStr += bytes[i].ToString("X2");
                }
            }
            return returnStr;
        }

        #region Compare
        private class RowComparer : System.Collections.IComparer
        {
            private static int sortOrderModifier = 1;

            public RowComparer(SortOrder sortOrder)
            {
                if (sortOrder == SortOrder.Descending)
                {
                    sortOrderModifier = -1;
                }
                else if (sortOrder == SortOrder.Ascending) { sortOrderModifier = 1; }
            }
            public int Compare(object x, object y)
            {
                DataGridViewRow DataGridViewRow1 = (DataGridViewRow)x;
                DataGridViewRow DataGridViewRow2 = (DataGridViewRow)y;
                // Try to sort based on the Scan time column.
                string value1 = DataGridViewRow1.Cells["colItemCode"].Value.ToString();
                string value2 = DataGridViewRow2.Cells["colItemCode"].Value.ToString();
                string type1 = DataGridViewRow1.Cells["colType"].Value.ToString();
                string type2 = DataGridViewRow2.Cells["colType"].Value.ToString();
                int CompareResult = 0;
                if (type1 == type2)
                {
                    CompareResult = value1.CompareTo(value2);
                }
                else
                {
                    CompareResult = type1.CompareTo(type2);
                }
                return CompareResult * sortOrderModifier;
            }
        }
        #endregion

        #region Equipment
        private void InitTab()
        {
            if (this.gridSetup.RowCount == 0)
            {
                this.tabSetup.Parent = null;
            }
            else
            {
                this.tabSetup.Parent = this.tabControl1;
            }
            if (this.gridMachine.RowCount == 0)
            {
                this.tabMachine.Parent = null;
            }
            else
            {
                this.tabMachine.Parent = this.tabControl1;
            }
            if (this.dgvEquipment.RowCount == 0)
            {
                this.tabEquipment.Parent = null;
            }
            else
            {
                this.tabEquipment.Parent = this.tabControl1;
            }
        }

        Dictionary<string, string> dicSetupPPNAndQty = null;
        List<string> productPN = null;
        private void InitSetupGrid()
        {
            this.gridSetup.Rows.Clear();
            dicSetupPPNAndQty = new Dictionary<string, string>();
            productPN = new List<string>();
            GetMaterialBinData getMaterial = new GetMaterialBinData(sessionContext, initModel, this);
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            List<string> matList = new List<string>();
            foreach (var wo in wos)
            {
                DataTable dt = getMaterial.GetBomMaterialData(wo);
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["ProductFlag"].ToString() == "1")
                        {
                            productPN.Add(row["PartNumber"].ToString());
                        }
                        else
                        {
                            string strPPN = wo + "," + row["PartNumber"].ToString();
                            if (!dicSetupPPNAndQty.ContainsKey(strPPN))
                            {
                                dicSetupPPNAndQty[strPPN] = row["Quantity"].ToString();
                            }
                            if (!matList.Contains(row["PartNumber"].ToString()))
                            {
                                this.gridSetup.Rows.Add(new object[10] { ICTClient.Properties.Resources.Close, "", row["PartNumber"], row["PartDesc"], "", "", row["CompName"], row["Quantity"], "", "" });
                                matList.Add(row["PartNumber"].ToString());
                            }
                        }

                    }
                }
            }
            this.gridSetup.ClearSelection();
        }

        private void InitMachineGrid()
        {
            this.gridMachine.Rows.Clear();
            bool isOK = true;
            string equipmentNo = config.StationNumber;
            if (eqPartNumber == null || equipmentNo.Length == 0)
                return;
            //verify the machine equipment
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            if (values == null || values.Length == 0)
                isOK = false;
            if (values[0] != "0")
                isOK = false;
            if (!isOK)
            {
                errorHandler(3, message("The equipment is invalid"), "");
                this.gridMachine.Rows.Add(new object[7] { ICTClient.Properties.Resources.Close, equipmentNo, equipmentNo, "", "", "", "" });
            }
            else
            {
                this.gridMachine.Rows.Add(new object[7] { ICTClient.Properties.Resources.Close, equipmentNo, values[3], "", "", "", "" });
            }
            //setup machine equipment
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            int errorCode = eqManager.UpdateEquipmentData(wos[0], iProcessLayer, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    entityExt.PART_NUMBER = values[2];
                    SetMachineGridData(entityExt);
                    errorHandler(0, message("Process machine equipment number") + equipmentNo + message("SUCCESS"), "");
                }
            }
        }

        string startafter = "";
        private void InitAgingAttribute()
        {
            GetWorkPlanData workplanHandle = new GetWorkPlanData(sessionContext, initModel, this);
            string[] workorders = this.txbCDAMONumber.Text.Trim().Split(',');
            string[] workplanDataResultValues = workplanHandle.GetWorkplanDataForStation(workorders[0]);
            string workplanid = workplanDataResultValues[2];
            string workstep = workplanDataResultValues[3];

            GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
            string[] attributeResultValues = getattribute.GetAttributeValueForWorkStep("START_AFTER", workplanid, workstep);
            if (attributeResultValues.Length > 0)
            {
                startafter = attributeResultValues[1];
            }
        }
        string partdesc = "";
        private bool InitProcessLayer(string serialNumber)
        {
            bool isValiad = true;
            GetSerialNumberInfo getSNInfoHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
            string[] snValues = getSNInfoHandler.GetSNInfo(serialNumber);//"PART_DESC", "PART_NUMBER", "WORKORDER_NUMBER" 
            if (snValues != null && snValues.Length > 0)
            {
                this.Invoke(new MethodInvoker(delegate
                {
                    this.txbCDAPartNumber.Text = snValues[1];
                    this.txbCDAMONumber.Text = snValues[2];
                    partdesc = snValues[0];
                }));
                InitDocumentGrid();
                //CheckIPIStatus();
            }
            else
            {
                errorHandler(3, message("The serial number is invalid") + serialNumber, "");
                SetTopWindowMessage(serialNumber, message("The serial number is invalid"));
                isValiad = false;
            }

            if (isValiad)
            {
                GetMaterialBinData materialHanlder = new GetMaterialBinData(sessionContext, initModel, this);

                iProcessLayer = materialHanlder.GetProcessLayer(this.txbCDAMONumber.Text);
                this.txtLayer.Text = ConvertProcessLayerToString2(iProcessLayer.ToString());
            }

            return isValiad;
        }

        List<string> equipmentPNList = null;
        private void InitEquipmentGrid()
        {
            dgvEquipment.Rows.Clear();
            equipmentPNList = new List<string>();
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            if (string.IsNullOrEmpty(this.txbCDAMONumber.Text))
                return;
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            foreach (var wo in wos)
            {
                List<EquipmentEntity> listEntity = eqManager.GetRequiredEquipmentData(wo, iProcessLayer);//todo
                if (listEntity != null)
                {
                    foreach (var item in listEntity)
                    {
                        if (!equipmentPNList.Contains(item.PART_NUMBER))
                        {
                            this.dgvEquipment.Rows.Add(new object[8] { ICTClient.Properties.Resources.Close, wo, item.PART_NUMBER, item.EQUIPMENT_DESCRIPTION, "", "", "", "" });
                            equipmentPNList.Add(item.PART_NUMBER);
                        }
                    }
                }
            }
            this.dgvEquipment.ClearSelection();
        }

        private void SetupMachineAuto()
        {
            bool isOK = false;
            string equipmentNo = config.StationNumber;
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipmentNo);
            if (values == null || values.Length == 0)
                isOK = false;
            return;
            if (values[0] != "0")
                isOK = false;
            return;
            //setup machine equipment
            string[] wos = this.txbCDAMONumber.Text.Split(new char[] { ',' });
            int errorCode = eqManager.UpdateEquipmentData(wos[0], iProcessLayer, equipmentNo, 0);
            if (errorCode == 0)//1301 Equipment is already set up
            {
                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                if (entityExt != null)
                {
                    string eqWorkOrder = CheckEquipmentSetupWO(values[2]);
                    entityExt.PART_NUMBER = values[2];
                    entityExt.EQ_WORKORDER = eqWorkOrder;
                    SetEquipmentGridData(entityExt);
                    errorHandler(0, message("Process machine equipment number") + equipmentNo + message("SUCCESS"), "");
                }
            }
        }

        private bool CheckEquipmentSetup()
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["UsCount"].Value != null && row.Cells["UsCount"].Value.ToString().Length == 0)
                {
                    errorHandler(3, message("Equipment setup required"), "");
                    return false;
                }
            }
            return true;
        }

        private bool CheckMachineSetup()
        {
            foreach (DataGridViewRow row in this.gridMachine.Rows)
            {
                if (row.Cells["UsCountMac"].Value != null && row.Cells["UsCountMac"].Value.ToString().Length == 0)
                {
                    errorHandler(3, message("Equipment setup required"), "");
                    return false;
                }
                if (Convert.ToInt32(row.Cells["UsCountMac"].Value.ToString()) <= 0)
                {
                    errorHandler(3, message("Equipment usage count is not enough"), "");
                    return false;
                }
                if (row.Cells["NextMaintenanceMac"].Value != null && Convert.ToDateTime(row.Cells["NextMaintenanceMac"].Value) <= DateTime.Now)
                {
                    errorHandler(3, message("Equipment maintenance time is come"), "");
                    return false;
                }
            }
            return true;
        }

        private void SetEquipmentGridData(EquipmentEntityExt entityExt)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == entityExt.PART_NUMBER
                    && row.Cells["EQWorkorder"].Value.ToString() == entityExt.EQ_WORKORDER
                    && (row.Cells["EquipNo"].Value == null || row.Cells["EquipNo"].Value.ToString() == ""))
                {
                    row.Cells["NextMaintenance"].Value = DateTime.Now.AddSeconds(Convert.ToDouble(entityExt.SECONDS_BEFORE_EXPIRATION)).ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["ScanTime"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["UsCount"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                    row.Cells["EquipNo"].Value = entityExt.EQUIPMENT_NUMBER;
                    row.Cells["Status"].Value = ICTClient.Properties.Resources.ok;
                    row.Cells["eqPartNumber"].Style.BackColor = Color.FromArgb(0, 192, 0);
                }
            }
        }

        private void SetMachineGridData(EquipmentEntityExt entityExt)
        {
            foreach (DataGridViewRow row in this.gridMachine.Rows)
            {
                if (row.Cells["eqPartNumberMac"].Value != null && row.Cells["eqPartNumberMac"].Value.ToString() == entityExt.PART_NUMBER
                    && (row.Cells["EquipNoMac"].Value == null || row.Cells["EquipNoMac"].Value.ToString() == ""))
                {
                    row.Cells["NextMaintenanceMac"].Value = DateTime.Now.AddSeconds(Convert.ToDouble(entityExt.SECONDS_BEFORE_EXPIRATION)).ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["ScanTimeMac"].Value = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                    row.Cells["UsCountMac"].Value = entityExt.USAGES_BEFORE_EXPIRATION;
                    row.Cells["EquipNoMac"].Value = entityExt.EQUIPMENT_NUMBER;
                    row.Cells["StatusMac"].Value = ICTClient.Properties.Resources.ok;
                    row.Cells["eqPartNumberMac"].Style.BackColor = Color.FromArgb(0, 192, 0);
                }
            }
            this.gridMachine.ClearSelection();
        }

        private string CheckEquipmentSetupWO(string partNumber)
        {
            string setupWO = "";
            foreach (DataGridViewRow row in dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == partNumber
                && row.Cells["EquipNo"].Value.ToString() == "")
                {
                    setupWO = row.Cells["EQWorkorder"].Value.ToString();
                    break;
                }
            }
            return setupWO;
        }
        private bool CheckSNExistInGrid(string serialNumber)
        {
            bool isExist = false;
            foreach (DataGridViewRow row in this.dgvSN.Rows)
            {
                if (row.Cells[1].Value.ToString() == serialNumber)
                {
                    isExist = true;
                    break;
                }
            }
            return isExist;
        }
        private bool CheckEquipmentIsExist(string partNumber, string equipmentNo)
        {
            foreach (DataGridViewRow row in this.dgvEquipment.Rows)
            {
                if (row.Cells["eqPartNumber"].Value != null && row.Cells["eqPartNumber"].Value.ToString() == partNumber
                    && row.Cells["EquipNo"].Value.ToString() != "")
                {
                    if (row.Cells["EquipNo"].Value.ToString() == equipmentNo)
                    {
                        errorHandler(3, message("Equipment already setup"), "");
                        return false;
                    }
                    else
                    {
                        DialogResult dr = MessageBox.Show(message("Do you want to replace equipment"), message("Information"), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
                        if (dr == DialogResult.OK)
                        {
                            //Remove previous equipment and equUpdateEquipmentData 
                            string removeEquipNo = row.Cells["EquipNo"].Value.ToString();
                            string equipWO = row.Cells["EQWorkorder"].Value.ToString();
                            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
                            int errorCode = eqManager.UpdateEquipmentData(equipWO, iProcessLayer, equipmentNo, 1);
                            //rige equipment 
                            int errorCodeSetup = eqManager.UpdateEquipmentData(equipWO, iProcessLayer, equipmentNo, 0);
                            if (errorCodeSetup == 0)//1301 Equipment is already set up
                            {
                                EquipmentEntityExt entityExt = eqManager.GetSetupEquipmentData(equipmentNo);
                                if (entityExt != null)
                                {
                                    entityExt.PART_NUMBER = partNumber;
                                    entityExt.EQ_WORKORDER = equipWO;
                                    SetEquipmentGridData(entityExt);
                                    SetTipMessage(MessageType.OK, message("Process equipment number") + equipmentNo + message("SUCCESS"));
                                }
                            }
                        }
                        return false;
                    }
                }
            }
            return true;
        }

        private bool CheckEquipmentValid(string[] values)
        {
            bool isValid = false;
            if (values == null || values.Length == 0)
                return false;
            if (values[0] != "0")
                return false;
            foreach (DataGridViewRow item in this.dgvEquipment.Rows)
            {
                if (item.Cells["eqPartNumber"].Value.ToString() == values[2])// "EQUIPMENT_STATE", "ERROR_CODE", "PART_NUMBER"
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }

        #endregion

        #region Document
        static string cachePN = "";
        private void InitDocumentGrid()
        {
            if (config.FilterByFileName == "disable") //by station
            {
                if (gridDocument.Rows.Count <= 0)
                {
                    GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                    List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                    if (listDoc != null && listDoc.Count > 0)
                    {
                        foreach (DocumentEntity item in listDoc)
                        {
                            gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                        }
                    }
                }
            }
            else //by station & filename(partno)
            {
                if (this.txbCDAPartNumber.Text == "" || cachePN == this.txbCDAPartNumber.Text)
                    return;
                cachePN = this.txbCDAPartNumber.Text;
                gridDocument.Rows.Clear();
                this.Invoke(new MethodInvoker(delegate
                {
                    webBrowser1.Navigate("about:blank");
                }));
                GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
                List<DocumentEntity> listDoc = getDocument.GetDocumentDataByStation();
                if (listDoc != null && listDoc.Count > 0)
                {
                    foreach (DocumentEntity item in listDoc)
                    {
                        string filename = item.MDA_FILE_NAME;
                        Match name = Regex.Match(filename, config.FileNamePattern);
                        if (name.Success)
                        {
                            if (name.Groups.Count > 1)
                            {
                                string partno = name.Groups[1].ToString();
                                if (partno == this.txbCDAPartNumber.Text)
                                {
                                    gridDocument.Rows.Add(new object[2] { item.MDA_DOCUMENT_ID, item.MDA_FILE_NAME });
                                }
                            }
                        }
                    }
                }
            }
        }

        private void GetDocumentCollections()
        {
            GetDocumentData getDocument = new GetDocumentData(sessionContext, initModel, this);
            //get advice id
            Advice[] adviceArray = getDocument.GetAdviceByStationAndPN(this.txbCDAPartNumber.Text);
            if (adviceArray != null && adviceArray.Length > 0)
            {
                int iAdviceID = adviceArray[0].id;
                List<DocumentEntity> list = getDocument.GetDocumentDataByAdvice(iAdviceID);
                if (list != null && list.Count > 0)
                {
                    foreach (var item in list)
                    {
                        string docID = item.MDA_DOCUMENT_ID;
                        string fileName = item.MDA_FILE_NAME;
                        SetDocumentControl(docID, fileName);
                        break;
                    }
                }
            }
        }

        private void SetDocumentControl(string docID, string fileName)
        {
            GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
            byte[] content = documentHandler.GetDocumnetContentByID(Convert.ToInt64(docID));
            if (content != null)
            {
                string path = config.MDAPath;
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                string filePath = path + @"/" + fileName;
                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                Encoding.GetEncoding("gb2312");
                fs.Write(content, 0, content.Length);
                fs.Flush();
                fs.Close();
            }
        }

        private void SetDocumentControlForDoc(long documentID, string fileName)
        {
            string path = config.MDAPath;
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            string filePath = path + @"/" + fileName;
            if (!File.Exists(filePath))
            {
                GetDocumentData documentHandler = new GetDocumentData(sessionContext, initModel, this);
                byte[] content = documentHandler.GetDocumnetContentByID(documentID);
                if (content != null)
                {
                    FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                    fs.Write(content, 0, content.Length);
                    fs.Flush();
                    fs.Close();
                }
            }
            this.webBrowser1.Navigate(filePath);
        }
        #endregion

        #endregion

        #region Network status
        private string strNetMsg = "Network Connected";
        private void picNet_MouseHover(object sender, EventArgs e)
        {
            this.toolTip1.Show(message(strNetMsg), this.picNet);
        }

        private void AvailabilityChanged(object sender, NetworkAvailabilityEventArgs e)
        {
            this.Invoke(new MethodInvoker(delegate
            {
                if (e.IsAvailable)
                {
                    this.picNet.Image = ICTClient.Properties.Resources.NetWorkConnectedGreen24x24;
                    this.toolTip1.Show(message("Network Connected"), this.picNet);
                    strNetMsg = message("Network Connected");
                    LogHelper.Info(strNetMsg);
                    SetTopWindowMessage(strNetMsg, "");
                }
                else
                {
                    this.picNet.Image = ICTClient.Properties.Resources.NetWorkDisconnectedRed24x24;
                    this.toolTip1.Show(message("Network Disconnected"), this.picNet);
                    strNetMsg = message("Network Disconnected");
                    LogHelper.Error(strNetMsg);
                    SetTopWindowMessage("Network Status", strNetMsg);
                }

            }));

        }
        #endregion

        #region CheckList
        private void btnAddTask_Click(object sender, EventArgs e)
        {
            int iHour = DateTime.Now.Hour;
            if (8 <= iHour && iHour <= 18)
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
            }
            else
            {
                gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
            }
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clResult1"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clSeq"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clDate"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clShift"].ReadOnly = true;
            gridCheckList.Rows[this.gridCheckList.Rows.Count - 1].Cells["clStatus"].ReadOnly = true;
            gridCheckList.ClearSelection();
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            //if (!VerifyCheckList())
            //{
            //    errorHandler(2, message("checklist_first"), "");
            //    return;
            //}
            CheckListsCreate();
            #region
            //if (gridCheckList.Rows.Count > 0)
            //{
            //    string targetFileName = "";
            //    string shortFileName = config.StationNumber + "_" + this.gridCheckList.Rows[0].Cells["clShift"].Value.ToString() + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            //    bool isOK = CreateTemplate(shortFileName, ref targetFileName);
            //    if (isOK)
            //    {
            //        Excel.Application xlsApp = null;
            //        Excel._Workbook xlsBook = null;
            //        Excel._Worksheet xlsSheet = null;
            //        try
            //        {
            //            GC.Collect();
            //            xlsApp = new Excel.Application();
            //            xlsApp.DisplayAlerts = false;
            //            xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
            //            xlsBook = xlsApp.ActiveWorkbook;
            //            xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;

            //            int iBeginIndex = 7;
            //            Excel.Range range = null;
            //            foreach (DataGridViewRow row in gridCheckList.Rows)
            //            {
            //                range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
            //                range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
            //                string strSeq = row.Cells["clSeq"].Value.ToString();
            //                string strItemName = row.Cells["clItemName"].Value.ToString();
            //                string strItemPoint = row.Cells["clPoint"].Value.ToString();
            //                string strItemStandard = row.Cells["clStandard"].Value.ToString();
            //                string strItemMethod = row.Cells["clMethod"].Value.ToString();
            //                string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
            //                string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            //                string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
            //                string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
            //                string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
            //                string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
            //                string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
            //                string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();
            //                xlsSheet.Cells[iBeginIndex, 1] = strSeq;
            //                xlsSheet.Cells[iBeginIndex, 2] = strItemName;
            //                xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
            //                xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
            //                xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
            //                xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
            //                xlsSheet.Cells[iBeginIndex, 7] = strCheckDate;
            //                xlsSheet.Cells[iBeginIndex, 8] = strException;
            //                xlsSheet.Cells[iBeginIndex, 9] = strHappendTime;
            //                xlsSheet.Cells[iBeginIndex, 10] = strProcessContent;
            //                xlsSheet.Cells[iBeginIndex, 11] = strProcessPersion;
            //                xlsSheet.Cells[iBeginIndex, 12] = strOperator;
            //                xlsSheet.Cells[iBeginIndex, 13] = strLeader;
            //                iBeginIndex++;
            //            }
            //            xlsBook.Save();
            //            errorHandler(0, "Save Production Check List success.(" + targetFileName + ")", "");
            //        }
            //        catch (Exception ex)
            //        {
            //            LogHelper.Error(ex);
            //        }
            //        finally
            //        {
            //            xlsBook.Close(false, Type.Missing, Type.Missing);
            //            xlsApp.Quit();
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
            //            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

            //            xlsSheet = null;
            //            xlsBook = null;
            //            xlsApp = null;

            //            GC.Collect();
            //            GC.WaitForPendingFinalizers();
            //        }
            //    }
            //}
            #endregion
        }
        #region add by qy
        private void CheckListsCreate()
        {
            if (gridCheckList.Rows.Count > 0)
            {
                string targetFileName = "";
                string shortFileName = config.StationNumber + "_VI_" + DateTime.Now.ToString("yyyyMM");
                bool isOK = CreateTemplate(shortFileName, ref targetFileName);
                if (isOK)
                {
                    Excel.Application xlsApp = null;
                    Excel._Workbook xlsBook = null;
                    Excel._Worksheet xlsSheet = null;
                    try
                    {
                        GC.Collect();
                        xlsApp = new Excel.Application();
                        xlsApp.DisplayAlerts = false;
                        xlsApp.Workbooks.Open(targetFileName, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        xlsBook = xlsApp.ActiveWorkbook;
                        xlsSheet = (Excel._Worksheet)xlsBook.ActiveSheet;
                        int count = xlsSheet.UsedRange.Cells.Rows.Count;

                        int iBeginIndex = count;
                        Excel.Range range = null;
                        foreach (DataGridViewRow row in gridCheckList.Rows)
                        {
                            range = (Excel.Range)xlsSheet.Rows[iBeginIndex, Missing.Value];
                            range.Rows.Insert(Excel.XlDirection.xlDown, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove);
                            string strSeq = row.Cells["clSeq"].Value.ToString();
                            string strItemName = row.Cells["clItemName"].Value.ToString();
                            string strItemPoint = row.Cells["clPoint"].Value.ToString();
                            string strItemStandard = row.Cells["clStandard"].Value.ToString();
                            string strItemMethod = row.Cells["clMethod"].Value.ToString();
                            string strItemResult = GetCheckItemResult(row.Cells["clResult1"].Value.ToString(), row.Cells["clResult2"].Value.ToString());
                            string strCheckDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                            string strShift = row.Cells["clShift"].Value.ToString();
                            string strException = row.Cells["clException"].Value == null ? "" : row.Cells["clException"].Value.ToString();
                            string strHappendTime = row.Cells["clChangeDate"].Value == null ? "" : row.Cells["clChangeDate"].Value.ToString();
                            string strProcessContent = row.Cells["clContent"].Value == null ? "" : row.Cells["clContent"].Value.ToString();
                            string strProcessPersion = row.Cells["clPersion"].Value == null ? "" : row.Cells["clPersion"].Value.ToString();
                            string strOperator = row.Cells["clOperator"].Value == null ? "" : row.Cells["clOperator"].Value.ToString();
                            string strLeader = row.Cells["clLeader"].Value == null ? "" : row.Cells["clLeader"].Value.ToString();

                            xlsSheet.Cells[iBeginIndex, 1] = iBeginIndex - 7;
                            xlsSheet.Cells[iBeginIndex, 2] = strItemName;
                            xlsSheet.Cells[iBeginIndex, 3] = strItemPoint;
                            xlsSheet.Cells[iBeginIndex, 4] = strItemStandard;
                            xlsSheet.Cells[iBeginIndex, 5] = strItemMethod;
                            xlsSheet.Cells[iBeginIndex, 6] = strItemResult;
                            xlsSheet.Cells[iBeginIndex, 7] = strShift;
                            xlsSheet.Cells[iBeginIndex, 8] = strCheckDate;
                            xlsSheet.Cells[iBeginIndex, 9] = strException;
                            xlsSheet.Cells[iBeginIndex, 10] = strHappendTime;
                            xlsSheet.Cells[iBeginIndex, 11] = strProcessContent;
                            xlsSheet.Cells[iBeginIndex, 12] = strProcessPersion;
                            xlsSheet.Cells[iBeginIndex, 13] = strOperator;
                            xlsSheet.Cells[iBeginIndex, 14] = strLeader;

                            iBeginIndex++;
                        }
                        xlsBook.Save();
                        SetAlarmStatusText("");
                        errorHandler(0, message("Save Production Check List success") + targetFileName + ")", "");
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex);
                    }
                    finally
                    {
                        xlsBook.Close(false, Type.Missing, Type.Missing);
                        xlsApp.Quit();
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsApp);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsBook);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(xlsSheet);

                        xlsSheet = null;
                        xlsBook = null;
                        xlsApp = null;

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
            }
        }

        private bool CheckCheckList()
        {
            bool result = true;
            foreach (DataGridViewRow row in gridCheckList.Rows)
            {
                string status = row.Cells["clStatus"].Value.ToString();
                if (status != "OK")
                {
                    result = false;
                    errorHandler(2, message("Production checklist is required"), "");
                    break;
                }
            }
            return result;
        }

        #endregion
        private string GetCheckItemResult(string result1, string result2)
        {
            if (string.IsNullOrEmpty(result1))
                return result2;
            if (string.IsNullOrEmpty(result2))
                return result1;
            else
                return "NA";
        }

        private void InitTaskData()
        {
            try
            {
                Dictionary<string, List<CheckListItemEntity>> dicTask = new Dictionary<string, List<CheckListItemEntity>>();
                XDocument xdc = XDocument.Load("TaskFile.xml");
                var stationNodes = from item in xdc.Descendants("StationNumber")
                                   where item.Attribute("value").Value == config.StationNumber
                                   select item;
                XElement stationNode = stationNodes.FirstOrDefault();
                var tasks = from item in stationNode.Descendants("shift")
                            select item;
                foreach (XElement node in tasks.ToList())
                {
                    string shiftValue = GetNoteAttributeValues(node, "value");
                    List<CheckListItemEntity> itemList = new List<CheckListItemEntity>();
                    var items = from item in node.Descendants("Item")
                                select item;
                    foreach (XElement subItem in items.ToList())
                    {
                        CheckListItemEntity entity = new CheckListItemEntity();
                        entity.ItemName = GetNoteAttributeValues(subItem, "name");
                        entity.ItemPoint = GetNoteAttributeValues(subItem, "point");
                        entity.ItemStandard = GetNoteAttributeValues(subItem, "standard");
                        entity.ItemMethod = GetNoteAttributeValues(subItem, "method");
                        entity.ItemInputType = GetNoteAttributeValues(subItem, "inputType");
                        itemList.Add(entity);
                    }
                    if (!dicTask.ContainsKey(shiftValue))
                    {
                        dicTask[shiftValue] = itemList;
                    }
                }
                //init check list grid
                string strInputValue = GetNoteDescendantsValues(stationNode, "DataInputType");
                string[] strInputValues = strInputValue.Split(new char[] { ',' });
                DataTable dtInput = new DataTable();
                dtInput.Columns.Add("name");
                dtInput.Columns.Add("value");
                DataRow rowEmpty = dtInput.NewRow();
                rowEmpty["name"] = "";
                rowEmpty["value"] = "";
                dtInput.Rows.Add(rowEmpty);
                foreach (var strValues in strInputValues)
                {
                    DataRow row = dtInput.NewRow();
                    row["name"] = strValues;
                    row["value"] = strValues;
                    dtInput.Rows.Add(row);
                }
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DataSource = dtInput;
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).DisplayMember = "Name";
                ((DataGridViewComboBoxColumn)this.gridCheckList.Columns["clResult2"]).ValueMember = "Value";

                int iHour = DateTime.Now.Hour;
                int seq = 1;
                if (8 <= iHour && iHour <= 18)
                {
                    if (dicTask.ContainsKey("白班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["白班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "白班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
                else
                {
                    if (dicTask.ContainsKey("晚班"))
                    {
                        List<CheckListItemEntity> itemList = dicTask["晚班"];
                        if (itemList != null && itemList.Count > 0)
                        {
                            foreach (var item in itemList)
                            {
                                object[] objValues = new object[11] { seq, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", item.ItemName, item.ItemPoint, item.ItemStandard, item.ItemMethod, "", "", "", item.ItemInputType };
                                this.gridCheckList.Rows.Add(objValues);
                                seq++;
                            }
                            SetCheckListInputStatus();
                            this.gridCheckList.ClearSelection();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }

        private void SetCheckListInputStatus()
        {
            foreach (DataGridViewRow row in this.gridCheckList.Rows)
            {
                if (row.Cells["clInputType"].Value.ToString() == "1")
                {
                    row.Cells["clResult1"].ReadOnly = true;
                }
                else if (row.Cells["clInputType"].Value.ToString() == "2")
                {
                    row.Cells["clResult2"].ReadOnly = true;
                }
            }
        }
        #region Grid ComboBox
        int iRowIndex = -1;
        private void gridCheckList_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBox_SelectedIndexChanged);
            }
        }

        public void ComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(combox_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "OK";
                    }
                    else
                    {
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                        this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Style.BackColor = Color.White;
                    this.gridCheckList.Rows[iRowIndex].Cells["clStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public void combox_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBox_SelectedIndexChanged);
        }
        #endregion

        int iIndexCheckList = -1;
        private void gridCheckList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.gridCheckList.Rows.Count == 0)
                    return;
                this.gridCheckList.ContextMenuStrip = contextMenuStrip2;
                iIndexCheckList = ((DataGridView)sender).CurrentRow.Index;
            }
        }

        private void checkListAdd_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                int iHour = DateTime.Now.Hour;
                if (8 <= iHour && iHour <= 18)
                {
                    this.gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "白班", "", "", "", "", "", "", "", "" });
                }
                else
                {
                    this.gridCheckList.Rows.Add(new object[] { this.gridCheckList.Rows.Count + 1, DateTime.Now.ToString("yyyy/MM/dd"), "晚班", "", "", "", "", "", "", "", "" });
                }
                this.gridCheckList.ClearSelection();
            }
        }

        private void checkListDelete_Click(object sender, EventArgs e)
        {
            if (iIndexCheckList > -1)
            {
                this.gridCheckList.Rows.RemoveAt(iIndexCheckList);
                int seq = 1;
                foreach (DataGridViewRow row in this.gridCheckList.Rows)
                {
                    row.Cells["clSeq"].Value = seq;
                    seq++;
                }
                this.gridCheckList.ClearSelection();
            }
        }

        private bool CreateTemplate(string strFileName, ref string targetFileName)
        {
            bool bFlag = true;
            targetFileName = "";
            string filePath = Assembly.GetExecutingAssembly().GetModules()[0].FullyQualifiedName;
            string _appDir = Path.GetDirectoryName(filePath);
            string strExportPath = _appDir + @"\CheckListFiles\";
            //临时文件目录
            if (Directory.Exists(strExportPath) == false)
            {
                Directory.CreateDirectory(strExportPath);
            }
            string strSourceFileName = strExportPath + @"CheckListTemplate.xls";
            string strTargetFileName = config.CheckListFolder + strFileName + ".xls";
            targetFileName = strTargetFileName;

            if (System.IO.File.Exists(strSourceFileName))
            {
                try
                {
                    if (System.IO.File.Exists(strTargetFileName))
                        return true;

                    System.IO.File.Copy(strSourceFileName, strTargetFileName, true);
                    //去掉文件Readonly,避免不可写
                    FileInfo file = new FileInfo(strTargetFileName);
                    if ((file.Attributes & FileAttributes.ReadOnly) > 0)
                    {
                        file.Attributes ^= FileAttributes.ReadOnly;
                    }
                }
                catch (Exception ex)
                {
                    bFlag = false;
                    LogHelper.Error(ex);
                    throw ex;
                }
            }
            else
            {
                bFlag = false;
            }

            return bFlag;
        }

        private string GetNoteAttributeValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Attribute(attributename).Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        private string GetNoteDescendantsValues(XElement node, string attributename)
        {
            string strValue = "";
            try
            {
                strValue = node.Descendants(attributename).First().Value;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
            return strValue;
        }

        #endregion

        #region Listen files add by qy

        Dictionary<string, string> dicFailureMap = new Dictionary<string, string>();
        Dictionary<int, string> dicMap = new Dictionary<int, string>();
        private void InitFailureMapTable()
        {
            GetFailureData getfailuredata = new GetFailureData(sessionContext, initModel, this);
            string[] snStateResultValues1 = { };
            snStateResultValues1 = getfailuredata.MdataGetFailureDataforStation();
            if (snStateResultValues1.Length > 0)
            {
                for (int i = 0; i < snStateResultValues1.Length / 2; i++)
                {
                    if (!dicFailureMap.ContainsKey(snStateResultValues1[2 * i + 1]))
                    {
                        dicFailureMap[snStateResultValues1[2 * i + 1].ToUpper()] = snStateResultValues1[2 * i + 0];
                    }
                }
            }

            string[] mapList = File.ReadAllLines("FailureMap.txt", Encoding.Default);
            foreach (var line in mapList)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    string[] strs = line.Split(new char[] { ';' });
                    if (!dicMap.ContainsKey(Convert.ToInt32(strs[0])))
                    {
                        dicMap[Convert.ToInt32(strs[0])] = strs[1];
                    }
                }
            }

            //string[] LineList = File.ReadAllLines("FailureMap.txt", Encoding.Default);
            //foreach (var line in LineList)
            //{
            //    if (!string.IsNullOrEmpty(line))
            //    {
            //        string[] strs = line.Split(new char[] { ';' });
            //        if (!dicFailureMap.ContainsKey(strs[1]))
            //        {
            //            dicFailureMap[strs[1]] = strs[0];
            //        }
            //    }
            //}
        }

        private int ConvertResultToInt(string snstate)
        {
            int iValue = 0;
            switch (snstate)
            {
                case "PASS":
                    iValue = 0;
                    break;
                case "FAIL":
                    iValue = 1;
                    break;
                case "P":
                    iValue = 0;
                    break;
                case "F":
                    iValue = 1;
                    break;
                default:
                    iValue = 1;
                    break;
            }
            return iValue;
        }

        private string TransValue(string svalue)
        {
            int KValue = Convert.ToInt32(config.kValue);
            int MEGValue = Convert.ToInt32(config.MEGValue);
            double MValue = Convert.ToDouble(config.MValue);
            double UValue = Convert.ToDouble(config.UValue);
            double NValue = Convert.ToDouble(config.NValue);
            double PValue = Convert.ToDouble(config.PValue);
            if (svalue.Contains("K"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 1)) * KValue;
            }
            else if (svalue.Contains("MEG"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 3)) * MEGValue;
            }
            else if (svalue.Contains("M"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 1)) * MValue;
            }
            else if (svalue.Contains("U"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 1)) * UValue;
            }
            else if (svalue.Contains("N"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 1)) * NValue;
            }
            else if (svalue.Contains("P"))
            {
                return "" + Convert.ToDouble(svalue.Substring(0, svalue.Length - 1)) * PValue;
            }
            else
                return svalue;
        }

        private static string looperName(string name, List<string> namelist)
        {
            int count = 1;

            string newname = name;
            while (namelist.Contains(newname))
            {
                newname = string.Format("{0}({1})", name, count++);
                count++;
            }
            return newname;
        }

        private List<string> AddMeasureInfoToList(string[] LineList)    //trUploadResultDataAndRecipe--resultUploadValues
        {
            List<string> listItem = new List<string>(); //储存upload数据

            for (int i = 6; i < LineList.Length; i++)
            {
                string linecontent = LineList[i].Trim();
                if (linecontent.Contains("//") || linecontent == "" || linecontent.Contains("BEGIN") || linecontent.Contains("END") || linecontent.Contains("DISCH"))//去除无效的信息
                {
                    continue;
                }
                Match match = Regex.Match(linecontent, config.LineRegular);
                if (match.Success)
                {
                    GroupCollection linegroups = match.Groups;

                    string lowerLimit = linegroups[Convert.ToInt32(config.LowerLimit)].ToString().Trim();
                    bool b_parse = false;
                    double d_lLimit = 0;
                    if (lowerLimit != "")
                    {
                        lowerLimit = TransValue(lowerLimit);
                        b_parse = Double.TryParse(lowerLimit, out d_lLimit);
                        if (b_parse)
                        {
                            lowerLimit = (decimal)d_lLimit + "";
                        }
                    }
                    listItem.Add("0");//ErrorCode
                    listItem.Add(lowerLimit);//LowerLimit

                    string measureFailCode = linegroups[Convert.ToInt32(config.MeasureFailCode)].ToString().Trim();
                    listItem.Add(measureFailCode);//measureFailCode

                    //检查是否有重复的measurename,有就在后面加xxxx(1),xxxx(2)
                    string[] msnameRegular = config.MeasureName.Split('|');
                    string name = linegroups[Convert.ToInt32(msnameRegular[0])].ToString().Trim() + linegroups[Convert.ToInt32(msnameRegular[1])].ToString().Trim();
                    string measurename = looperName(name, listItem);
                    if (measurename.Length > 78)
                    {
                        measurename = measurename.Substring(0, 78);
                    }
                    listItem.Add(measurename); //MeasureName

                    //拆分Msr_V
                    Match match1 = Regex.Match(linegroups[Convert.ToInt32(config.MeasureValue)].ToString().Trim(), config.MNRegular);
                    if (match1.Success)
                    {
                        GroupCollection MNgroups = match1.Groups;
                        string measurevalue = MNgroups[1].ToString().Trim();
                        if (measurevalue != "")
                        {
                            measurevalue = TransValue(measurevalue);
                            b_parse = Double.TryParse(measurevalue, out d_lLimit);
                            if (b_parse)
                            {
                                measurevalue = (decimal)d_lLimit + "";
                            }
                        }
                        listItem.Add(measurevalue);//measurevalue

                        string unit = MNgroups[2].ToString().Trim();
                        listItem.Add(unit);//unit
                    }

                    listItem.Add("");//nominal
                    listItem.Add("");//新增 Remark
                    listItem.Add("0");//Tolerance

                    string maxlimit = linegroups[Convert.ToInt32(config.UpperLimit)].ToString().Trim();
                    if (maxlimit != "")
                    {
                        maxlimit = TransValue(maxlimit);
                        b_parse = Double.TryParse(maxlimit, out d_lLimit);
                        if (b_parse)
                        {
                            maxlimit = (decimal)d_lLimit + "";
                        }
                    }
                    listItem.Add(maxlimit);//UpperLimit
                }
            }
            return listItem;
        }

        private Dictionary<string, List<string>> AddIntoMeasureValues(string[] LineList) //trUploadFailureAndResultData
        {
            string serialnumberpos = "";   //每次读取新文档清空serialnumberpos
            Dictionary<string, List<string>> dicmeasure = new Dictionary<string, List<string>>();

            for (int i = 6; i < LineList.Length; i++)
            {
                List<string> listItem = new List<string>(); //储存upload数据
                string linecontent = LineList[i].Trim();
                if (linecontent.Contains("//") || linecontent == "" || linecontent.Contains("BEGIN") || linecontent.Contains("END") || linecontent.Contains("DISCH"))//去除无效的信息
                {
                    continue;
                }
                Match match = Regex.Match(linecontent, config.LineRegular);
                if (match.Success)
                {
                    GroupCollection linegroups = match.Groups;

                    bool b_parse = false;
                    double d_measurev = 0;
                    listItem.Add("0");//ErrorCode

                    string measureFailCode = linegroups[Convert.ToInt32(config.MeasureFailCode)].ToString().Trim();
                    listItem.Add(measureFailCode);//measureFailCode

                    //检查是否有重复的measurename,有就在后面加xxxx(1),xxxx(2)
                    string[] msnameRegular = config.MeasureName.Split('|');
                    string name = linegroups[Convert.ToInt32(msnameRegular[0])].ToString().Trim() + linegroups[Convert.ToInt32(msnameRegular[1])].ToString().Trim();
                    string measurename = looperName(name, listItem);
                    if (measurename.Length > 78)
                    {
                        measurename = measurename.Substring(0, 78);
                    }
                    listItem.Add(measurename); //MeasureName

                    //serialnumberpos
                    string[] pos = linegroups[Convert.ToInt32(msnameRegular[1])].ToString().Trim().Split('#');
                    serialnumberpos = pos[1];

                    //拆分Msr_V
                    Match match1 = Regex.Match(linegroups[Convert.ToInt32(config.MeasureValue)].ToString().Trim(), config.MNRegular);
                    if (match1.Success)
                    {
                        GroupCollection MNgroups = match1.Groups;
                        string measurevalue = MNgroups[1].ToString().Trim();
                        if (measurevalue != "")
                        {
                            measurevalue = TransValue(measurevalue);
                            b_parse = Double.TryParse(measurevalue, out d_measurev);
                            if (b_parse)
                            {
                                measurevalue = (decimal)d_measurev + "";
                            }
                        }
                        listItem.Add(measurevalue);//measurevalue
                    }

                    if (!dicmeasure.ContainsKey(serialnumberpos))
                    {
                        dicmeasure.Add(serialnumberpos, listItem);
                    }
                    else
                    {
                        dicmeasure[serialnumberpos].AddRange(listItem);
                    }
                }
            }
            return dicmeasure;
        }

        private string FailureCodeMap(int code)
        {
            string failuremap = "";
            foreach (var key in dicMap.Keys)
            {
                if (Convert.ToInt32(key.ToString()) == code)
                {
                    failuremap = dicMap[code].ToString().ToUpper();
                }
            }
            if (failuremap == "")
            {
                failuremap = dicMap[100].ToString().ToUpper();
            }
            return failuremap;
        }

        private void MoveFileToOKFolder(string filepath)
        {
            string OkFolder = config.LogTransOK;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to ok folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(OkFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }

                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to ok folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = config.LogTransOK + strDirName + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(OkFolder)) Directory.CreateDirectory(OkFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, "Move file:" + filepath + " to OK folder success.", "");
            }
            catch (Exception e)
            {
                errorHandler(2, "move file error " + e.Message, "");
            }
        }

        private void MoveFileToErrorFolder(string filepath, string errorMsg)
        {
            string errorFolder = config.LogTransError;
            string strDir = Path.GetDirectoryName(filepath) + @"\";
            string strDirCopy = Path.GetDirectoryName(filepath);
            string strDestDir = "";
            try
            {
                if (strDir == config.LogFileFolder)//move file to error folder
                {
                    FileInfo fInfo = new FileInfo(@"" + filepath);
                    string fileNameOnly = Path.GetFileNameWithoutExtension(filepath);
                    string extension = Path.GetExtension(filepath);
                    string newFullPath = null;
                    if (config.ChangeFileName.ToUpper() == "ENABLE")
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + extension);
                    }
                    else
                    {
                        newFullPath = Path.Combine(errorFolder, fileNameOnly + extension);
                    }
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (File.Exists(newFullPath))
                    {
                        File.Delete(newFullPath);
                    }
                    fInfo.MoveTo(@"" + newFullPath);
                }
                else//move Directory to error folder
                {
                    string strDirName = strDirCopy.Substring(strDirCopy.LastIndexOf(@"\") + 1);
                    strDestDir = errorFolder + strDirName + "_" + errorMsg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
                    if (!Directory.Exists(errorFolder)) Directory.CreateDirectory(errorFolder);
                    if (Directory.Exists(strDestDir))
                    {
                        Directory.Delete(strDestDir, true);
                    }
                    Directory.Move(strDir, strDestDir);
                }
                errorHandler(1, "Move file:" + filepath + " to error folder success.", "");
            }
            catch (Exception e)
            {
                errorHandler(2, "move file error " + e.Message, "");
            }
        }
        #endregion

        #region Tray
        private void GetSerialNumberByTray()
        {
            GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
            string value = "";
            string[] attributeResultValues = getattribute.GetAttributeValueForEquipment("TraySeq", TrayNumber);
            if (attributeResultValues.Length > 0)
            {
                value = TrayNumber + "_" + attributeResultValues[1];

                string[] objectResultValues = getattribute.GetObjectValeForSN("TrayNo", value);
                if (objectResultValues.Length > 0)
                {
                    this.Invoke(new MethodInvoker(delegate
                    {
                        dgvSN.Rows.Clear();
                        dgvSNInfo.Rows.Clear();
                        dgvCompName.Rows.Clear();
                        dgvDefect.Rows.Clear();
                    }));
                    foreach (var item in objectResultValues)
                    {
                        this.Invoke(new MethodInvoker(delegate
                        {
                            LoadVIC(item);
                        }));
                    }

                }
            }
            else
            {
                this.errorHandler(2, message("Tray and SerialNumber has been released"), message("Tray and SerialNumber has been released"));
                return;
            }
        }

        private bool CheckTrayExistInItac(string equipment)
        {
            bool isValid = true;
            EquipmentManager eqManager = new EquipmentManager(sessionContext, initModel, this);
            string[] values = eqManager.GetEquipmentDetailData(equipment);
            if (values == null || values.Length == 0)
            {
                isValid = false;
            }

            return isValid;
        }
        #endregion

        #region inspection

        private bool LoadVICSingle(string serialnumber)
        {
            //gate keeper,check serial state
            CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
            snentitys = checkHandler.GetSerialNumberData(serialnumber, iProcessLayer);

            if (snentitys.Count > 0)
            {
                //dgvSN.Rows.Clear();
                //LoadFailureType();
                for (int i = 0; i < snentitys.Count; i++)
                {
                    if (snentitys[i].serialNumber == serialnumber) //add by qy
                    {
                        string state = snentitys[i].serialNumberState;
                        if (state == (int)EnumCommon.EnumResultState.SCRAP + "")
                        {
                            state = EnumCommon.EnumResultState.SCRAP.ToString();
                        }
                        else if (state == (int)EnumCommon.EnumResultState.PASS + "")
                        {
                            state = EnumCommon.EnumResultState.PASS.ToString();
                        }
                        else if (state == (int)EnumCommon.EnumResultState.FAIL + "")
                        {
                            state = EnumCommon.EnumResultState.FAIL.ToString();
                        }
                        if (!CheckSNExistInGrid(snentitys[i].serialNumber))
                        {
                            dgvSN.Rows.Add(new string[3] { snentitys[i].serialNumberPos, snentitys[i].serialNumber, state });
                        }
                        if (state == EnumCommon.EnumResultState.SCRAP.ToString())
                        {
                            dgvSN.Rows[i].DefaultCellStyle.BackColor = Color.Gray;
                        }
                    }
                }
                //dgvSN_CellClick(null, new DataGridViewCellEventArgs(0, 0));

            }
            else
            {
                errorHandler(2, message("Check Serial Number State Error."), "");
                return false;
            }
            return true;
        }
        List<SerialNumberStateEntity> snentitys = null;
        private bool LoadVIC(string serialnumber)
        {
            //gate keeper,check serial state
            CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
            snentitys = checkHandler.GetSerialNumberData(serialnumber, iProcessLayer);

            if (snentitys.Count > 0)
            {
                dgvSN.Rows.Clear();
                //LoadFailureType();
                for (int i = 0; i < snentitys.Count; i++)
                {
                    //if (snentitys[i].serialNumber == serialnumber) //add by qy
                    //{
                    string state = snentitys[i].serialNumberState;
                    if (state == (int)EnumCommon.EnumResultState.SCRAP + "")
                    {
                        state = EnumCommon.EnumResultState.SCRAP.ToString();
                    }
                    else if (state == (int)EnumCommon.EnumResultState.PASS + "")
                    {
                        state = EnumCommon.EnumResultState.PASS.ToString();
                    }
                    else if (state == (int)EnumCommon.EnumResultState.FAIL + "")
                    {
                        state = EnumCommon.EnumResultState.FAIL.ToString();
                    }
                    else if (state == (int)EnumCommon.EnumResultState.ERROR + "")
                    {
                        state = EnumCommon.EnumResultState.ERROR.ToString();
                    }
                    if (!CheckSNExistInGrid(snentitys[i].serialNumber))
                    {
                        dgvSN.Rows.Add(new string[3] { snentitys[i].serialNumberPos, snentitys[i].serialNumber, state });
                    }
                    if (state == EnumCommon.EnumResultState.ERROR.ToString())
                    {
                        dgvSN.Rows[i].DefaultCellStyle.BackColor = Color.Gray;
                    }
                    if (state == EnumCommon.EnumResultState.SCRAP.ToString())
                    {
                        dgvSN.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    }
                    //}
                }
                this.dgvSN.Sort((new RowComparer2(SortOrder.Ascending)));
                dgvSN.ClearSelection();
                //dgvSN_CellClick(null, new DataGridViewCellEventArgs(0, 0));

            }
            else
            {
                errorHandler(2, message("Check Serial Number State Error."), "");
                return false;
            }
            return true;
        }

        private class RowComparer2 : System.Collections.IComparer
        {
            private static int sortOrderModifier = 1;

            public RowComparer2(SortOrder sortOrder)
            {
                if (sortOrder == SortOrder.Descending)
                {
                    sortOrderModifier = -1;
                }
                else if (sortOrder == SortOrder.Ascending) { sortOrderModifier = 1; }
            }
            public int Compare(object x, object y)
            {
                DataGridViewRow DataGridViewRow1 = (DataGridViewRow)x;
                DataGridViewRow DataGridViewRow2 = (DataGridViewRow)y;
                // Try to sort based on the Scan time column.
                string str1 = DataGridViewRow1.Cells["Column4"].Value.ToString();
                string str2 = DataGridViewRow2.Cells["Column4"].Value.ToString();
                int dt1 = Convert.ToInt32(DataGridViewRow1.Cells["Column2"].Value.ToString());
                int dt2 = Convert.ToInt32(DataGridViewRow2.Cells["Column2"].Value.ToString());
                int CompareResult = str1.CompareTo(str2);
                if (CompareResult > 0) return -1 * sortOrderModifier;
                else if (CompareResult < 0) return 1 * sortOrderModifier;

                //int CompareResult = str1.CompareTo(str2);
                //return CompareResult * sortOrderModifier;

                int CompareResult1 = dt1.CompareTo(dt2);
                return CompareResult1 * sortOrderModifier;
            }

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (dgvSNInfo.Rows.Count > 0)
            {
                bool isCancel = true;
                int rowindex = dgvSNInfo.CurrentRow.Index;
                if (rowindex > -1)
                {
                    string sn = dgvSNInfo.Rows[rowindex].Cells[1].Value.ToString();
                    dgvSNInfo.Rows.RemoveAt(rowindex);
                    foreach (DataGridViewRow row in this.dgvSNInfo.Rows)
                    {
                        if (row.Cells[1].Value.ToString() == sn)
                        {
                            isCancel = false;
                        }
                    }
                    if (isCancel)
                    {
                        EditDataToTraySNGrid(sn, "PASS");
                    }

                    this.dgvSNInfo.ClearSelection();
                    this.btnCancel.Enabled = false;
                }
            }
            else
            {
                //把状态给改回来 改为Pass
            }
        }

        private string StringToHexString(string s, Encoding encode)
        {
            byte[] b = encode.GetBytes(s);//按照指定编码将string编程字节数组
            string result = string.Empty;
            for (int i = 0; i < b.Length; i++)//逐字节变为16进制字符，以%隔开
            {
                result += Convert.ToString(b[i], 16);
            }
            return result;
        }

        List<string> sns = null;
        private void btnUpload_Click(object sender, EventArgs e)
        {
            int serialState = 0;
            List<string> snInfoValues = new List<string>();
            sns = new List<string>();
            int scrapcount = 0;
            if (dgvSN.Rows.Count > 0)
            {
                GetSerialNumberInfo getSNHandler = new GetSerialNumberInfo(sessionContext, initModel, this);
                int count = dgvSN.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    string serialNumber = dgvSN.Rows[i].Cells[1].Value.ToString();
                    if (!sns.Contains(serialNumber))
                    {
                        CheckSerialNumberState checkHandler = new CheckSerialNumberState(sessionContext, initModel, this);
                        List<SerialNumberStateEntity> snentitys = checkHandler.GetSerialNumberData(serialNumber, iProcessLayer);

                        if (snentitys.Count > 0)
                        {
                            for (int K = 0; K < snentitys.Count; K++)
                            {
                                string state = snentitys[K].serialNumberState;
                                sns.Add(snentitys[K].serialNumberPos);
                                sns.Add(snentitys[K].serialNumber);
                                sns.Add(state);
                                if (state == "2")
                                    scrapcount++;
                            }
                        }

                        //SerialNumberData[] snArray = getSNHandler.GetSerialNumberBySNRef(serialNumber);
                        //foreach (var item in snArray)
                        //{
                        //    sns.Add(item.serialNumberPos);
                        //    sns.Add(item.serialNumber);
                        //    sns.Add("0");
                        //}
                    }
                    //sns.Add(dgvSN.Rows[i].Cells[0].Value.ToString());
                    //sns.Add(dgvSN.Rows[i].Cells[1].Value.ToString());
                    //string state = dgvSN.Rows[i].Cells[2].Value.ToString();
                    //if (state == EnumCommon.EnumResultState.SCRAP.ToString())
                    //    state = (int)EnumCommon.EnumResultState.SCRAP + "";
                    //else if (state == EnumCommon.EnumResultState.PASS.ToString())
                    //    state = (int)EnumCommon.EnumResultState.PASS + "";
                    //else if (state == EnumCommon.EnumResultState.FAIL.ToString())
                    //    state = (int)EnumCommon.EnumResultState.FAIL + "";
                    //sns.Add(state);
                }
            }
            else
            {
                MessageBox.Show(message("Pls input the [Data Input]"));
                return;
            }

            //upload
            if (dgvSNInfo.Rows.Count > 0)
            {
                serialState = 1;
                int count1 = dgvSNInfo.Rows.Count;
                for (int i = 0; i < count1; i++)
                {
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[0].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[1].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[2].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[3].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[4].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[5].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[6].Value.ToString());
                    snInfoValues.Add(dgvSNInfo.Rows[i].Cells[7].Value.ToString());
                }
            }

            bool b = UploadFailureTypeAndCompNew(serialState, sns, snInfoValues);
            if (b)
            {

                this.Invoke(new MethodInvoker(delegate
                {
                    LoadYield();
                }));
                SetTopWindowMessage(tempserialnumber, "");
                errorHandler(0, message("Upload the Comp.Name and Defect is success") + tempserialnumber, "");
                if (snInfoValues.Count > 0)
                {
                    if (config.DataOutputInterface != "" && config.DataOutputInterface != null)
                    {
                        if (SendSN(config.NG_CHANNEL_OPEN))
                        {
                            Thread.Sleep(Convert.ToInt32(config.CommandSleepTime));
                            SendSN(config.NG_CHANGE_CLOSE);
                        }
                    }

                    UpdateIPIStatus("1");//首件检查 失败
                    UpdateIPIStatusForProductionInspection("1", tempserialnumber);//生产检查失败
                    SetAlarmStatusText("失败");
                }
                else
                {
                    if (config.DataOutputInterface != "" && config.DataOutputInterface != null)
                    {
                        if (SendSN(config.OK_CHANNEL_Open))
                        {
                            Thread.Sleep(Convert.ToInt32(config.CommandSleepTime));
                            SendSN(config.OK_CHANNEL_CLOSE);
                        }
                    }

                    UpdateIPIStatus("0");//首件检查 成功
                    if (config.IsPrint == "Enable")
                    {
                        //CheckLocalFile();
                        //Pscount = Convert.ToInt32(this.lblPCBQty.Text.Split('/')[0]);

                        #region add by qy 20161223 检测同一个条码在一个通箱中只能计数1次
                        List<string> scannedSNs = new List<string>();
                        if (File.Exists(@"LocalRackInfo.txt"))
                        {
                            string[] Linelists = File.ReadAllLines(@"LocalRackInfo.txt", Encoding.Default);
                            foreach (string content in Linelists)
                            {
                                if (content != "")
                                {
                                    string[] cons = content.Split(',');
                                    scannedSNs.Add(cons[4]);
                                }
                            }
                        }
                        foreach (string itemsn in scannedSNs)
                        {
                            if (itemsn.Contains(sns[1]) || sns[1].Contains(itemsn))
                            {
                                SetAlarmStatusText("通过");
                                dgvSNInfo.Rows.Clear();
                                dgvSN.Rows.Clear();
                                dgvCompName.Rows.Clear();
                                dgvDefect.Rows.Clear();
                                this.txbCDADataInput.Text = "";
                                this.txbCDADataInput.Focus();
                                tempserialnumber = "";
                                return;
                            }
                        }
                        #endregion

                        Pscount++;

                        //append "Rack SN"
                        QueueEntity qe = new QueueEntity();
                        qe.rackSNInfoExt = sns;
                        SEQueue.Enqueue(qe);

                        //AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                        //for (int s = 0; s < sns.Count / 3; s++)
                        //{
                        //    string item = sns[s * 3 + 1];
                        //    int state = Convert.ToInt32(sns[s * 3 + 2]);
                        //    if (state != 2 && state != -1)
                        //        appendAttri.AppendAttributeForAll(0, item, "-1", "RACK_SN", this.lblTextRackNo.Text);
                        //    //else
                        //    //    scrapcount++;
                        //}
                        sn_att = sns[1];
                        int passQty = Convert.ToInt32(lblPass.Text);
                        int failQty = Convert.ToInt32(lblFail.Text);
                        int lastQty = workorderTotal - passQty - failQty;
                        this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                        this.lblTotalSN.Text = Convert.ToString(int.Parse(this.lblTotalSN.Text) + sns.Count() / 3 - scrapcount);
                        WriteIntoLocalFile(lblTextRackNo.Text, Pscount, this.lblTotalSN.Text, sn_att);
                        if (Pscount == Convert.ToInt32(config.MaxiumCount))
                        {
                            PrintLabelExt(sn_att);
                            Pscount = 0;
                            this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                            this.lblTotalSN.Text = "0";
                            lblTextRackNo.Text = GetNextSN();
                            if (File.Exists(@"LocalRackInfo.txt"))
                                File.Delete(@"LocalRackInfo.txt");
                        }
                    }
                    SetAlarmStatusText("通过");
                }
            }
            else
            {
                SetTopWindowMessage("", message("Upload the Comp.Name and Defect is error") + tempserialnumber);
                errorHandler(3, message("Upload the Comp.Name and Defect is error") + tempserialnumber, "Error");
            }
            dgvSNInfo.Rows.Clear();
            dgvSN.Rows.Clear();
            dgvCompName.Rows.Clear();
            dgvDefect.Rows.Clear();
            this.txbCDADataInput.Text = "";
            this.txbCDADataInput.Focus();
            tempserialnumber = "";

        }

        private void BtnaddFailure_Click(object sender, EventArgs e)
        {
            int rowCompName = 0;
            if (config.IsNeedCompColumn == "Y")
            {
                if (dgvSN.Rows.Count > 0 && dgvCompName.Rows.Count > 0 && dgvDefect.Rows.Count > 0)
                {
                }
                else
                {
                    MessageBox.Show(message("Pls load the data of SN and CompName and Defect"));
                    return;
                }
                if (dgvSN.SelectedRows.Count <= 0 || dgvCompName.SelectedRows.Count <= 0 || dgvDefect.SelectedRows.Count <= 0)//dgvCompName.CurrentRow == null || dgvDefect.CurrentRow == null
                {
                    MessageBox.Show(message("Pls select SN or CompName or Defect"));
                    return;
                }
                foreach (DataGridViewRow row in dgvCompName.Rows)
                {
                    if (row.Selected)
                        rowCompName = row.Index;// dgvCompName.CurrentRow.Index;
                }
            }
            else
            {
                if (dgvSN.Rows.Count > 0 && dgvDefect.Rows.Count > 0)
                {
                }
                else
                {
                    MessageBox.Show(message("Pls load the data of SN and Defect"));
                    return;
                }
                if (dgvSN.SelectedRows.Count > 0 || dgvDefect.SelectedRows.Count > 0)
                {
                    MessageBox.Show(message("Pls select SN or Defect"));
                    return;
                }
            }

            int rowSN = -1;// dgvSN.CurrentRow.Index;//获得选种行的索引
            foreach (DataGridViewRow row in dgvSN.Rows)
            {
                if (row.Selected)
                    rowSN = row.Index;// dgvCompName.CurrentRow.Index;
            }
            //int rowCompName = dgvCompName.CurrentRow.Index;
            int rowDefect = -1;// dgvDefect.CurrentRow.Index;
            foreach (DataGridViewRow row in dgvDefect.Rows)
            {
                if (row.Selected)
                    rowDefect = row.Index;// dgvCompName.CurrentRow.Index;
            }
            if (rowSN > -1 && rowSN < dgvSN.Rows.Count && rowCompName > -1 && (config.IsNeedCompColumn == "Y" ? rowCompName < dgvCompName.Rows.Count : rowCompName == dgvCompName.Rows.Count) && rowDefect > -1 && rowDefect < dgvDefect.Rows.Count)
            {
                string sn = dgvSN.Rows[rowSN].Cells[1].Value.ToString().Trim();
                string pos = dgvSN.Rows[rowSN].Cells[0].Value.ToString().Trim();
                string state = dgvSN.Rows[rowSN].Cells[2].Value.ToString().Trim();
                string compName = config.IsNeedCompColumn == "Y" ? dgvCompName.Rows[rowCompName].Cells[0].Value.ToString().Trim() : "";
                string desc = config.IsNeedCompColumn == "Y" ? dgvCompName.Rows[rowCompName].Cells[1].Value.ToString().Trim() : "";
                string defect = dgvDefect.Rows[rowDefect].Cells[0].Value.ToString().Trim();
                string defectId = dgvDefect.Rows[rowDefect].Cells[1].Value.ToString().Trim();
                List<string> listvalues = new List<string>();
                if (dgvSN.Rows[rowSN].DefaultCellStyle.BackColor == Color.Gray)
                {
                    MessageBox.Show(message("The scrap SN can not be selected"));
                    return;
                }
                else
                {
                    if (dgvSNInfo.Rows.Count > 0)
                    {
                        for (int i = 0; i < dgvSNInfo.Rows.Count; i++)
                        {
                            string snold = dgvSNInfo.Rows[i].Cells[1].Value.ToString().Trim();
                            string compNameOld = dgvSNInfo.Rows[i].Cells[2].Value.ToString().Trim();
                            string defectidold = dgvSNInfo.Rows[i].Cells[5].Value.ToString().Trim();
                            listvalues.Add(snold + "|" + compNameOld + "|" + defectidold);
                        }
                        if (listvalues.Contains(sn + "|" + compName + "|" + defectId))
                        {
                            MessageBox.Show(message("You have chosen the duplicate data"));
                            return;
                        }
                    }
                    dgvSNInfo.Rows.Add(new string[8] { pos, sn, compName, desc, defect, defectId, state, this.txtInfo.Text.Trim() });
                    dgvSN.Rows[rowSN].Cells[2].Value = EnumCommon.EnumResultState.FAIL.ToString();
                    //dgvSN.Rows[rowSN].Cells[2].Style.Format = FontStyle.Italic.ToString();
                    dgvSN.Rows[rowSN].Cells[2].Style.ForeColor = Color.Black;
                    dgvSN.Rows[rowSN].Cells[2].Style.Font = new Font(dgvSN.DefaultCellStyle.Font.FontFamily, dgvSN.DefaultCellStyle.Font.Size, FontStyle.Italic);
                    this.BtnaddFailure.Enabled = false;
                    this.dgvCompName.ClearSelection();
                    this.dgvDefect.ClearSelection();
                    this.dgvSNInfo.ClearSelection();
                    this.txtInfo.Text = "";
                    errorHandler(0, message("add_failure_success"), "");
                }
            }
            else
            {
                MessageBox.Show(message("Pls select SN or CompName or Defect"));
                return;
            }
        }
        System.Collections.ArrayList myLst2 = null;
        private void LoadFailureType()
        {
            dgvDefect.Rows.Clear();
            myLst2 = new System.Collections.ArrayList();
            GetFailureData getfailuredata = new GetFailureData(sessionContext, initModel, this);
            string[] snStateResultValues1 = { };
            snStateResultValues1 = getfailuredata.MdataGetFailureDataforStation();
            if (snStateResultValues1.Length > 0)
            {
                for (int i = 0; i < snStateResultValues1.Length / 2; i++)
                {
                    dgvDefect.Rows.Add(new string[2] { snStateResultValues1[2 * i + 1], snStateResultValues1[2 * i + 0] });
                    myLst2.Add(snStateResultValues1[2 * i + 1]);
                }
            }
            foreach (string m in myLst2)
            {
                this.cmbDefct.AutoCompleteCustomSource.Add(m);
            }
        }
        private void dgeSN(string serialnumber)
        {
            GetMaterialBinData materialHandler = new GetMaterialBinData(sessionContext, initModel, this);
            DataTable dt = materialHandler.GetBomMaterialData(serialnumber);
            if (dt != null && dt.Rows.Count > 0)
            {
                foreach (DataRow row in dt.Rows)
                {
                    if (row["ProductFlag"].ToString() == "1")
                        continue;
                    string CompName = row["CompName"].ToString();
                    dgvCompName.Rows.Add(new string[1] { CompName });
                }
            }

        }
        private void dgvSN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            //string[] snStateResultValues1 = { };
            //if (e.RowIndex > -1 && e.RowIndex < dgvSN.Rows.Count)
            //{
            //    dgvCompName.Rows.Clear();
            //    int indexrow = e.RowIndex;
            //    string sn = dgvSN.Rows[indexrow].Cells[1].Value.ToString();
            //    GetWorkOrder wo = new GetWorkOrder(sessionContext, initModel, this);
            //    string[] bomDataResultValues = wo.GetBomData(sn);
            //    if (bomDataResultValues.Length >= 0)
            //    {
            //        for (int i = 0; i < bomDataResultValues.Length; i++)
            //        {
            //            string CompName = bomDataResultValues[i];
            //            string[] mountingPlaces = CompName.Split(',');
            //            foreach (string mountPlace in mountingPlaces)
            //            {
            //                dgvCompName.Rows.Add(new string[1] { mountPlace });
            //            }
            //        }
            //    }
            //}
            //LoadFailureType();
        }
        System.Collections.ArrayList myLst = null;
        private void dgvSN_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            myLst = new System.Collections.ArrayList();
            string[] snStateResultValues1 = { };
            if (e.RowIndex > -1 && e.RowIndex < dgvSN.Rows.Count)
            {
                if (config.IsNeedCompColumn == "Y")
                {
                    dgvCompName.Rows.Clear();
                    int indexrow = e.RowIndex;
                    string sn = dgvSN.Rows[indexrow].Cells[1].Value.ToString();
                    GetWorkOrder wo = new GetWorkOrder(sessionContext, initModel, this);
                    string[] bomDataResultValues = new string[] { };
                    if (config.GetCompStation == "")
                    {
                        bomDataResultValues = wo.GetBomData(sn, config.StationNumber);
                    }
                    else
                    {
                        bomDataResultValues = wo.GetBomData(sn, config.GetCompStation.Trim());
                    }
                    if (bomDataResultValues.Length >= 0)
                    {
                        for (int i = 0; i < bomDataResultValues.Length / 4; i++)
                        {
                            string setupflag = bomDataResultValues[i * 4 + 2];
                            int processlayer = Convert.ToInt32(bomDataResultValues[i * 4 + 3]);
                            if (setupflag == "1")
                            {
                                if (config.IsCheckLayer == "Y")
                                {
                                    if (processlayer != iProcessLayer)
                                        continue;
                                }
                                string CompName = bomDataResultValues[i * 4];
                                string desc = bomDataResultValues[i * 4 + 1];
                                string[] mountingPlaces = CompName.Split(',');
                                foreach (string mountPlace in mountingPlaces)
                                {
                                    dgvCompName.Rows.Add(new string[2] { mountPlace, desc });
                                    myLst.Add(mountPlace);
                                }
                            }
                        }
                    }
                }
                foreach (string m in myLst)
                {
                    this.cmbComp.AutoCompleteCustomSource.Add(m);
                }
                LoadFailureType();
                this.dgvCompName.ClearSelection();
                this.dgvDefect.ClearSelection();
                this.dgvSNInfo.ClearSelection();
            }
            //this.Invoke(new MethodInvoker(delegate
            //{
            //    dgeSN();
            //    LoadFailureType();
            //}));
        }
        private bool UploadFailureTypeAndCompNew(int serailState, List<string> sns, List<string> snInfoValues)
        {
            string[] snStationSettings = new string[] { };
            bool flagReturn = true;
            int duplicateSerialNumber = 0;
            string[] failureValues = new string[] { };
            string[] failureReturnValues = new string[] { };
            int result1 = 0;
            bool b = false;
            Dictionary<string, List<string>> snAndFailureValues = new Dictionary<string, List<string>>();
            List<string> listFailureValue = null;
            UploadFailureState uploadstate = new UploadFailureState(sessionContext, initModel, this);
            int snstatenew = 0;
            if (serailState == 0) // Inspection all is OK
            {
                List<string> passlist = new List<string>();
                for (int i = 0; i < sns.Count / 3; i++)
                {
                    snstatenew = Convert.ToInt32(sns[3 * i + 2]);// == (int)EnumCommon.EnumResultState.SCRAP ? (int)EnumCommon.EnumResultState.SCRAP : (int)EnumCommon.EnumResultState.PASS;
                    if (snstatenew != 2 && snstatenew != -1)
                    {
                        passlist.Add("0");
                        passlist.Add(sns[3 * i + 1]);
                        passlist.Add(sns[3 * i + 0]);
                        passlist.Add("0");
                    }

                    //if(config.IsSelectWO!="Y")
                    //{
                    //    InitProcessLayer(sns[3 * i + 1]);
                    //}

                    //result1 = uploadstate.UploadFailureAndResultData(sns[3 * i + 1], sns[3 * i + 0], iProcessLayer, new string[] { }, failureValues, duplicateSerialNumber, 0, -1, snstatenew);

                    //if (result1 == 0 && (snstatenew == 0 || snstatenew == 2))
                    //{
                    //    //UpdateGridDataAfterUploadState();
                    //    continue;
                    //}
                    //else
                    //{
                    //    flagReturn = false;
                    //    break;
                    //}

                }
                if (passlist.ToArray().Length > 0)
                {
                    result1 = uploadstate.UploadProcessResultCall(passlist.ToArray(), iProcessLayer, 0);
                    if (result1 == 0)
                    {
                    }
                    else
                    {
                        flagReturn = false;
                    }
                }
            }
            else //Inspection has some fail 
            {
                List<string> passlist = new List<string>();
                for (int i = 0; i < snInfoValues.Count / 8; i++)
                {
                    string sn = snInfoValues[8 * i + 1];
                    string pos = snInfoValues[8 * i + 0];
                    string compName = snInfoValues[8 * i + 2];
                    string defect = snInfoValues[8 * i + 4];
                    string defectcode = snInfoValues[8 * i + 5];
                    string info = snInfoValues[8 * i + 7];
                    //string snstate = snInfoValues[6 * i + 5];

                    if (snAndFailureValues.ContainsKey(sn))
                    {
                        bool flag = snAndFailureValues.TryGetValue(sn, out listFailureValue);
                        if (flag)
                        {
                            listFailureValue.Add(compName);
                            listFailureValue.Add(0 + "");
                            listFailureValue.Add(defectcode);
                            listFailureValue.Add(info);
                            snAndFailureValues.Remove(sn);
                            snAndFailureValues.Add(sn, listFailureValue);

                        }
                    }
                    else
                    {
                        listFailureValue = new List<string>();
                        listFailureValue.Add(compName);
                        listFailureValue.Add(0 + "");
                        listFailureValue.Add(defectcode);
                        listFailureValue.Add(info);
                        snAndFailureValues.Add(sn, listFailureValue);
                    }
                }//处理数据结束
                for (int i = 0; i < sns.Count / 3; i++)
                {
                    string sn = sns[3 * i + 1];
                    if (snAndFailureValues.ContainsKey(sn))
                    {
                        bool flag = snAndFailureValues.TryGetValue(sn, out listFailureValue);
                        if (flag)
                        {
                            failureValues = AddListToArrayNew(listFailureValue);
                            //if (config.IsSelectWO != "Y")
                            //{
                            //    InitProcessLayer(sns[3 * i + 1]);
                            //}
                            result1 = uploadstate.UploadFailureAndResultData(sn, sns[3 * i + 0], iProcessLayer, new string[] { }, failureValues, duplicateSerialNumber, 0, -1, 1);

                            if (result1 == 0)
                            {
                                //UpdateGridDataAfterUploadState();
                                //allcount++;
                                continue;
                            }
                            else
                            {
                                flagReturn = false;
                                break;
                            }
                        }
                    }
                    else
                    {
                        snstatenew = Convert.ToInt32(sns[3 * i + 2]);// == (int)EnumCommon.EnumResultState.SCRAP ? (int)EnumCommon.EnumResultState.SCRAP : (int)EnumCommon.EnumResultState.PASS;
                        if (snstatenew != 2 && snstatenew != -1)
                        {
                            passlist.Add("0");
                            passlist.Add(sn);
                            passlist.Add(sns[3 * i + 0]);
                            passlist.Add("0");
                        }

                        //failureValues = new string[] { };

                        ////if (config.IsSelectWO != "Y")
                        ////{
                        ////    InitProcessLayer(sns[3 * i + 1]);
                        ////}
                        //result1 = uploadstate.UploadFailureAndResultData(sn, sns[3 * i + 0], iProcessLayer, new string[] { }, failureValues, duplicateSerialNumber, 0, -1, snstatenew);

                        //if (result1 == 0 && (snstatenew == 0 || snstatenew == 2))
                        //{
                        //    //UpdateGridDataAfterUploadState();
                        //    continue;
                        //}
                        //else
                        //{
                        //    flagReturn = false;
                        //    break;
                        //}
                    }
                }
                if (passlist.ToArray().Length > 0)
                {
                    result1 = uploadstate.UploadProcessResultCall(passlist.ToArray(), iProcessLayer, 0);
                    if (result1 == 0)
                    {
                    }
                    else
                    {
                        flagReturn = false;
                    }
                }
            }
            return flagReturn;
        }

        private string[] AddListToArrayNew(List<string> list)
        {
            string[] values = new string[list.Count];
            for (int m = 0; m < list.Count; m++)
            {
                values[m] = list[m];
            }
            return values;
        }

        private void gridCheckList_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (this.gridCheckList.Columns[e.ColumnIndex].Name == "clResult1" && this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString() != "")
            {
                //verify the input value
                string strRegex = @"^(\d{0,9}.\d{0,9})-(\d{0,9}.\d{0,9}).*$";
                string strResult1 = this.gridCheckList.Rows[e.RowIndex].Cells["clResult1"].Value.ToString();
                string strStandard = this.gridCheckList.Rows[e.RowIndex].Cells["clStandard"].Value.ToString();
                Match match = Regex.Match(strStandard, strRegex);
                if (match.Success)
                {
                    if (match.Groups.Count > 2)
                    {
                        double iMin = Convert.ToDouble(match.Groups[1].Value);
                        double iMax = Convert.ToDouble(match.Groups[2].Value);
                        double iResult = Convert.ToDouble(strResult1);
                        if (iResult >= iMin && iResult <= iMax)
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "OK";
                        }
                        else
                        {
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                            this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                        }
                    }
                }
                else
                {
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Style.BackColor = Color.Red;
                    this.gridCheckList.Rows[e.RowIndex].Cells["clStatus"].Value = "NG";
                }
            }
        }

        private void dgvCompName_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
            {
                this.BtnaddFailure.Enabled = true;
            }
        }

        private void dgvDefect_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (config.IsNeedCompColumn == "Y")
            {
                if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                {
                    this.BtnaddFailure.Enabled = true;
                }
                else
                {
                    this.BtnaddFailure.Enabled = false;
                }
            }
            else
            {
                if (dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                {
                    this.BtnaddFailure.Enabled = true;
                }
                else
                {
                    this.BtnaddFailure.Enabled = false;
                }
            }
        }

        private void dgvSNInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex == -1)
                return;
            if (dgvSNInfo.SelectedRows.Count > 0)
            {
                this.btnCancel.Enabled = true;
            }
            else
            {
                this.btnCancel.Enabled = false;
            }
        }


        #endregion

        #region check userskill
        private bool CheckUserSkillBySN(string serialNumber)
        {
            string strTempSN = serialNumber;
            if (this.txbCDAMONumber.Text.Contains(","))
            {
                strTempSN = serialNumber + "L";
            }
            CheckUserSkill userSkillHandler = new CheckUserSkill(sessionContext, initModel, this);
            int resultCode = userSkillHandler.CheckUserSkillForSN(UserName, strTempSN, iProcessLayer);
            if (resultCode != 0)
            {
                errorHandler(2, message("UserSkillInvalid"), "");
                return false;
            }
            return true;
        }

        private bool CheckUserSkillByWO(string workorder)
        {
            CheckUserSkill userSkillHandler = new CheckUserSkill(sessionContext, initModel, this);
            int resultCode = userSkillHandler.CheckUserSkillForWS(UserName, workorder, iProcessLayer);
            if (resultCode != 0)
            {
                errorHandler(2, message("UserSkillInvalid"), "");
                return false;
            }
            return true;
        }

        #endregion

        #region PanelMessage
        public delegate void SetAlarmStatusTextDel(string strText);
        public void SetAlarmStatusText(string strText)
        {
            if (txtConsole.InvokeRequired)
            {
                SetAlarmStatusTextDel setMsgDel = new SetAlarmStatusTextDel(SetAlarmStatusText);
                Invoke(setMsgDel, new object[] { strText });
            }
            else
            {
                if (string.IsNullOrEmpty(strText))
                {
                    this.lblMsg.Text = strText;
                    this.panelMsg.BackColor = Color.FromArgb(184, 255, 160);
                }
                else
                {
                    if (strText.ToUpper() == "通过")
                    {
                        this.lblMsg.Text = strText + System.Environment.NewLine + "PASS";
                        this.panelMsg.BackColor = Color.FromArgb(184, 255, 160);
                    }
                    else
                    {
                        this.lblMsg.Text = strText + System.Environment.NewLine + "FAIL";
                        this.panelMsg.BackColor = Color.Red;
                    }
                }
            }
        }
        #endregion

        #region PFC
        private System.Timers.Timer CheckConnectTimer = new System.Timers.Timer();
        private SocketClientHandler cSocket = null;
        private object _lock1 = new Object();
        public void ProcessPFCMessage(string pfcMsg)
        {
            //#!PONGCRLF
            //#!BCTOPBARCODECRLF
            //#!BCBOTTOMBARCODECRLF
            //#!BOARDAVCRLF
            //#!TRANSFERBARCODECRLF
            lock (_lock1)
            {
                //errorHandler(0, "Receive message from PFC " + pfcMsg.TrimEnd(), "");
                SetConnectionText(0, "Receive message from PFC " + pfcMsg.TrimEnd());
                LogHelper.Info("Receive message from PFC " + pfcMsg.TrimEnd());
                if (pfcMsg.Length >= 10)
                {
                    bool isOK = true;
                    string messageType = pfcMsg.Substring(2, 8).TrimEnd();
                    switch (messageType)
                    {
                        case "PONG":
                            PFCStartTime = DateTime.Now;
                            break;
                        case "BCTOP":
                            string serialNumber = pfcMsg.Substring(10).TrimEnd();
                            iProcessLayer = 0;
                            //Match match3 = Regex.Match(serialNumber, config.SDLExtractPattern);
                            //if (match3.Success)
                            //{
                            //    if (this.dgvEquipment.RowCount > 0)
                            //    {
                            //        if (!CheckEquipmentSetup())
                            //        {
                            //            return;
                            //        }
                            //    }
                            //    this.Invoke(new MethodInvoker(delegate
                            //    {
                            //        this.tabControl1.SelectedTab = this.tabDetail;
                            //    }));
                            //    isOK = ProcessSerialNumberSingle(match3.ToString());
                            //}
                            //else
                            //{
                            Match match = Regex.Match(serialNumber, config.DLExtractPattern);
                            if (match.Success)
                            {
                                if (this.dgvEquipment.RowCount > 0)
                                {
                                    if (!CheckEquipmentSetup())
                                    {
                                        return;
                                    }
                                }
                                this.Invoke(new MethodInvoker(delegate
                                {
                                    this.tabControl1.SelectedTab = this.tabDetail;
                                }));
                                isOK = ProcessSerialNumber(match.ToString());
                            }
                            //}

                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber);
                            }
                            break;
                        case "BCBOTTOM":
                            string serialNumber1 = pfcMsg.Substring(10).TrimEnd();
                            iProcessLayer = 1;
                            //Match match1 = Regex.Match(serialNumber1, config.SDLExtractPattern);
                            //if (match1.Success)
                            //{
                            //    if (this.dgvEquipment.RowCount > 0)
                            //    {
                            //        if (!CheckEquipmentSetup())
                            //        {
                            //            return;
                            //        }
                            //    }
                            //    this.Invoke(new MethodInvoker(delegate
                            //    {
                            //        this.tabControl1.SelectedTab = this.tabDetail;
                            //    }));
                            //    isOK = ProcessSerialNumberSingle(match1.ToString());
                            //}
                            //else
                            //{
                            Match match2 = Regex.Match(serialNumber1, config.DLExtractPattern);
                            if (match2.Success)
                            {
                                if (this.dgvEquipment.RowCount > 0)
                                {
                                    if (!CheckEquipmentSetup())
                                    {
                                        return;
                                    }
                                }
                                this.Invoke(new MethodInvoker(delegate
                                {
                                    this.tabControl1.SelectedTab = this.tabDetail;
                                }));
                                isOK = ProcessSerialNumber(match2.ToString());
                            }
                            //}

                            if (isOK)
                            {
                                SendMsessageToPFC(PFCMessage.GO, serialNumber1);
                            }
                            break;
                        case "BOARDAV"://todo
                            //isOK = ProcessSerialNumberData();
                            //if (isOK)
                            //{
                            //    SendMsessageToPFC(PFCMessage.GO, "");
                            //    BoardCome = false;
                            //}
                            break;
                        case "TRANSFER":
                            string serialNumber2 = pfcMsg.Substring(10).TrimEnd();
                            SendMsessageToPFC(PFCMessage.COMPLETE, serialNumber2);
                            SendMessageToCOM(serialNumber2);
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    MessageBox.Show("Receive message length less then 10.");
                }
            }
        }

        public delegate void SetConnectionTextDel(int typeOfError, string strMessage);
        public void SetConnectionText(int typeOfError, string strMessage)
        {
            if (txtConnection.InvokeRequired)
            {
                SetConnectionTextDel connectDel = new SetConnectionTextDel(SetConnectionText);
                Invoke(connectDel, new object[] { typeOfError, strMessage });
            }
            else
            {
                String errorBuilder = null;
                String isSucces = null;
                switch (typeOfError)
                {
                    case 0:
                        isSucces = "SUCCESS";
                        txtConnection.SelectionColor = Color.Black;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        //LogHelper.Info(strMessage);
                        break;
                    case 1:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        LogHelper.Error(strMessage);
                        break;
                    default:
                        isSucces = "FAIL";
                        txtConnection.SelectionColor = Color.Red;
                        errorBuilder = "# " + DateTime.Now.ToString("HH:mm:ss") + " >> " + isSucces + " >< " + strMessage + "\n";
                        break;
                }

                txtConnection.AppendText(errorBuilder);
                txtConnection.ScrollToCaret();
            }
        }

        private object _lockSend = new Object();
        public void SendMsessageToPFC(PFCMessage msgType, string serialNumber)
        {
            lock (_lockSend)
            {
                //#!PINGCRLF
                //#!GOBARCODECRLF
                //#!COMPLETEBARCODECRLF
                string prefix = "#!";
                string suffix = HexToStr1("0D") + HexToStr1("0A");
                string sendMessage = "";
                switch (msgType)
                {
                    case PFCMessage.PING:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                    case PFCMessage.GO:
                        sendMessage = prefix + PFCMessage.GO.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.COMPLETE:
                        sendMessage = prefix + PFCMessage.COMPLETE.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    case PFCMessage.CONFIRM:
                        sendMessage = prefix + PFCMessage.CONFIRM.ToString().PadRight(8, ' ') + serialNumber + suffix;
                        break;
                    default:
                        sendMessage = prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix;
                        break;
                }
                //send message through socket
                try
                {
                    if (DateTime.Now.Subtract(PFCStartTime).Seconds >= 20)
                    {
                        cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                        PFCStartTime = DateTime.Now;
                        Thread.Sleep(1000);
                    }
                    bool isOK = cSocket.send(sendMessage);
                    if (isOK)
                    {
                        //errorHandler(1, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    else
                    {
                        //errorHandler(2, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                        bool isConnectOK = cSocket.connect(config.IPAddress, config.Port);
                        if (isConnectOK)
                        {
                            isOK = cSocket.send(sendMessage);
                            if (isOK)
                            {
                                SetConnectionText(0, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                            else
                            {
                                SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                            }
                        }
                        else
                        {
                            SetConnectionText(1, "Conncet to PFC error");
                        }
                    }
                }
                catch (Exception ex)
                {
                    cSocket.send(prefix + PFCMessage.PING.ToString().PadRight(8, ' ') + suffix);
                    bool isOK = cSocket.send(sendMessage);
                    if (isOK)
                    {
                        //errorHandler(0, "Send message to PFC:" + sendMessage.TrimEnd(), "");
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    else
                    {
                        SetConnectionText(1, "Send message to PFC:" + sendMessage.TrimEnd());
                    }
                    LogHelper.Error(ex.Message, ex);
                }
            }
        }

        public static string HexToStr1(string mHex) // 返回十六进制代表的字符串
        {
            mHex = mHex.Replace(" ", "");
            if (mHex.Length <= 0) return "";
            byte[] vBytes = new byte[mHex.Length / 2];
            for (int i = 0; i < mHex.Length; i += 2)
                if (!byte.TryParse(mHex.Substring(i, 2), NumberStyles.HexNumber, null, out vBytes[i / 2]))
                    vBytes[i / 2] = 0;
            return ASCIIEncoding.Default.GetString(vBytes);
        }

        public void GetTimerStart()
        {
            // 循环间隔时间(1分钟)
            CheckConnectTimer.Interval = 2 * 1000;
            // 允许Timer执行
            CheckConnectTimer.Enabled = true;
            // 定义回调
            CheckConnectTimer.Elapsed += new ElapsedEventHandler(CheckConnectTimer_Elapsed);
            // 定义多次循环
            CheckConnectTimer.AutoReset = true;
        }

        private void CheckConnectTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SendMsessageToPFC(PFCMessage.PING, "");
        }

        private void SendMessageToCOM(string msg)
        {
            try
            {
                initModel.scannerHandler.handler().Write(msg);
                LogHelper.Info("Send message to COM :" + msg);
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex);
            }
        }
        #endregion

        #region 控制轨道 controlbox
        private static byte[] strToToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0)
                hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2), 16);
            return returnBytes;
        }

        public static int CRC16(byte[] str, int len)
        {
            int Base = 0xFFFF;
            int Oper = 0xA001;
            for (int i = 0; i < len; i++)
            {
                Base = ((int)str[i]) ^ Base;
                for (int j = 0; j < 8; j++)
                {
                    if ((Base & 0x01) != 0)
                    {
                        Base = Base >> 1;
                        Base = Base ^ Oper;
                    }
                    else
                    {
                        Base = Base >> 1;
                    }
                }
            }
            return Base;
        }

        private string CRCCommand(string command)//自动计算校验码
        {
            string str = "";
            byte[] byt = strToToHexByte(command);

            int rtdata = CRC16(byt, byt.Length);
            string end = Convert.ToString(rtdata, 16).ToUpper().PadLeft(4, '0');
            str = command + " " + end.Substring(2, 2) + " " + end.Substring(0, 2);

            return str;
        }

        private string ReturnOpenCommand(string CoilAddress)
        {
            string iValue = "";
            switch (CoilAddress)
            {
                case "Y0":
                    iValue = CRCCommand("01 05 00 00 FF 00");
                    break;
                case "Y1":
                    iValue = CRCCommand("01 05 00 01 FF 00");
                    break;
                case "Y2":
                    iValue = CRCCommand("01 05 00 02 FF 00");
                    break;
                case "Y3":
                    iValue = CRCCommand("01 05 00 03 FF 00");
                    break;
                case "Y4":
                    iValue = CRCCommand("01 05 00 04 FF 00");
                    break;
                case "Y5":
                    iValue = CRCCommand("01 05 00 05 FF 00");
                    break;
                case "Y6":
                    iValue = CRCCommand("01 05 00 06 FF 00");
                    break;
                case "Y7":
                    iValue = CRCCommand("01 05 00 07 FF 00");
                    break;
                default:
                    iValue = "";
                    break;
            }
            return iValue;
        }

        private string ReturnCloseCommand(string CoilAddress)
        {
            string iValue = "";
            switch (CoilAddress)
            {
                case "Y0":
                    iValue = CRCCommand("01 05 00 00 00 00");
                    break;
                case "Y1":
                    iValue = CRCCommand("01 05 00 01 00 00");
                    break;
                case "Y2":
                    iValue = CRCCommand("01 05 00 02 00 00");
                    break;
                case "Y3":
                    iValue = CRCCommand("01 05 00 03 00 00");
                    break;
                case "Y4":
                    iValue = CRCCommand("01 05 00 04 00 00");
                    break;
                case "Y5":
                    iValue = CRCCommand("01 05 00 05 00 00");
                    break;
                case "Y6":
                    iValue = CRCCommand("01 05 00 06 00 00");
                    break;
                case "Y7":
                    iValue = CRCCommand("01 05 00 07 00 00");
                    break;
                default:
                    iValue = "";
                    break;
            }
            return iValue;
        }

        #endregion

        #region 首件检查
        string IPIstatus = "";
        //static string cacheWO = "";
        private void CheckIPIStatus()
        {
            //if(cacheWO!=this.txbCDAMONumber.Text)
            //{
            //    cacheWO = this.txbCDAMONumber.Text;
            if (config.IPI_STATUS_CHECK == "ENABLE" || config.Production_Inspection_CHECK == "ENABLE")
            {
                GetWorkPlanData workplanHandle = new GetWorkPlanData(sessionContext, initModel, this);
                string[] workplanDataResultValues = workplanHandle.GetWorkplanDataForStation(this.txbCDAMONumber.Text);
                string workplanid = workplanDataResultValues[2];
                string workstep = workplanDataResultValues[3];

                GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                string[] attributeResultValues = getattribute.GetAttributeValueForWorkStep("IPI_STATE_UPDATE", workplanid, workstep);
                if (attributeResultValues.Length > 0)
                {
                    if (attributeResultValues[1] == "Y")
                    {
                        string[] valuesAttri = getattribute.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS");
                        if (valuesAttri != null && valuesAttri.Length > 0)
                        {
                            IPIstatus = valuesAttri[1];
                        }
                    }
                }
            }

            //}
        }
        private void UpdateIPIStatus(string result)
        {
            if (config.IPI_STATUS_CHECK == "ENABLE")
            {
                if (IPIstatus == "0")
                {
                    AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                    if (result == "0")//成功
                        appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "1");
                    else//失败
                        appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "-1");

                    IPIstatus = "";
                }
            }
        }
        #endregion

        #region print label
        private void PrintLabelExt(string sn)
        {
            //ZblLabBuild LabBuild = new ZblLabBuild();
            //string content = LabBuild.BuildLabContent(this.txtPartNumber.Text, this.txtLotNumber.Text, this.txtDateCode.Text, this.txtQty.Text);

            if (string.IsNullOrEmpty(TemplateText))
            {
                TemplateText = ReadLabelTemplate(@"" + config.TemplateFolder);
            }
            string content = GeneratePrintString(TemplateText, sn);
            Byte[] bytes = Encoding.UTF8.GetBytes(content);
            if (config.PrinterTypeName == "ZPL")
            {
                bytes = Encoding.Default.GetBytes(content);
            }
            FileStream fs = new FileStream(@"Last_TravelerSlip.txt", FileMode.Create);
            //byte[] bt = Encoding.UTF8.GetBytes(bytes.ToString());
            fs.Seek(0, SeekOrigin.Begin);
            fs.Write(bytes, 0, bytes.Length);
            fs.Flush();
            fs.Close();
            try
            {
                if (config.PrinterTypeName == "ZPL")
                {
                    if (config.PrintType == "COM")
                    {
                        USBControl.comPrint(config.PrintSerialPort, bytes);
                    }
                    else
                    {
                        ZebraPrintHelper.SendBytesToPrinter(this.cbxPrinter.Text, bytes);
                        //USBControl.SendBytesToPrinter(this.cbxPrinter.Text, @"Last_TravelerSlip.txt");//.SendBytesToPrinter(this.cbxPrinter.Text, bytes,0);

                    }
                }
                else
                {
                    if (config.PrintType == "COM")
                    {
                        USBControl.comPrint(config.PrintSerialPort, bytes);
                    }
                    else
                    {
                        try
                        {
                            streamToPrint = new StreamReader(@"Last_TravelerSlip.txt");
                            PrintDocument printDocument = new PrintDocument();
                            printDocument.PrinterSettings.PrinterName = this.cbxPrinter.Text;
                            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
                            printDocument.Print();
                        }
                        finally
                        {
                            streamToPrint.Close();
                        }
                    }
                }

                errorHandler(0, message("Print success"), "");
            }
            catch (Exception ex)
            {
                errorHandler(2, message("Print fail"), "");
                LogHelper.Error(ex);
            }
        }

        private void InitPrinterList()
        {
            //获取本地连接打印机列表加载到下拉框中
            PrinterSettings.StringCollection list = PrinterSettings.InstalledPrinters;
            foreach (string pkInstalledPrinters in list)
            {
                cbxPrinter.Items.Add(pkInstalledPrinters);
                //本地默认的打印机为默认选择项
                PrintDocument prtdoc = new PrintDocument();
                string strDefaultPrinter = prtdoc.PrinterSettings.PrinterName;//获取默认的打印机名 
                if (pkInstalledPrinters == strDefaultPrinter)
                //把本地默认打印机设为缺省值 
                {
                    cbxPrinter.SelectedIndex = cbxPrinter.Items.IndexOf(pkInstalledPrinters);
                }
            }
        }
        string TemplateText = "";
        private void cmbLabelTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string filePath = @"" + config.TemplateFolder;
            if (filePath != "System.Data.DataRowView")
            {
                TemplateText = ReadLabelTemplate(filePath);
            }
        }
        private string ReadLabelTemplate(string filePath)
        {
            string[] contents = File.ReadAllLines(filePath, Encoding.UTF8);
            if (config.PrinterTypeName == "ZPL")
                contents = File.ReadAllLines(filePath, Encoding.Default);
            StringBuilder sb = new StringBuilder();
            if (contents != null && contents.Length > 0)
            {
                foreach (var item in contents)
                {
                    sb.AppendLine(item);
                }
            }
            return sb.ToString();
        }

        string sn_att = "";
        private string GeneratePrintString(string text, string sn)
        {
            if (text.Contains(@",,WO.NO"))//WorkOrder
            {
                text = text.Replace(@",,WO.NO", txbCDAMONumber.Text).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,WO.PN_NO"))//PartNumber
            {
                text = text.Replace(@",,WO.PN_NO", txbCDAPartNumber.Text).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,WO.PN_DESC"))//PartNumber
            {
                text = text.Replace(@",,WO.PN_DESC", partdesc).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            //if (text.Contains(@"{MB.Quantity}"))//PartNumber
            //{
            //    text = text.Replace(@"{MB.Quantity}", config.MaxiumCount);
            //}

            if (text.Contains(@",,DATE_TIME[]")) //datetime eg. 2016-07-14 21:00:00
            {
                text = text.Replace(@",,DATE_TIME[]", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,USER_NAME")) //username
            {
                text = text.Replace(@",,USER_NAME", lblUser.Text).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,LINE.NO")) //lineno
            {
                string lineno = "";
                GetWorkOrder WOHandle = new GetWorkOrder(sessionContext, initModel, this);
                string[] lineinfo = WOHandle.GetLineNumber(config.StationNumber);
                if (lineinfo.Length > 0)
                {
                    lineno = lineinfo[0];
                }
                text = text.Replace(@",,LINE.NO", lineno).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,LINE.DESC")) //lineno
            {
                string linedesc = "";
                GetWorkOrder WOHandle = new GetWorkOrder(sessionContext, initModel, this);
                string[] lineinfo = WOHandle.GetLineNumber(config.StationNumber);
                if (lineinfo.Length > 0)
                {
                    linedesc = lineinfo[1];
                }
                text = text.Replace(@",,LINE.DESC", linedesc).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,STN.NO")) //station number
            {
                text = text.Replace(@",,STN.NO", config.StationNumber).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,STN.DESC")) //station desc 如果没有激活工单则无法抓取
            {
                text = text.Replace(@",,STN.DESC", initModel.currentSettings.stationDesc).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,DATE[]")) //date eg. 2016-07-14
            {
                text = text.Replace(@",,DATE[]", DateTime.Now.ToString("yyyy-MM-dd")).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,STN.NO")) //station number
            {
                text = text.Replace(@",,COUNT[SN.ATTR(attribute code)", config.StationNumber).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,COUNT[")) //Quantity in Magazine Rack
            {
                string regular = @"(,,COUNT\[SN.ATTR\((\S.*)\)\])";
                string replacetext = "";
                string attributecode = "";
                string attributevalue = "";
                Match match = Regex.Match(text, regular);
                if (match.Success)
                {
                    replacetext = match.ToString();
                    if (match.Groups.Count > 1)
                    {
                        attributecode = match.Groups[2].ToString();
                        replacetext = match.Groups[1].ToString();
                    }
                    GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                    string[] attributeResultValues = getattribute.GetAttributeValueForAll(0, sn, "-1", attributecode);
                    if (attributeResultValues.Length > 0)
                    {
                        attributevalue = attributeResultValues[1];
                    }
                }

                if (attributecode != "")
                {
                    text = text.Replace(@"" + replacetext, attributevalue).Replace("\"", "").Replace("{", "").Replace("}", "");
                }
                else
                {
                    text = text.Replace(@",,COUNT[]", Pscount.ToString()).Replace("\"", "").Replace("{", "").Replace("}", "");//config.MaxiumCount
                }
            }
            if (text.Contains(@",,PANEL.COUNT[]")) //station number
            {
                text = text.Replace(@",,PANEL.COUNT[]", this.lblTotalSN.Text).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,WO.ATTR")) //Work Order Attribute, attribute code
            {
                string regular = @"(,,WO.ATTR\((\S.*)\))";
                string replacetext = "";
                string attributecode = "";
                string attributevalue = "";
                Match match = Regex.Match(text, regular);
                if (match.Success)
                {
                    replacetext = match.ToString();
                    if (match.Groups.Count > 1)
                    {
                        attributecode = match.Groups[2].ToString();
                        replacetext = match.Groups[1].ToString();
                    }
                }
                GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                string[] attributeResultValues = getattribute.GetAttributeValueForAll(1, txbCDAMONumber.Text, "-1", attributecode);
                if (attributeResultValues.Length > 0)
                {
                    attributevalue = attributeResultValues[1];
                }

                text = text.Replace(@"" + replacetext, attributevalue).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,SN.ATTR")) //Work Order Attribute, attribute code
            {
                string regular = @"(,,SN.ATTR\((\S.*)\))";
                string replacetext = "";
                string attributecode = "";
                string attributevalue = "";
                Match match = Regex.Match(text, regular);
                if (match.Success)
                {
                    replacetext = match.ToString();
                    if (match.Groups.Count > 1)
                    {
                        attributecode = match.Groups[2].ToString();
                        replacetext = match.Groups[1].ToString();
                    }
                }
                GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                string[] attributeResultValues = getattribute.GetAttributeValueForAll(0, sn, "-1", attributecode);
                if (attributeResultValues.Length > 0)
                {
                    attributevalue = attributeResultValues[1];
                }

                text = text.Replace(@"" + replacetext, attributevalue).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            if (text.Contains(@",,Rack.NO"))//通箱号
            {
                text = text.Replace(@",,Rack.NO", lblTextRackNo.Text).Replace("\"", "").Replace("{", "").Replace("}", "");
            }
            return text;
        }

        private string GetNextSN()
        {
            string rackno = "";
            GetNextSerialNumber NTSHandle = new GetNextSerialNumber(sessionContext, initModel, this);
            SerialNumberData[] serialNumberArray = new SerialNumberData[] { };
            serialNumberArray = NTSHandle.GetSerialNumber(config.Temp_PartNo, 1);
            if (serialNumberArray.Length > 0)
                rackno = serialNumberArray[0].serialNumber;

            return rackno;
        }
        StreamReader streamToPrint = null;

        //打印最后保存到本地的文本Last_TravelerSlip.txt
        private void btnPrintLast_Click(object sender, EventArgs e)
        {
            if (!File.Exists(@"Last_TravelerSlip.txt"))
            {
                errorHandler(2, message("No file need to print"), "");
                return;
            }

            string LastTemplate = ReadLabelTemplate(@"Last_TravelerSlip.txt");
            Byte[] bytes = Encoding.UTF8.GetBytes(LastTemplate);
            if (config.PrinterTypeName == "ZPL")
            {
                bytes = Encoding.Default.GetBytes(LastTemplate);
            }
            try
            {
                if (config.PrinterTypeName == "ZPL")
                {
                    if (config.PrintType == "COM")
                    {
                        USBControl.comPrint(config.PrintSerialPort, bytes);
                    }
                    else
                    {
                        ZebraPrintHelper.SendBytesToPrinter(this.cbxPrinter.Text, bytes);
                        //USBControl.SendBytesToPrinter(this.cbxPrinter.Text, @"Last_TravelerSlip.txt");//.SendBytesToPrinter(this.cbxPrinter.Text, bytes,0);

                    }
                }
                else
                {
                    if (config.PrintType == "COM")
                    {
                        USBControl.comPrint(config.PrintSerialPort, bytes);
                    }
                    else
                    {
                        try
                        {
                            streamToPrint = new StreamReader(@"Last_TravelerSlip.txt");
                            PrintDocument printDocument = new PrintDocument();
                            printDocument.PrinterSettings.PrinterName = this.cbxPrinter.Text;
                            printDocument.PrintPage += new PrintPageEventHandler(printDocument_PrintPage);
                            printDocument.Print();
                        }
                        finally
                        {
                            streamToPrint.Close();
                        }
                    }
                }

                errorHandler(0, message("Print success"), "");
            }
            catch (Exception ex)
            {
                errorHandler(2, message("Print fail"), "");
                LogHelper.Error(ex);
            }
        }

        private void printDocument_PrintPage(object sender, PrintPageEventArgs ev)
        {
            float linesPerPage = 0;
            float yPos = 0;
            int count = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            String line = null;
            Font printFont = new Font("宋体", 12);
            // Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height /
               printFont.GetHeight(ev.Graphics);
            // Iterate over the file, printing each line. 
            while (count < linesPerPage &&
                ((line = streamToPrint.ReadLine()) != null))
            {
                if (line.Length > 0)
                {
                    yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));
                    ev.Graphics.DrawString(line, new Font(new FontFamily("宋体"), 11), Brushes.Black,
                       leftMargin, yPos, new StringFormat());
                }

                count++;
            }

            // If more lines exist, print another page. 
            if (line != null)
                ev.HasMorePages = true;
            else
                ev.HasMorePages = false;
        }

        //当未达到最大打印数量时，可强制打印
        private void btnCloseRack_Click(object sender, EventArgs e)
        {
            if (config.IsPrint == "Enable")
            {
                if (Pscount > 0)
                {
                    PrintLabelExt(sn_att);
                    Pscount = 0;
                    this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                    this.lblTotalSN.Text = "0";
                    lblTextRackNo.Text = GetNextSN();
                    if (File.Exists(@"LocalRackInfo.txt"))
                        File.Delete(@"LocalRackInfo.txt");
                }
                else
                {
                    errorHandler(2, message("PCB can not print"), "");
                }
            }
        }

        //防止如果程序意外关闭时数据丢失，将Rack信息存到本地文件中，打印完则删除
        private void WriteIntoLocalFile(string rackno, int PCBcount, string panlecount, string sn)
        {
            try
            {
                FileStream fs1 = new FileStream(@"LocalRackInfo.txt", FileMode.Append, FileAccess.Write, FileShare.ReadWrite);
                StreamWriter sw = new StreamWriter(fs1);

                sw.WriteLine(rackno + "," + PCBcount + "," + panlecount + "," + this.txbCDAMONumber.Text + "," + sn);

                sw.Close();
                fs1.Close();

                LogHelper.Info(message("write serialnumber success") + rackno + "," + PCBcount + "," + panlecount + "," + this.txbCDAMONumber.Text + "," + sn);
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return;
            }
        }

        //打开程序或者目检PASS都要先检查有没有Rack本地文件，有的话在LocalRackInfo.txt的基础上进行打印
        private void CheckLocalFile()
        {
            string content = "";
            if (File.Exists(@"LocalRackInfo.txt"))
            {
                string[] Linelists = File.ReadAllLines(@"LocalRackInfo.txt", Encoding.Default);
                content = Linelists[Linelists.Length - 1];

                if (content != "")
                {
                    string[] cons = content.Split(',');
                    string wo = cons[3];
                    if (wo == txbCDAMONumber.Text)
                    {
                        this.lblTextRackNo.Text = cons[0];
                        this.lblTotalSN.Text = cons[2];
                        this.lblPCBQty.Text = cons[1] + "/" + config.MaxiumCount;
                        sn_att = cons[4];
                    }
                }
            }
        }
        #endregion

        #region active wo
        private void InitWorkOrderList()
        {
            GetWorkOrder getWOHandler = new GetWorkOrder(sessionContext, initModel, this);
            DataTable dt = getWOHandler.GetAllWorkordersExt();
            DataView dv = dt.DefaultView;
            dv.Sort = "Info desc";
            dt = dv.Table;
            if (dt != null)
            {
                this.gridWorkorder.DataSource = dt;
                this.gridWorkorder.ClearSelection();
            }
            for (int i = 0; i < gridWorkorder.Rows.Count; i++)
            {
                gridWorkorder.Rows[i].Cells["columnRunId"].Value = i + 1 + "";
            }
        }

        private delegate void SetWorkorderGridStatusHandle();
        private void SetWorkorderGridStatus()
        {
            if (this.gridWorkorder.InvokeRequired)
            {
                SetWorkorderGridStatusHandle setStatusDel = new SetWorkorderGridStatusHandle(SetWorkorderGridStatus);
                Invoke(setStatusDel, new object[] { });
            }
            else
            {
                for (int i = 0; i < this.gridWorkorder.Rows.Count; i++)
                {
                    if (this.txbCDAMONumber.Text.Trim() == this.gridWorkorder.Rows[i].Cells["columnWoNumber"].Value.ToString())
                    {
                        ((DataGridViewImageCell)gridWorkorder.Rows[i].Cells["Activated"]).Value = ICTClient.Properties.Resources.ok;
                    }
                    else
                    {
                        ((DataGridViewImageCell)gridWorkorder.Rows[i].Cells["Activated"]).Value = ICTClient.Properties.Resources.Close;
                    }
                }
            }
        }

        private void btnActivate_Click(object sender, EventArgs e)
        {
            //if (config.IsCheckList == "Y")
            //{
            //    if (!CheckCheckList())
            //    {
            //        return;
            //    }
            //}
            string workorder = "";
            strShiftChecklist = "";
            bool isInitChecklist = false;
            string WorkorderPre = txbCDAMONumber.Text;
            int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
            if (this.gridWorkorder.SelectedRows.Count > 0)
            {
                workorder = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                ActivateWorkorder activateHandler = new ActivateWorkorder(sessionContext, initModel, this);
                int error = activateHandler.ActivateWorkorderExt(initModel.configHandler.StationNumber, workorder, 1, ConvertProcessLayerToString(this.cmbLayer.Text));//1 = Activate work order for the station only
                if (error == 0)
                {
                    if (this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString() != this.txbCDAMONumber.Text)
                    {
                        Pscount = 0;
                        this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                        this.lblTotalSN.Text = "0";
                        lblTextRackNo.Text = GetNextSN();
                        if (File.Exists(@"LocalRackInfo.txt"))
                            File.Delete(@"LocalRackInfo.txt");

                    }
                    this.txbCDAMONumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                    this.txbCDAPartNumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnPn"].Value.ToString();
                    this.txtLayer.Text = ConvertProcessLayerToString2(ConvertProcessLayerToString(this.cmbLayer.Text));
                    GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
                    GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
                    initModel.currentSettings = model;
                    if (model != null && model.workorderNumber != null)
                    {
                        GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                        List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                        if (listData != null && listData.Count > 0)
                        {
                            MdataGetPartData mData = listData[0];
                            initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                        }
                    }
                    LoadYield();
                    if (workorder != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                    {
                        strShiftChecklist = "";
                        InitWorkOrderType();
                        InitShift2(WorkorderPre);
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                InitTaskData_SOCKET("开线点检");
                                isStartLineCheck = true;
                            }

                        }
                        isInitChecklist = true;
                    }
                    InitSetupGrid();
                    InitWorkOrderList();
                    SetWorkorderGridStatus();
                    PanelQty = GetPCBPanelQty();
                    iProcessLayer = initModel.currentSettings.processLayer;
                    if (config.IsMaterialSetup == "Y")
                    {
                        InitSetupGrid();
                    }
                    if (config.IsEquipSetup == "Y")
                    {
                        InitEquipmentGrid();
                        SetupMachineAuto();
                    }
                    InitCintrolLanguage();
                    InitDocumentGrid();
                    if (!isInitChecklist)
                    {
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            InitShift2(WorkorderPre);
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                            }
                            else
                            {
                                if (!ReadCheckListFile())
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                        }
                        else
                        {
                            InitTaskData();
                        }
                    }
                    SetAlarmStatusText("");
                    SetTipMessage(MessageType.OK, message("Activated work order success."));
                }
            }
        }

        private void btnActivateExt_Click(object sender, EventArgs e)
        {
            //if (config.IsCheckList == "Y")
            //{
            //    if (!CheckCheckList())
            //    {
            //        return;
            //    }
            //}
            string workorder = "";
            strShiftChecklist = "";
            bool isInitChecklist = false;
            string WorkorderPre = txbCDAMONumber.Text;
            int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
            if (this.gridWorkorder.SelectedRows.Count > 0)
            {
                workorder = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                ActivateWorkorder activateHandler = new ActivateWorkorder(sessionContext, initModel, this);
                int error = activateHandler.ActivateWorkorderExt(initModel.configHandler.StationNumber, workorder, 2, ConvertProcessLayerToString(this.cmbLayer.Text));//2 = Activate work order for entire line
                if (error == 0)
                {
                    if (this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString() != this.txbCDAMONumber.Text)
                    {
                        Pscount = 0;
                        this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                        this.lblTotalSN.Text = "0";
                        lblTextRackNo.Text = GetNextSN();
                        if (File.Exists(@"LocalRackInfo.txt"))
                            File.Delete(@"LocalRackInfo.txt");
                    }
                    this.txbCDAMONumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnWoNumber"].Value.ToString();
                    this.txbCDAPartNumber.Text = this.gridWorkorder.SelectedRows[0].Cells["columnPn"].Value.ToString();
                    this.txtLayer.Text = ConvertProcessLayerToString2(ConvertProcessLayerToString(this.cmbLayer.Text));
                    GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
                    GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
                    initModel.currentSettings = model;
                    if (model != null && model.workorderNumber != null)
                    {
                        GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                        List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                        if (listData != null && listData.Count > 0)
                        {
                            MdataGetPartData mData = listData[0];
                            initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                        }
                    }
                    LoadYield();
                    if (workorder != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                    {
                        strShiftChecklist = "";
                        InitWorkOrderType();
                        InitShift2(WorkorderPre);
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                InitTaskData_SOCKET("开线点检");
                                isStartLineCheck = true;
                            }
                        }
                        isInitChecklist = true;
                    }
                    InitSetupGrid();
                    InitWorkOrderList();
                    SetWorkorderGridStatus();
                    PanelQty = GetPCBPanelQty();
                    iProcessLayer = initModel.currentSettings.processLayer;
                    if (config.IsMaterialSetup == "Y")
                    {
                        InitSetupGrid();
                    }
                    if (config.IsEquipSetup == "Y")
                    {
                        InitEquipmentGrid();
                        SetupMachineAuto();
                    }
                    InitCintrolLanguage();
                    InitDocumentGrid();
                    if (!isInitChecklist)
                    {
                        if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                        {
                            InitShift2(WorkorderPre);
                            if (!CheckShiftChange2())
                            {
                                InitTaskData_SOCKET("开线点检;设备点检");
                                isStartLineCheck = true;
                            }
                            else
                            {
                                if (!ReadCheckListFile())
                                {
                                    InitTaskData_SOCKET("开线点检");
                                    isStartLineCheck = true;
                                }
                            }
                        }
                        else
                        {
                            InitTaskData();
                        }
                    }
                    SetAlarmStatusText("");
                    SetTipMessage(MessageType.OK, message("Activated work order success."));
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //if (config.IsCheckList == "Y")
            //{
            //    if (!CheckCheckList())
            //    {
            //        return;
            //    }
            //}
            string WorkorderPre = txbCDAMONumber.Text;
            int processlayerPre = initModel.currentSettings == null ? -1 : initModel.currentSettings.processLayer;
            strShiftChecklist = "";//20161215 add by qy
            GetCurrentWorkorder getCurrentHandler = new GetCurrentWorkorder(sessionContext, initModel, this);
            GetStationSettingModel model = getCurrentHandler.GetCurrentWorkorderResultCall();
            initModel.currentSettings = model;
            if (model != null && model.workorderNumber != null)
            {
                GetNumbersOfSingleBoards getNumBoard = new GetNumbersOfSingleBoards(sessionContext, initModel, this);
                List<MdataGetPartData> listData = getNumBoard.GetNumbersOfSingleBoardsResultCall(model.partNumber);
                if (listData != null && listData.Count > 0)
                {
                    MdataGetPartData mData = listData[0];
                    initModel.numberOfSingleBoards = mData.quantityMultipleBoard;
                }
                if (model.workorderNumber != this.txbCDAMONumber.Text)
                {
                    Pscount = 0;
                    this.lblPCBQty.Text = Pscount.ToString() + "/" + config.MaxiumCount;
                    this.lblTotalSN.Text = "0";
                    lblTextRackNo.Text = GetNextSN();
                    if (File.Exists(@"LocalRackInfo.txt"))
                        File.Delete(@"LocalRackInfo.txt");
                }
                this.txbCDAMONumber.Text = model.workorderNumber;
                this.txbCDAPartNumber.Text = model.partNumber;
                this.txtLayer.Text = ConvertProcessLayerToString2(initModel.currentSettings.processLayer.ToString());
                LoadYield();
                bool isInitChecklist = false;
                if (model.workorderNumber != WorkorderPre || processlayerPre != initModel.currentSettings.processLayer)
                {
                    strShiftChecklist = "";
                    InitWorkOrderType();
                    InitShift2(WorkorderPre);//20161215 add by qy
                    if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                    {
                        if (!CheckShiftChange2())
                        {
                            InitTaskData_SOCKET("开线点检;设备点检");
                            isStartLineCheck = true;
                        }
                        else
                        {
                            InitTaskData_SOCKET("开线点检");
                            isStartLineCheck = true;
                        }
                    }
                    isInitChecklist = true;
                }
                InitSetupGrid();
                InitWorkOrderList();
                SetWorkorderGridStatus();
                PanelQty = GetPCBPanelQty();
                iProcessLayer = initModel.currentSettings.processLayer;
                if (config.IsMaterialSetup == "Y")
                {
                    InitSetupGrid();
                }
                if (config.IsEquipSetup == "Y")
                {
                    InitEquipmentGrid();
                    SetupMachineAuto();
                }
                InitCintrolLanguage();
                InitDocumentGrid();
                if (!isInitChecklist)
                {
                    if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                    {
                        InitShift2(WorkorderPre);
                        if (!CheckShiftChange2())
                        {
                            InitTaskData_SOCKET("开线点检;设备点检");
                            isStartLineCheck = true;
                        }
                        else
                        {
                            if (!ReadCheckListFile())
                            {
                                InitTaskData_SOCKET("开线点检");
                                isStartLineCheck = true;
                            }
                        }
                    }
                    else
                    {
                        InitTaskData();
                    }
                }
                SetAlarmStatusText("");
                SetTipMessage(MessageType.OK, message("Activated work order success."));
            }
            else
            {
                this.txbCDAMONumber.Text = "";
                this.txbCDAPartNumber.Text = "";
            }
        }

        #endregion

        #region production inspection
        private void UpdateIPIStatusForProductionInspection(string result, string serialnumber)
        {
            if (config.Production_Inspection_CHECK == "ENABLE")
            {
                if (IPIstatus == "1")
                {
                    UploadFailureState getHandle = new UploadFailureState(sessionContext, initModel, this);
                    int error = getHandle.GetSerialNumberByref(serialnumber);
                    if (error == -203)
                        serialnumber = serialnumber.Substring(0, serialnumber.Length - 3);

                    GetAttributeValue getattribute = new GetAttributeValue(sessionContext, initModel, this);
                    string[] valuesAttri = getattribute.GetAttributeValueForAll(0, serialnumber, "-1", "IPI");
                    if (valuesAttri != null && valuesAttri.Length > 0)
                    {
                        AppendAttribute appendAttri = new AppendAttribute(sessionContext, initModel, this);
                        if (result != "0")
                            appendAttri.AppendAttributeForAll(1, this.txbCDAMONumber.Text, "-1", "IPI_STATUS", "-2");

                        IPIstatus = "";
                    }

                }
            }
        }

        #endregion

        #region
        string OKlist = "";
        string NGlist = "";
        private void InitCheckResultMapping()
        {
            string[] LineList = File.ReadAllLines("CheckResultMappingFile.txt", Encoding.Default);
            foreach (var line in LineList)
            {
                if (!string.IsNullOrEmpty(line))
                {
                    string[] strs = line.Split(new char[] { ';' });
                    if (strs[0] == "OK")
                    {
                        OKlist = OKlist + "," + strs[1];
                    }
                    else
                    {
                        NGlist = NGlist + "," + strs[1];
                    }
                }
            }
        }
        #endregion

        #region 模糊查询
        List<string> listNew = new List<string>();
        private void cmbComp_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                this.cmbComp.Items.Clear();
                listNew.Clear();
                foreach (var item in myLst)
                {
                    string s = PinYin.getSpell(item.ToString());
                    if (item.ToString().Contains(this.cmbComp.Text.ToUpper()) || s == this.cmbComp.Text.ToUpper())
                    {
                        listNew.Add(item.ToString());
                    }
                }
                this.cmbComp.Items.AddRange(listNew.ToArray());
                this.cmbComp.SelectionStart = this.cmbComp.Text.Length;
                Cursor = Cursors.Default;
                this.cmbComp.DroppedDown = true;
                this.cmbComp.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.ToString());
            }
        }
        List<string> listNew2 = new List<string>();
        private void cmbDefct_TextUpdate(object sender, EventArgs e)
        {
            try
            {
                this.cmbDefct.Items.Clear();
                listNew2.Clear();
                foreach (var item in myLst2)
                {
                    string s = PinYin.getSpell(item.ToString());
                    if (item.ToString().Contains(this.cmbDefct.Text.ToUpper()) || s == this.cmbDefct.Text.ToUpper())
                    {
                        listNew2.Add(item.ToString());
                    }
                }
                this.cmbDefct.Items.AddRange(listNew2.ToArray());
                this.cmbDefct.SelectionStart = this.cmbDefct.Text.Length;
                Cursor = Cursors.Default;
                this.cmbDefct.DroppedDown = true;
                this.cmbDefct.SelectedIndex = -1;
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.ToString());
            }
        }

        private void cmbComp_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                foreach (DataGridViewRow row in dgvCompName.Rows)
                {
                    string compname = row.Cells[0].Value.ToString();
                    if (compname == this.cmbComp.Text)
                    {
                        int rowindex = row.Index;
                        dgvCompName.FirstDisplayedScrollingRowIndex = rowindex;
                        dgvCompName.ClearSelection();
                        row.Selected = true;
                        this.cmbComp.Text = "";
                        if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                        {
                            this.BtnaddFailure.Enabled = true;
                        }
                    }
                }
            }
        }
        private void cmbDefct_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyData == Keys.Enter)
            {
                foreach (DataGridViewRow row in dgvDefect.Rows)
                {
                    string defect = row.Cells[0].Value.ToString();
                    if (defect == this.cmbDefct.Text)
                    {
                        int rowindex = row.Index;
                        dgvDefect.FirstDisplayedScrollingRowIndex = rowindex;
                        dgvDefect.ClearSelection();
                        row.Selected = true;
                        this.cmbDefct.Text = "";
                        if (config.IsNeedCompColumn == "Y")
                        {
                            if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                            {
                                this.BtnaddFailure.Enabled = true;
                            }
                            else
                            {
                                this.BtnaddFailure.Enabled = false;
                            }
                        }
                        else
                        {
                            if (dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                            {
                                this.BtnaddFailure.Enabled = true;
                            }
                            else
                            {
                                this.BtnaddFailure.Enabled = false;
                            }
                        }
                    }
                }
            }
        }
        private void cmbComp_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvCompName.Rows)
            {
                string compname = row.Cells[0].Value.ToString();
                if (compname == this.cmbComp.Text)
                {
                    int rowindex = row.Index;
                    dgvCompName.FirstDisplayedScrollingRowIndex = rowindex;
                    dgvCompName.ClearSelection();
                    row.Selected = true;
                    this.cmbComp.SelectedIndex = -1;
                    if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                    {
                        this.BtnaddFailure.Enabled = true;
                    }
                }
            }
        }
        private void cmbDefct_SelectedIndexChanged(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dgvDefect.Rows)
            {
                string defect = row.Cells[0].Value.ToString();
                if (defect == this.cmbDefct.Text)
                {
                    int rowindex = row.Index;
                    dgvDefect.FirstDisplayedScrollingRowIndex = rowindex;
                    dgvDefect.ClearSelection();
                    row.Selected = true;
                    this.cmbDefct.SelectedIndex = -1;
                    if (config.IsNeedCompColumn == "Y")
                    {
                        if (dgvCompName.SelectedRows.Count > 0 && dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                        {
                            this.BtnaddFailure.Enabled = true;
                        }
                        else
                        {
                            this.BtnaddFailure.Enabled = false;
                        }
                    }
                    else
                    {
                        if (dgvDefect.SelectedRows.Count > 0 && dgvSN.SelectedRows.Count > 0)
                        {
                            this.BtnaddFailure.Enabled = true;
                        }
                        else
                        {
                            this.BtnaddFailure.Enabled = false;
                        }
                    }
                }
            }
        }
        #endregion
        private string ConvertProcessLayerToString(string str)
        {
            string iValue = "";
            switch (str)
            {
                case "T":
                    iValue = "0";
                    break;
                case "B":
                    iValue = "1";
                    break;
                default:
                    iValue = "2";
                    break;
            }
            return iValue;
        }
        private string ConvertProcessLayerToString2(string str)
        {
            string iValue = "";
            switch (str)
            {
                case "0":
                    iValue = "正面";
                    break;
                case "1":
                    iValue = "反面";
                    break;
                default:
                    iValue = "不分面";
                    break;
            }
            return iValue;
        }

        #region Panel Position and Direction graphics
        PosGraghicsForm frm = null;
        private void btnShowPCB_Click(object sender, EventArgs e)
        {
            if (frm != null && frm.pictureBox1.Image != null)
            {
                frm.Hide();
            }
            if (txbCDAPartNumber.Text == "")
            {
                errorHandler(2, message("No work order"), "");
                return;
            }
            string processlayer = "";
            if (config.IsSelectWO == "Y" || this.txbCDAMONumber.Text != "")
            {
                processlayer = initModel.currentSettings == null ? "-1" : initModel.currentSettings.processLayer.ToString();
            }
            else
            {
                processlayer = iProcessLayer.ToString();
            }
            frm = new PosGraghicsForm(this, sessionContext, initModel, txbCDAPartNumber.Text, processlayer);
            frm.Show();
            if (frm.pictureBox1.Image == null)
                frm.Hide();
        }
        #endregion

        #region checklist from OA
        bool Supervisor = false;
        bool IPQC = true;
        private void InitTaskData_SOCKET(string djclass)
        {
            try
            {
                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, message("no active wo"), "");
                    return;
                }
                this.Invoke(new MethodInvoker(delegate
                {
                    try
                    {
                        this.dgvCheckListTable.Rows.Clear();

                        Supervisor = false;
                        IPQC = true;
                        GetWorkPlanData handle = new GetWorkPlanData(sessionContext, initModel, this);
                        int firstSN = int.Parse(this.lblPass.Text) + int.Parse(this.lblFail.Text) + int.Parse(this.lblScrap.Text);
                        if (firstSN == 0)
                        {
                            djclass = djclass + ";首末件点检";
                        }

                        string workstep_text = handle.GetWorkStepInfobyWorkPlan(this.txbCDAMONumber.Text, initModel.currentSettings.processLayer);
                        if (workstep_text != "")
                        {
                            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
                            string[] processCode = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "TTE_PROCESS_CODE");
                            if (processCode != null && processCode.Length > 0)
                            {
                                string process = processCode[1];
                                string sedmessage = "{getCheckListItem;" + PartNumber + ";" + process + ";[" + workstep_text + "];[" + djclass + "];" + "}";
                                string returnMsg = checklist_cSocket.SendData(sedmessage);

                                if (returnMsg != "" && returnMsg != null)
                                {
                                    string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                    string status = values[1];
                                    if (status == "0")//“0” , or “-1” (error)  
                                    {
                                        int seq = 1;
                                        string itemregular = @"\{[^\{\}]+\}"; //@"\[[^\[\]]+\]";
                                        MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                        if (match.Count <= 0)
                                        {
                                            errorHandler(2, message("No checklist data"), "");
                                            return;
                                        }
                                        else
                                            SetTipMessage(MessageType.OK, "");
                                        for (int i = 0; i < match.Count; i++)
                                        {
                                            string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                            //string[] datas = data.Split(';');
                                            string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                            string sourceclass = datas[4];//数据来源
                                            string formno = datas[0];//对应单号
                                            string itemno = datas[1];//机种品号
                                            string itemnname = datas[2];//机种品名
                                            string sbno = datas[5];//设备编号
                                            string sbname = datas[6];//设备名称
                                            string gcno = datas[7];//过程编号
                                            string gcname = datas[8];//过程名称
                                            string lbclass = datas[9];//类别
                                            string djxmname = datas[10];//点检项目
                                            string specvalue = datas[11];//规格值
                                            string djkind = datas[12];//点检类型
                                            string maxvalue = datas[14];//上限值
                                            string minvalue = datas[13];//下限值
                                            string djclase = datas[15];//点检类别
                                            string djversion = datas[3];//版本
                                            string dataclass = datas[16];//状态

                                            object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                            this.dgvCheckListTable.Rows.Add(objValues);
                                            seq++;
                                            SetCheckListInputStatusTable();

                                            if (djkind == "判断值")
                                            {
                                                string[] strInputValues = new string[] { "Y", "N" };
                                                DataTable dtInput = new DataTable();
                                                dtInput.Columns.Add("name");
                                                dtInput.Columns.Add("value");
                                                DataRow rowEmpty = dtInput.NewRow();
                                                rowEmpty["name"] = "";
                                                rowEmpty["value"] = "";
                                                dtInput.Rows.Add(rowEmpty);
                                                foreach (var strValues in strInputValues)
                                                {
                                                    DataRow row = dtInput.NewRow();
                                                    row["name"] = strValues;
                                                    row["value"] = strValues;
                                                    dtInput.Rows.Add(row);
                                                }

                                                DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                ComboBoxCell.DataSource = dtInput;
                                                ComboBoxCell.DisplayMember = "Name";
                                                ComboBoxCell.ValueMember = "Value";
                                                dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                            }

                                            this.dgvCheckListTable.ClearSelection();
                                        }
                                    }
                                    else
                                    {
                                        string errormsg = values[1];
                                        errorHandler(2, errormsg, "");
                                    }
                                }
                                else
                                {
                                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                                    returnMsg = checklist_cSocket.SendData(sedmessage);

                                    if (returnMsg != "" && returnMsg != null)
                                    {
                                        string[] values = returnMsg.TrimEnd(';').Replace("{", "").Replace("}", "").Replace("#", "").Split(new string[] { ";" }, StringSplitOptions.None);
                                        string status = values[1];
                                        if (status == "0")//“0” , or “-1” (error)  
                                        {
                                            int seq = 1;
                                            string itemregular = @"\{[^\{\}]+\}";
                                            MatchCollection match = Regex.Matches(returnMsg.TrimStart('{').Substring(0, returnMsg.Length - 2), itemregular);
                                            if (match.Count <= 0)
                                            {
                                                errorHandler(2, message("No checklist data"), "");
                                                return;
                                            }
                                            else
                                                SetTipMessage(MessageType.OK, "");
                                            for (int i = 0; i < match.Count; i++)
                                            {
                                                string data = match[i].ToString().TrimStart('{').TrimEnd('}');
                                                //string[] datas = data.Split(';');
                                                string[] datas = Regex.Split(data, "#!#", RegexOptions.IgnoreCase);
                                                string sourceclass = datas[4];//数据来源
                                                string formno = datas[0];//对应单号
                                                string itemno = datas[1];//机种品号
                                                string itemnname = datas[2];//机种品名
                                                string sbno = datas[5];//设备编号
                                                string sbname = datas[6];//设备名称
                                                string gcno = datas[7];//过程编号
                                                string gcname = datas[8];//过程名称
                                                string lbclass = datas[9];//类别
                                                string djxmname = datas[10];//点检项目
                                                string specvalue = datas[11];//规格值
                                                string djkind = datas[12];//点检类型
                                                string maxvalue = datas[14];//上限值
                                                string minvalue = datas[13];//下限值
                                                string djclase = datas[15];//点检类别
                                                string djversion = datas[3];//版本
                                                string dataclass = datas[16];//状态

                                                object[] objValues = new object[] { seq, djclase, djxmname, gcname, specvalue, "", "", "", djkind, gcno, maxvalue, minvalue, lbclass, sourceclass, formno, itemno, itemnname, sbno, sbname, djversion, dataclass, "" };
                                                this.dgvCheckListTable.Rows.Add(objValues);
                                                seq++;
                                                SetCheckListInputStatusTable();

                                                if (djkind == "判断值")
                                                {
                                                    string[] strInputValues = new string[] { "Y", "N" };
                                                    DataTable dtInput = new DataTable();
                                                    dtInput.Columns.Add("name");
                                                    dtInput.Columns.Add("value");
                                                    DataRow rowEmpty = dtInput.NewRow();
                                                    rowEmpty["name"] = "";
                                                    rowEmpty["value"] = "";
                                                    dtInput.Rows.Add(rowEmpty);
                                                    foreach (var strValues in strInputValues)
                                                    {
                                                        DataRow row = dtInput.NewRow();
                                                        row["name"] = strValues;
                                                        row["value"] = strValues;
                                                        dtInput.Rows.Add(row);
                                                    }

                                                    DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                                    ComboBoxCell.DataSource = dtInput;
                                                    ComboBoxCell.DisplayMember = "Name";
                                                    ComboBoxCell.ValueMember = "Value";
                                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                                }

                                                this.dgvCheckListTable.ClearSelection();
                                            }
                                        }
                                        else
                                        {
                                            string errormsg = values[1];
                                            errorHandler(2, errormsg, "");
                                        }
                                    }
                                }
                            }
                            else
                            {
                                errorHandler(2, message("no TTE_PROCESS_CODE"), "");//20161213 edit by qy
                                return;
                            }
                        }
                        else
                        {
                            errorHandler(2, message("no workstep text"), "");
                            return;
                        }
                    }
                    catch (Exception ex)
                    {
                        LogHelper.Error(ex.Message, ex);
                    }
                }));

            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private void SetCheckListInputStatusTable()
        {
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjkind"].Value.ToString() == "判断值")
                {
                    row.Cells["tabResult1"].ReadOnly = true;
                }
                else if (row.Cells["tabdjkind"].Value.ToString() == "输入值" || row.Cells["tabdjkind"].Value.ToString() == "范围值")
                {
                    row.Cells["tabResult2"].ReadOnly = true;
                }
                row.Cells["tabNo"].ReadOnly = true;
                row.Cells["tabStatus"].ReadOnly = true;
            }
        }

        private void btnSupervisor_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQC_Click(object sender, EventArgs e)
        {
            if (gridCheckList.RowCount <= 0)
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }

        public void SupervisorConfirm(string user)//班长确认
        {
            DialogResult dr = MessageBox.Show(message("produtc or not"), message("Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                Supervisor = true;
                errorHandler(0, message("supervisor confirm OK"), "");
            }
            else
            {
                Supervisor = false;
                errorHandler(2, message("supervisor confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;1;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }

        }

        public void IPQCConfirm(string user)//IPQC巡检
        {
            DialogResult dr = MessageBox.Show(message("IPQC produtc or not"), message("Warning"), MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
            if (dr == DialogResult.Yes)
            {
                IPQC = true;
                errorHandler(0, message("IPQC confirm OK"), "");
            }
            else
            {
                IPQC = false;
                errorHandler(2, message("IPQC confirm NG"), "");
            }
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                SaveCheckList();
                string result = "N";
                if (Supervisor)
                    result = "Y";
                string endsendmessage = "{updateCheckListResult;2;" + user + ";" + result + ";" + sequece + "}";
                checklist_cSocket.SendData(endsendmessage);
            }
        }

        private void btnAddCheckListTable_Click(object sender, EventArgs e)
        {
            dgvCheckListTable.Rows.Add(new object[] { this.dgvCheckListTable.Rows.Count + 1, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" });

            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult1"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabNo"].ReadOnly = true;
            dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabStatus"].ReadOnly = true;
            dgvCheckListTable.ClearSelection();
        }
        string sequece = "";
        private void dgvCheckListTable_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //20161208 edit by qy
            try
            {
                if (e.RowIndex == -1)
                    return;
                if (this.dgvCheckListTable.Columns[e.ColumnIndex].Name == "tabResult1" && this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() != "")
                {
                    if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "范围值")
                    {
                        //verify the input value
                        string strRegex = @"^\d{0,9}\.\d{0,9}|-\d{0,9}\.\d{0,9}";//@"^(\d{0,9}.\d{0,9})～(\d{0,9}.\d{0,9}).*$";"^(\-|\+?\d{0,9}.\d{0,9})～(\-|\+?\d{0,9}.\d{0,9})$"
                        string strResult1 = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString();
                        string strStandard = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString().Replace("（", "").Replace("）", "");
                        string strMax = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabmaxvalue"].Value.ToString();
                        string strMin = this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabminvalue"].Value.ToString();
                        Match match1 = Regex.Match(strMax, strRegex);
                        Match match2 = Regex.Match(strMin, strRegex);
                        if (match1.Success && match2.Success)
                        {
                            //if (match.Groups.Count > 2)
                            //{
                            //double iMin = Convert.ToDouble(match.Groups[1].Value);
                            //double iMax = Convert.ToDouble(match.Groups[2].Value);
                            double iMin = Convert.ToDouble(match2.ToString());
                            double iMax = Convert.ToDouble(match1.ToString());
                            double iResult = Convert.ToDouble(strResult1);
                            if (iResult >= iMin && iResult <= iMax)
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                            }
                            else
                            {
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                                this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                            }
                            //}
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                    else if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabdjkind"].Value.ToString() == "输入值")
                    {
                        if (this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabResult1"].Value.ToString() == this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabspecname"].Value.ToString())
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "OK";
                        }
                        else
                        {
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                            this.dgvCheckListTable.Rows[e.RowIndex].Cells["tabStatus"].Value = "NG";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
            }

        }

        private void dgvCheckListTable_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridView dgv = sender as DataGridView;
            if (dgv.CurrentCell.GetType().Name == "DataGridViewComboBoxCell" && dgv.CurrentCell.RowIndex != -1)
            {
                iRowIndex = dgv.CurrentCell.RowIndex;
                (e.Control as ComboBox).SelectedIndexChanged += new EventHandler(ComboBoxTable_SelectedIndexChanged);
            }
        }

        public void ComboBoxTable_SelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.Leave += new EventHandler(comboxtable_Leave);
            try
            {
                if (combox.SelectedItem != null && combox.Text != "")
                {
                    if (OKlist.Contains(combox.Text))
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "OK";
                    }
                    else
                    {
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.Red;
                        this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "NG";
                    }
                }
                else
                {
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Style.BackColor = Color.White;
                    this.dgvCheckListTable.Rows[iRowIndex].Cells["tabStatus"].Value = "";
                }
                Thread.Sleep(100);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public void comboxtable_Leave(object sender, EventArgs e)
        {
            ComboBox combox = sender as ComboBox;
            combox.SelectedIndexChanged -= new EventHandler(ComboBoxTable_SelectedIndexChanged);
        }

        int iIndexCheckListTable = -1;
        private void dgvCheckListTable_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                if (this.dgvCheckListTable.Rows.Count == 0)
                    return;
                ((DataGridView)sender).CurrentRow.Selected = true;
                iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
                this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;

                if (iIndexCheckListTable == -1)
                    this.dgvCheckListTable.ContextMenuStrip = null;

            }
        }
        private void dgvCheckListTable_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            //if (e.Button == MouseButtons.Right)
            //{
            //    if (this.dgvCheckListTable.Rows.Count == 0)
            //        return;

            //    if (e.RowIndex == -1)
            //    {
            //        this.dgvCheckListTable.ContextMenuStrip = null;
            //        return;
            //    }

            //    iIndexCheckListTable = ((DataGridView)sender).CurrentRow.Index;
            //    this.dgvCheckListTable.ContextMenuStrip = contextMenuStrip2;
            //    ((DataGridView)sender).CurrentRow.Selected = true;
            //}
        }

        private void SaveCheckList()
        {
            try
            {
                string path = @"CheckList.txt";
                string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                StringBuilder sb = new StringBuilder();
                sb.AppendLine(datetime);
                sb.AppendLine(txbCDAMONumber.Text + ";" + initModel.currentSettings.processLayer);
                sb.AppendLine(Supervisor.ToString());
                sb.AppendLine(IPQC.ToString());
                sb.AppendLine(sequece);
                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                {
                    string sourceclass = row.Cells["tabsourceclass"].Value.ToString();//数据来源
                    string formno = row.Cells["tabformno"].Value.ToString();//对应单号
                    string itemno = row.Cells["tabitemno"].Value.ToString();//机种品号
                    string itemnname = row.Cells["tabitemname"].Value.ToString();//机种品名
                    string sbno = row.Cells["tabsbno"].Value.ToString();//设备编号
                    string sbname = row.Cells["tabsnname"].Value.ToString();//设备名称
                    string gcno = row.Cells["tabgcno"].Value.ToString();//过程编号
                    string gcname = row.Cells["tabgcname"].Value.ToString();//过程名称
                    string lbclass = row.Cells["tablbclass"].Value.ToString();//类别
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();//点检项目
                    string specvalue = row.Cells["tabspecname"].Value.ToString();//规格值
                    string result1 = row.Cells["tabResult1"].Value.ToString();
                    string result2 = row.Cells["tabResult2"].Value == null ? "" : row.Cells["tabResult2"].Value.ToString();// row.Cells["tabResult2"].Value.ToString();
                    string status = row.Cells["tabstatus"].Value.ToString();//结果
                    string djkind = row.Cells["tabdjkind"].Value.ToString();//点检类型
                    string maxvalue = row.Cells["tabmaxvalue"].Value.ToString();//上限值
                    string minvalue = row.Cells["tabminvalue"].Value.ToString();//下限值
                    string djclase = row.Cells["tabdjclass"].Value.ToString();//点检类别
                    string djversion = row.Cells["tabdjversion"].Value.ToString();//版本
                    string dataclass = row.Cells["tabdataclass"].Value.ToString();//状态

                    //string cell13 = row.Cells[13].Value == null ? "" : row.Cells[13].Value.ToString();
                    string linedata = sourceclass + "￥" + formno + "￥" + itemno + "￥" + itemnname + "￥" + sbno + "￥" + sbname + "￥" + gcno + "￥" + gcname + "￥" + lbclass + "￥" + djxmname + "￥" + specvalue + "￥" + result1 + "￥" + result2 + "￥" + status + "￥" + djkind + "￥" + maxvalue + "￥" + minvalue + "￥" + djclase + "￥" + djversion + "￥" + dataclass;
                    //string linedata = row.Cells[1].Value.ToString() + ";" + row.Cells[2].Value.ToString() + ";" + row.Cells[3].Value.ToString() + ";" + row.Cells[4].Value.ToString() + ";" + row.Cells[5].Value.ToString() + ";" + row.Cells[6].Value.ToString() + ";" + row.Cells[7].Value.ToString() + ";" + row.Cells[8].Value.ToString() + ";" + row.Cells[9].Value.ToString() + ";" + row.Cells[10].Value.ToString() + ";" + row.Cells[11].Value.ToString() + ";" + row.Cells[12].Value.ToString() + ";" + cell13 + ";" + row.Cells[14].Value.ToString() + ";" + row.Cells[15].Value.ToString() + ";" + row.Cells[16].Value.ToString() + ";" + row.Cells[17].Value.ToString() + ";" + row.Cells[18].Value.ToString() + ";" + row.Cells[19].Value.ToString() + ";" + row.Cells[20].Value.ToString() + ";" + djkind;
                    sb.AppendLine(linedata);
                }

                FileStream fs = new FileStream(path, FileMode.Create);
                byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
                fs.Seek(0, SeekOrigin.Begin);
                fs.Write(bt, 0, bt.Length);
                fs.Flush();
                fs.Close();
                LogHelper.Info("Save checklist file success.");
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
            }
        }

        private bool ReadCheckListFile()
        {
            try
            {
                string path = @"CheckList.txt";
                if (File.Exists(path))
                {
                    string[] linelist = File.ReadAllLines(path);
                    string datetimespan = linelist[0];
                    string workorder = linelist[1];
                    Supervisor = Convert.ToBoolean(linelist[2]);
                    IPQC = Convert.ToBoolean(linelist[3]);
                    sequece = linelist[4];
                    TimeSpan span = DateTime.Now - Convert.ToDateTime(datetimespan);

                    if (span.TotalMinutes > Convert.ToInt32(config.RESTORE_TIME))//判断是否大于10分钟，大于10分钟则不自动点检
                    {
                        return false;
                    }
                    else
                    {
                        string[] workorders = workorder.Split(';');
                        if (workorders.Length > 1)
                        {
                            if (workorders[0] == this.txbCDAMONumber.Text)//判断工单是否有变化，无变化则自动点检
                            {
                                //if (workorders[1] == initModel.currentSettings.processLayer.ToString())//判断面次是否有变化
                                //{
                                #region setup checklist
                                int seq = 1;
                                if (linelist.Count() <= 6)
                                    return false;
                                this.dgvCheckListTable.Rows.Clear();
                                for (int i = 5; i < linelist.Count(); i++)
                                {
                                    string line = linelist[i];
                                    if (string.IsNullOrEmpty(line.Trim()))
                                        continue;

                                    string[] datas = line.Split('￥');
                                    object[] objValues = new object[] { seq, datas[17], datas[9], datas[7], datas[10], datas[11], "", datas[13], datas[14], datas[6], datas[15], datas[16], datas[8], datas[0], datas[1], datas[2], datas[3], datas[4], datas[5], datas[18], datas[19], "" };
                                    this.dgvCheckListTable.Rows.Add(objValues);
                                    seq++;
                                    SetCheckListInputStatusTable();
                                    if (datas[14] == "判断值")
                                    {
                                        string[] strInputValues = new string[] { "Y", "N" };
                                        DataTable dtInput = new DataTable();
                                        dtInput.Columns.Add("name");
                                        dtInput.Columns.Add("value");
                                        DataRow rowEmpty = dtInput.NewRow();
                                        rowEmpty["name"] = "";
                                        rowEmpty["value"] = "";
                                        dtInput.Rows.Add(rowEmpty);
                                        foreach (var strValues in strInputValues)
                                        {
                                            DataRow row = dtInput.NewRow();
                                            row["name"] = strValues;
                                            row["value"] = strValues;
                                            dtInput.Rows.Add(row);
                                        }

                                        DataGridViewComboBoxCell ComboBoxCell = new DataGridViewComboBoxCell();
                                        ComboBoxCell.DataSource = dtInput;
                                        ComboBoxCell.DisplayMember = "Name";
                                        ComboBoxCell.ValueMember = "Value";
                                        dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"] = ComboBoxCell;
                                    }
                                    dgvCheckListTable.Rows[this.dgvCheckListTable.Rows.Count - 1].Cells["tabResult2"].Value = datas[12];
                                    this.dgvCheckListTable.ClearSelection();

                                }
                                foreach (DataGridViewRow row in dgvCheckListTable.Rows)
                                {
                                    if (row.Cells["tabStatus"].Value.ToString() == "OK")
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.FromArgb(0, 192, 0);
                                    }
                                    else if ((row.Cells["tabStatus"].Value.ToString() == "NG"))
                                    {
                                        row.Cells["tabStatus"].Style.BackColor = Color.Red;
                                    }

                                }
                                return true;
                                #endregion
                                //}
                                //else
                                //{
                                //    return false;
                                //}
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
                else
                {
                    return false;
                }
            }
            catch (Exception ex)
            {
                LogHelper.Error(ex.Message, ex);
                return false;
            }
        }
        string strShift = "";
        private void WriteIntoShift2()
        {
            string datetime = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            strShift = datetime;
            string path = @"CheckListShiftTemp.txt";
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(datetime + ";" + this.txbCDAMONumber.Text);
            FileStream fs = new FileStream(path, FileMode.OpenOrCreate);
            byte[] bt = Encoding.UTF8.GetBytes(sb.ToString());
            fs.Seek(0, SeekOrigin.Begin);
            fs.Write(bt, 0, bt.Length);
            fs.Flush();
            fs.Close();
        }

        //检查有没有到换班时间，如果到换班时间
        string strShiftChecklist = "";
        private bool CheckShiftChange2()
        {
            try
            {
                bool isValid = false;
                if (strShiftChecklist == "")
                    return false;

                string[] shifchangetimes = config.SHIFT_CHANGE_TIME.Split(';');
                List<string> shiftList = new List<string>();
                string nowDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                for (int i = 0; i < shifchangetimes.Length; i++)
                {

                    shiftList.Add(DateTime.Now.ToString("yyyy/MM/dd ") + shifchangetimes[i].Substring(0, 2) + ":" + shifchangetimes[i].Substring(2, 2));

                }

                shiftList.Sort();

                for (int j = shiftList.Count - 1; j < shiftList.Count; j--)
                {
                    if (j == -1)
                        break;
                    LogHelper.Debug("shift time: " + shiftList[j]);
                    string shitftime = shiftList[j];

                    if (Convert.ToDateTime(nowDate) > Convert.ToDateTime(shiftList[j])) //当前时间与设定的时间做比较，如果到换班时间则比较上次点检的时间
                    {
                        if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        {
                            isValid = true;
                        }
                        break;
                    }
                    else
                    {
                        if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        {
                            string covert_datetime = nowDate;
                            if (j == shiftList.Count - 1)
                            {
                                covert_datetime = shiftList[j - 1];
                            }
                            else if (j == 0)
                            {
                                covert_datetime = shiftList[j];
                            }
                            if (Convert.ToDateTime(strShiftChecklist) < Convert.ToDateTime(nowDate) && Convert.ToDateTime(nowDate) < Convert.ToDateTime(covert_datetime))
                            {
                                isValid = true;
                            }
                            break;
                        }

                        //if (Convert.ToDateTime(strShiftChecklist).ToString("yyyy/MM/dd") != Convert.ToDateTime(nowDate).ToString("yyyy/MM/dd"))//add by qy
                        //{
                        //    shitftime = Convert.ToDateTime(shitftime).AddDays(-1).ToString("yyyy/MM/dd HH:mm:ss");

                        //    if (Convert.ToDateTime(strShiftChecklist) > Convert.ToDateTime(shitftime))
                        //    {
                        //        isValid = true;
                        //    }
                        //    break;
                        //}
                    }
                }

                return isValid;
            }
            catch (Exception ex)
            {
                LogHelper.Debug(ex.Message, ex);
                return false;
            }

        }

        private void InitShift2(string wo)
        {
            string path = @"CheckListShiftTemp.txt";
            if (File.Exists(path))
            {
                string[] content = File.ReadAllLines(path);

                foreach (var item in content)
                {
                    if (item != "")
                    {
                        string[] items = item.Split(';');
                        //if (items[1] == wo)
                        //{
                        strShiftChecklist = items[0];
                        break;
                        //}
                    }
                }
            }
        }
        DateTime next_checklist_time = DateTime.Now;
        string checklist_freq_time = "";
        private void InitWorkOrderType()
        {

            Dictionary<string, string> dicfreq = new Dictionary<string, string>();
            string CHECKLIST_FREQ = config.CHECKLIST_FREQ;
            string[] freqs = CHECKLIST_FREQ.Split(';');
            foreach (var item in freqs)
            {
                string[] items = item.Split(',');
                string key = items[0];
                if (key == "")
                    key = "OTHERS";
                dicfreq[key] = items[1];
            }

            GetAttributeValue getAttriHandler = new GetAttributeValue(sessionContext, initModel, this);
            string[] valuesAttri = getAttriHandler.GetAttributeValueForAll(1, this.txbCDAMONumber.Text, "-1", "WORKORDER_TYPE");
            if (valuesAttri != null && valuesAttri.Length > 0)
            {
                string value = valuesAttri[1];

                if (CHECKLIST_FREQ.Contains(value))
                {
                    checklist_freq_time = dicfreq[value];
                }
                else
                {
                    checklist_freq_time = dicfreq["OTHERS"];
                }
            }
            else
            {
                checklist_freq_time = dicfreq["OTHERS"];
            }
            if (strShiftChecklist != "")
            {
                next_checklist_time = Convert.ToDateTime(strShiftChecklist).AddMinutes(double.Parse(checklist_freq_time) * 60);
            }
            else
            {
                next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
            }

        }

        private void InitProductionChecklist()
        {
            if (DateTime.Now > next_checklist_time)
            {
                if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
                {
                    InitTaskData_SOCKET("过程点检");
                    isStartLineCheck = false;
                    next_checklist_time = DateTime.Now.AddMinutes(double.Parse(checklist_freq_time) * 60);
                }
            }
        }

        public void GetRestoreTimerStart()
        {

            if (RestoreMaterialTimer != null && RestoreMaterialTimer.Enabled)
                return;
            RestoreMaterialTimer = new System.Timers.Timer();
            // 循环间隔时间(1分钟)
            RestoreMaterialTimer.Interval = Convert.ToInt32(config.RESTORE_TREAD_TIMER) * 1000;
            // 允许Timer执行
            RestoreMaterialTimer.Enabled = true;
            // 定义回调
            RestoreMaterialTimer.Elapsed += new ElapsedEventHandler(RestoreMaterialTimer_Elapsed);
            // 定义多次循环
            RestoreMaterialTimer.AutoReset = true;
        }

        private void RestoreMaterialTimer_Elapsed(object sender, ElapsedEventArgs e)
        {
            SaveCheckList();
            InitProductionChecklist();
            InitShiftCheckList();
        }

        bool IsGetShiftCheckList = false;
        private void InitShiftCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //InitShift2(txbCDAMONumber.Text);
                if (!CheckShiftChange2())
                {
                    if (this.dgvCheckListTable.Rows.Count <= 0 || (this.dgvCheckListTable.Rows.Count > 0 && !isStartLineCheck))//!IsShiftCheck()
                    {
                        InitTaskData_SOCKET("开线点检;设备点检");
                        isStartLineCheck = true;
                    }
                }
            }
        }
        private bool IsShiftCheck()//true 表示已经带出开线点检的内容了
        {
            bool isValid = false;
            foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
            {
                if (row.Cells["tabdjclass"].Value.ToString() == "开线点检")
                {
                    isValid = true;
                    break;
                }
            }
            return isValid;
        }

        public void OpenScanPort()
        {
            initModel.scannerHandler = new ScannerHeandler(initModel, this);
            initModel.scannerHandler.handler().DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
            initModel.scannerHandler.handler().Open();
        }

        private void btnSupervisorTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(4, this, "");
            LogForm.ShowDialog();
        }

        private void btnIPQCTable_Click(object sender, EventArgs e)
        {
            if (sequece == "")
            {
                return;
            }
            if (config.LogInType == "COM" && initModel.scannerHandler.handler().IsOpen)
                initModel.scannerHandler.handler().Close();

            LoginForm LogForm = new LoginForm(5, this, "");
            LogForm.ShowDialog();
        }

        private void btnConfirmTable_Click(object sender, EventArgs e)
        {
            try
            {

                string PartNumber = this.txbCDAPartNumber.Text;
                if (PartNumber == "")
                {
                    errorHandler(2, message("no active wo"), "");
                    return;
                }
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value == null || row.Cells["tabStatus"].Value.ToString() == "")
                    {
                        errorHandler(2, message("Verify_CheckList"), "");
                        return;
                    }
                }

                string headmessage = "{appendCheckListResult;" + PartNumber;
                string sedmessage = "";
                string date = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    string gdcode = this.txbCDAMONumber.Text;
                    string itemno = PartNumber;
                    string itemname = initModel.currentSettings.partdesc;
                    string gczcode = initModel.configHandler.StationNumber;
                    string gczname = "";
                    string lineclass = "";
                    string lbclass = row.Cells["tablbclass"].Value.ToString();
                    string djxmname = row.Cells["tabdjxmname"].Value.ToString();
                    string specvalue = "";
                    if (row.Cells["tabResult1"].Value.ToString() != "")
                        specvalue = row.Cells["tabResult1"].Value.ToString();
                    else
                        specvalue = row.Cells["tabResult2"].Value.ToString();
                    string djkind = row.Cells["tabdjkind"].Value.ToString();
                    string maxvalues = row.Cells["tabmaxvalue"].Value.ToString();
                    string minvalues = row.Cells["tabminvalue"].Value.ToString();
                    string djclass = row.Cells["tabdjclass"].Value.ToString();
                    string djversion = row.Cells["tabdjversion"].Value.ToString();
                    string djuser = lblUser.Text;
                    string djremark = "";
                    string djdate = date;
                    string jcuser = lblUser.Text;
                    string qruser = "";
                    string pguser = "";

                    string msgrow = "{" + gdcode + "#!#" + itemno + "#!#" + itemname + "#!#" + gczcode + "#!#" + gczname + "#!#" + lineclass + "#!#" + lbclass + "#!#" + djxmname + "#!#" + specvalue + "#!#" + djkind + "#!#" + maxvalues + "#!#" + minvalues + "#!#" + djclass + "#!#" + djversion + "#!#" + djuser + "#!#" + djremark + "#!#" + djdate + "#!#" + jcuser + "#!#" + qruser + "#!#" + pguser + "}";
                    if (sedmessage == "")
                        sedmessage = msgrow;
                    else
                        sedmessage = sedmessage + ";" + msgrow;
                }
                string endsendmessage = headmessage + ";" + sedmessage + "}";
                string returnMsg = checklist_cSocket.SendData(endsendmessage);
                if (returnMsg != null && returnMsg != "")
                {
                    returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                    string[] Msgs = returnMsg.Split(';');
                    if (Msgs[1] == "0")
                    {
                        if (Supervisor_OPTION == "1")
                        {
                            Supervisor = true;
                            errorHandler(0, message("Send_CheckList_Success"), "");
                        }
                        else
                        {
                            errorHandler(0, message("Send_CheckList_Success,please supervisor confirm"), "");
                        }

                        sequece = Msgs[3];
                        SaveCheckList();
                        WriteIntoShift2();
                        InitShift2(txbCDAMONumber.Text);
                    }
                    else
                    {
                        errorHandler(2, message("Send_CheckList_fail"), "");
                    }
                }
                else
                {
                    isOK = checklist_cSocket.connect(config.CHECKLIST_IPAddress, config.CHECKLIST_Port);
                    returnMsg = checklist_cSocket.SendData(endsendmessage);
                    if (returnMsg != null && returnMsg != "")
                    {
                        returnMsg = returnMsg.TrimStart('{').TrimEnd('}');
                        string[] Msgs = returnMsg.Split(';');
                        if (Msgs[1] == "0")
                        {
                            if (Supervisor_OPTION == "1")
                            {
                                Supervisor = true;
                                errorHandler(0, message("Send_CheckList_Success"), "");
                            }
                            else
                            {
                                errorHandler(0, message("Send_CheckList_Success,please supervisor confirm"), "");
                            }

                            sequece = Msgs[3];
                            SaveCheckList();
                            WriteIntoShift2();
                            InitShift2(txbCDAMONumber.Text);
                        }
                        else
                        {
                            errorHandler(2, message("Send_CheckList_fail"), "");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //20161208 edit by qy
                LogHelper.Error(ex.Message, ex);
            }
        }

        private bool VerifyCheckList()
        {
            if (config.CHECKLIST_SOURCE.ToUpper() == "TABLE")
            {
                //if (!CheckShiftChange2())
                //{
                //    if (this.dgvCheckListTable.Rows.Count <= 0 && dgvCheckListTable.Rows[0].Cells["tabdjclass"].Value.ToString() != "开线点检")
                //    {
                //        InitTaskData_SOCKET("开线点检");
                //    }

                //}
                foreach (DataGridViewRow row in this.dgvCheckListTable.Rows)
                {
                    if (row.Cells["tabStatus"].Value.ToString() != "OK")
                    {
                        errorHandler(2, message("Verify_CheckList"), "");
                        return false;
                    }
                }
                if (this.dgvCheckListTable.Rows.Count > 0)
                {
                    if (!Supervisor)
                    {
                        errorHandler(2, message("Superivisor_check_fail"), "");
                        return false;
                    }
                    if (!IPQC)
                    {

                        errorHandler(2, message("IPQC_check_fail"), "");
                        return false;
                    }
                }

                return true;
            }
            else
            {
                foreach (DataGridViewRow row in gridCheckList.Rows)
                {
                    if (row.Cells["clStatus"].Value.ToString() != "OK")
                    {

                        errorHandler(2, message("Verify_CheckList"), "");
                        return false;
                    }
                }
                //if (this.gridCheckList.Rows.Count > 0)
                //{
                //    if (!Supervisor)
                //    {

                //        errorHandler(2, message("Superivisor_check_fail"), "");
                //        return false;
                //    }
                //    if (!IPQC)
                //    {

                //        errorHandler(2, message("IPQC_check_fail"), "");
                //        return false;
                //    }
                //}
                return true;
            }
        }

        #endregion

        #region 多线程
        public ConcurrentQueue<QueueEntity> SEQueue = new ConcurrentQueue<QueueEntity>();
        public Thread process;//创建并控制线程
        public class QueueEntity
        {
            public List<string> rackSNInfoExt = new List<string>();
        }
        private void ProcessRecipeCommand()
        {
            Thread.CurrentThread.IsBackground = true; //后台线程
            try
            {
                while (true)
                {
                    if (!SEQueue.IsEmpty)
                    {
                        QueueEntity seEntity = null;
                        bool isHas = SEQueue.TryDequeue(out seEntity);
                        if (isHas)
                        {
                            //content
                            AppendRackSN(seEntity.rackSNInfoExt);

                        }
                    }
                    Thread.Sleep(1000);
                }

            }
            catch (Exception ex)
            {
                LogHelper.Error("ProcessSocketCommand" + ex);
            }
        }

        private void AppendRackSN(List<string> sns)
        {
            AppendAttribute appendAttri = new AppendAttribute(sessionContext1, initModel, this);
            for (int s = 0; s < sns.Count / 3; s++)
            {
                string item = sns[s * 3 + 1];
                int state = Convert.ToInt32(sns[s * 3 + 2]);
                if (state != 2 && state != -1)
                    appendAttri.AppendAttributeForAll(0, item, "-1", "RACK_SN", this.lblTextRackNo.Text);
                //else
                //    scrapcount++;
            }
        }
        #endregion

        public bool isFormOutPoump = false;

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            this.txbCDADataInput.Text = "";
            LogoutForm frmOut = new LogoutForm(UserName, this, initModel,sessionContext);
            DialogResult dr = frmOut.ShowDialog();

            if (dr == DialogResult.OK)
            {
                UserName = frmOut.UserName;
                lblUser.Text = UserName;
                lblLoginTime.Text = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                sessionContext = frmOut.sessionContext;
                if (config.LogInType == "COM")
                {
                    SerialPort serialPort = new SerialPort();
                    serialPort.PortName = config.SerialPort;
                    serialPort.BaudRate = int.Parse(config.BaudRate);
                    serialPort.Parity = (Parity)int.Parse(config.Parity);
                    serialPort.StopBits = (StopBits)1;
                    serialPort.Handshake = Handshake.None;
                    serialPort.DataBits = int.Parse(config.DataBits);
                    serialPort.NewLine = "\r";
                    serialPort.DataReceived += new SerialDataReceivedEventHandler(DataRecivedHeandler);
                    serialPort.Open();
                    initModel.scannerHandler.SetSerialPortData(serialPort);
                }
            }
        }
    }
}
