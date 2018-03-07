using com.amtec.action;
using com.amtec.forms;
using com.amtec.model;
using System;
using System.IO.Ports;

namespace com.amtec.device
{
    public class ScannerHeandler
    {
        private SerialPort serialPort;
        private SerialPort outputSP;
        private InitModel init;
        private MainView view;

        public ScannerHeandler(InitModel init, MainView view)
        {
            this.init = init;
            this.view = view;
            if (init.configHandler.SerialPort != "" && init.configHandler.SerialPort != null)
            {
                serialPort = new SerialPort();
                serialPort.PortName = init.configHandler.SerialPort;
                serialPort.BaudRate = int.Parse(init.configHandler.BaudRate);
                serialPort.Parity = (Parity)int.Parse(init.configHandler.Parity);
                serialPort.StopBits = (StopBits)1;
                serialPort.Handshake = Handshake.None;
                serialPort.DataBits = int.Parse(init.configHandler.DataBits);
                serialPort.NewLine = "\r";
            }
            if (init.configHandler.DataOutputInterface == "COM")
            {
                outputSP = new SerialPort();
                outputSP.PortName = init.configHandler.OutSerialPort;
                outputSP.BaudRate = int.Parse(init.configHandler.OutBaudRate);
                outputSP.Parity = (Parity)int.Parse(init.configHandler.OutParity);
                outputSP.StopBits = (StopBits)1;
                outputSP.Handshake = Handshake.None;
                outputSP.DataBits = int.Parse(init.configHandler.OutDataBits);
                outputSP.NewLine = "\r";
            }
        }

        public SerialPort handler()
        {
            return serialPort;
        }
        public void SetSerialPortData(SerialPort setSP)
        {
            serialPort = setSP;
        }
        public SerialPort OutputCOM()
        {
            return outputSP;
        }

        public void endCommand()
        {
            //char[] charArray;
            //String text = init.configHandler.EndCommand;
            //String tmpString = text.Trim();
            //tmpString = tmpString.Replace("ESC", "*");
            //charArray = tmpString.ToCharArray();

            //for (int i = 0; i < charArray.Length; i++)
            //{
            //    if (charArray[i].Equals((char)42))
            //    {
            //        charArray[i] = (char)27;
            //    }
            //}

            //try
            //{
            //    serialPort.Write(charArray, 0, charArray.Length);
            //    LogHelper.Info("Send end command:" + text);
            //    view.errorHandler(0, "Send end command", "Send end command");
            //}
            //catch
            //{
            //    LogHelper.Info("Send end command error");
            //    view.errorHandler(0, "Send end command error", "Send end command error");
            //}
        }

        public void sendHigh()
        {

            //char[] charArray;
            //String text = init.configHandler.High;
            //String tmpString = text.Trim();
            //tmpString = tmpString.Replace("ESC", "*");
            //charArray = tmpString.ToCharArray();

            //for (int i = 0; i < charArray.Length; i++)
            //{
            //    if (charArray[i].Equals((char)42))
            //    {
            //        charArray[i] = (char)27;
            //    }
            //}

            //try
            //{
            //    serialPort.Write(charArray, 0, charArray.Length);
            //    view.errorHandler(0, "SEND HIGH", "SEND HIGH");
            //    LogHelper.Info("SEND HIGH:" + text);
            //}
            //catch
            //{
            //    view.errorHandler(2, "SEND HIGH ERROR", "SEND HIGH ERROR");
            //    LogHelper.Info("SEND HIGH ERROR");
            //}
        }

        public void sendLow()
        {
            //char[] charArray;
            //String text = init.configHandler.Low;
            //String tmpString = text.Trim();
            //tmpString = tmpString.Replace("ESC", "*");
            //charArray = tmpString.ToCharArray();

            //for (int i = 0; i < charArray.Length; i++)
            //{
            //    if (charArray[i].Equals((char)42))
            //    {
            //        charArray[i] = (char)27;
            //    }
            //}

            //try
            //{
            //    serialPort.Write(charArray, 0, charArray.Length);
            //    view.errorHandler(0, "SEND LOW", "SEND LOW");
            //    LogHelper.Info("SEND LOW:" + text);
            //}
            //catch
            //{
            //    view.errorHandler(2, "SEND LOW ERROR", "SEND LOW ERROR");
            //    LogHelper.Info("SEND LOW ERROR");
            //}
        }
    }
}
