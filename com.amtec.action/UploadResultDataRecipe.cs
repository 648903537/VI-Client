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
    public class UploadResultDataRecipe
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private MainView view;
        private InitModel init;

        public UploadResultDataRecipe(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int trUploadResultDataAndRecipe(float cycletime, string serialNum, string serialNumPos, string serialnumState, List<string> listItem, int multiboard, out string[] SNStateResultValues)
        {
            SNStateResultValues = new string[] { };
            var uploadValues = new string[] { };

            string[] UploadKeyN = { "ErrorCode", "LowerLimit", "MeasureFailCode", "MeasureName", "MeasureValue", "Unit", "Nominal", "Remark", "Tolerance", "UpperLimit" };

            uploadValues = listItem.ToArray();

            //recipeVersionMode  要从-1 改为0 之前是有问题的  第二次更改从0改为1 感觉没什么区别.
            //recipeVersionMode =-1 或者为空 recipeVersionId=-1 measurename 名称长度不能超过80个字符 超过 报-10  duplicateserialnumber =0 表示只上传大板中某一个位置的SN =1/-1 上传时都会复制给大板中其他位置的小板
            int error = imsapi.trUploadResultDataAndRecipe(sessionContext, init.configHandler.StationNumber, init.currentSettings.processLayer, -1, serialNum,// 不知道为什么processLayer 只能=1 以前的都是2的
                serialNumPos, Convert.ToInt32(serialnumState), multiboard, -1, cycletime, 1, UploadKeyN, uploadValues, out SNStateResultValues);//snState 默认设置为1   
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error == 0)
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " trUploadResultDataAndRecipe " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " trUploadResultDataAndRecipe " + error+"("+errorMsg+")", "");
            }

            return error;
        }

    }
}
