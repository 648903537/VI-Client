using com.amtec.forms;
using com.amtec.model;
using com.itac.mes.imsapi.client.dotnet;
using com.itac.mes.imsapi.domain.container;

namespace com.amtec.action
{
    public class ProcessMaterialBinData
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public ProcessMaterialBinData(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public int UpdateMaterialBinBooking(string materialBin, double strQty)
        {
            string[] materialBinBookingUploadKeys = new string[] { "ERROR_CODE", "MATERIAL_BIN_NUMBER", "QUANTITY", "TRANSACTION_CODE" };
            string[] materialBinBookingUploadValues = new string[] { "0", materialBin, strQty.ToString(), "0" };
            string[] materialBinBookingResultValues = new string[] { };
            LogHelper.Info("begin api mlUploadMaterialBinBooking (material bin number =" + materialBin + ",quantity=" + strQty + ")");
            int error = imsapi.mlUploadMaterialBinBooking(sessionContext, init.configHandler.StationNumber, materialBinBookingUploadKeys, materialBinBookingUploadValues, out materialBinBookingResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("end api mlUploadMaterialBinBooking (result code = " + error + ")");
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mlUploadMaterialBinBooking " + error+"("+errorMsg+")", "");
                return error;
            }
            view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mlUploadMaterialBinBooking " + error, "");
            return error;
        }
    }
}
