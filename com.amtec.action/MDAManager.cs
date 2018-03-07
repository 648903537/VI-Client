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
    public class MDAManager
    {
        private static IMSApiDotNet imsapi = IMSApiDotNet.loadLibrary();
        private IMSApiSessionContextStruct sessionContext;
        private InitModel init;
        private MainView view;

        public MDAManager(IMSApiSessionContextStruct sessionContext, InitModel init, MainView view)
        {
            this.sessionContext = sessionContext;
            this.init = init;
            this.view = view;
        }

        public List<DocumentEntity> GetDocumentDataByStation()
        {
            List<DocumentEntity> entityList = new List<DocumentEntity>();
            KeyValue[] attributeFilters = new KeyValue[] { new KeyValue("STATION_NUMBER", init.configHandler.StationNumber) };
            KeyValue[] dataTypeFilters = new KeyValue[] { new KeyValue("MDA_ACTIVE", "1"), new KeyValue("MDA_DATA_TYPE", "3") };
            string[] mdaResultKeys = new string[] { "MDA_ACTIVE", "MDA_DATA_TYPE", "MDA_DESC", "MDA_DOC_TYPE", "MDA_DOCUMENT_ID", "MDA_FILE_ID", "MDA_FILE_NAME"
                , "MDA_FILE_PATH", "MDA_NAME", "MDA_STATUS", "MDA_URL_NAME", "MDA_VERSION", "MDA_VERSION_DESC", "MDA_VERSION_NAME" };
            string[] mdaResultValues = new string[] { };
            LogHelper.Info("begin api mdaGetDocuments (Station number:" + init.configHandler.StationNumber + ")");
            int errorCode = imsapi.mdaGetDocuments(sessionContext, init.configHandler.StationNumber, attributeFilters, dataTypeFilters, mdaResultKeys, out mdaResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api mdaGetDocuments (result code = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode, "");
            }
            if (errorCode == 0)
            {
                int loop = mdaResultKeys.Length;
                int count = mdaResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    DocumentEntity entity = new DocumentEntity();
                    entity.MDA_ACTIVE = mdaResultValues[i];
                    entity.MDA_DATA_TYPE = mdaResultValues[i + 1];
                    entity.MDA_DESC = mdaResultValues[i + 2];
                    entity.MDA_DOC_TYPE = mdaResultValues[i + 3];
                    entity.MDA_DOCUMENT_ID = mdaResultValues[i + 4];
                    entity.MDA_FILE_ID = mdaResultValues[i + 5];
                    entity.MDA_FILE_NAME = mdaResultValues[i + 6];
                    entity.MDA_FILE_PATH = mdaResultValues[i + 7];
                    entity.MDA_NAME = mdaResultValues[i + 8];
                    entity.MDA_STATUS = mdaResultValues[i + 9];
                    entity.MDA_URL_NAME = mdaResultValues[i + 10];
                    entity.MDA_VERSION = mdaResultValues[i + 11];
                    entity.MDA_VERSION_DESC = mdaResultValues[i + 12];
                    entity.MDA_VERSION_NAME = mdaResultValues[i + 13];
                    entityList.Add(entity);
                }
            }
            return entityList;
        }

        public List<DocumentEntity> GetDocumentDataByPN(string partNumber)
        {
            List<DocumentEntity> entityList = new List<DocumentEntity>();
            KeyValue[] attributeFilters = new KeyValue[] { new KeyValue("PART_NUMBER", partNumber) };
            KeyValue[] dataTypeFilters = new KeyValue[] { new KeyValue("MDA_ACTIVE", "1"), new KeyValue("MDA_DATA_TYPE", "3") };
            string[] mdaResultKeys = new string[] { "MDA_ACTIVE", "MDA_DATA_TYPE", "MDA_DESC", "MDA_DOC_TYPE", "MDA_DOCUMENT_ID", "MDA_FILE_ID", "MDA_FILE_NAME"
                , "MDA_FILE_PATH", "MDA_NAME", "MDA_STATUS", "MDA_URL_NAME", "MDA_VERSION", "MDA_VERSION_DESC", "MDA_VERSION_NAME" };
            string[] mdaResultValues = new string[] { };
            LogHelper.Info("begin api mdaGetDocuments (Part number:" + partNumber + ")");
            int errorCode = imsapi.mdaGetDocuments(sessionContext, init.configHandler.StationNumber, attributeFilters, dataTypeFilters, mdaResultKeys, out mdaResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api mdaGetDocuments (result code = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode, "");
            }
            if (errorCode == 0)
            {
                int loop = mdaResultKeys.Length;
                int count = mdaResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    DocumentEntity entity = new DocumentEntity();
                    entity.MDA_ACTIVE = mdaResultValues[i];
                    entity.MDA_DATA_TYPE = mdaResultValues[i + 1];
                    entity.MDA_DESC = mdaResultValues[i + 2];
                    entity.MDA_DOC_TYPE = mdaResultValues[i + 3];
                    entity.MDA_DOCUMENT_ID = mdaResultValues[i + 4];
                    entity.MDA_FILE_ID = mdaResultValues[i + 5];
                    entity.MDA_FILE_NAME = mdaResultValues[i + 6];
                    entity.MDA_FILE_PATH = mdaResultValues[i + 7];
                    entity.MDA_NAME = mdaResultValues[i + 8];
                    entity.MDA_STATUS = mdaResultValues[i + 9];
                    entity.MDA_URL_NAME = mdaResultValues[i + 10];
                    entity.MDA_VERSION = mdaResultValues[i + 11];
                    entity.MDA_VERSION_DESC = mdaResultValues[i + 12];
                    entity.MDA_VERSION_NAME = mdaResultValues[i + 13];
                    entityList.Add(entity);
                }
            }
            return entityList;
        }

        public List<DocumentEntity> GetDocumentDataByWorkplan(string stationNo, string workorder)
        {
            string[] strValues = GetWorkplanDataForStation(stationNo, workorder);
            if (strValues == null)
                return null;
            List<DocumentEntity> entityList = new List<DocumentEntity>();
            KeyValue[] attributeFilters = new KeyValue[] { new KeyValue("WORKPLAN_NUMBER", strValues[0]), new KeyValue("WORKPLAN_VERS", strValues[1]), new KeyValue("WORKSTEP_NUMBER", strValues[2]) };
            KeyValue[] dataTypeFilters = new KeyValue[] { new KeyValue("MDA_ACTIVE", "1"), new KeyValue("MDA_DATA_TYPE", "3") };
            string[] mdaResultKeys = new string[] { "MDA_ACTIVE", "MDA_DATA_TYPE", "MDA_DESC", "MDA_DOC_TYPE", "MDA_DOCUMENT_ID", "MDA_FILE_ID", "MDA_FILE_NAME"
                , "MDA_FILE_PATH", "MDA_NAME", "MDA_STATUS", "MDA_URL_NAME", "MDA_VERSION", "MDA_VERSION_DESC", "MDA_VERSION_NAME" };
            string[] mdaResultValues = new string[] { };
            LogHelper.Info("begin api mdaGetDocuments (Station number:" + workorder + ")");
            int errorCode = imsapi.mdaGetDocuments(sessionContext, stationNo, attributeFilters, dataTypeFilters, mdaResultKeys, out mdaResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api mdaGetDocuments (errorcode = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode, "");
            }
            if (errorCode == 0)
            {
                int loop = mdaResultKeys.Length;
                int count = mdaResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    DocumentEntity entity = new DocumentEntity();
                    entity.MDA_ACTIVE = mdaResultValues[i];
                    entity.MDA_DATA_TYPE = mdaResultValues[i + 1];
                    entity.MDA_DESC = mdaResultValues[i + 2];
                    entity.MDA_DOC_TYPE = mdaResultValues[i + 3];
                    entity.MDA_DOCUMENT_ID = mdaResultValues[i + 4];
                    entity.MDA_FILE_ID = mdaResultValues[i + 5];
                    entity.MDA_FILE_NAME = mdaResultValues[i + 6];
                    entity.MDA_FILE_PATH = mdaResultValues[i + 7];
                    entity.MDA_NAME = mdaResultValues[i + 8];
                    entity.MDA_STATUS = mdaResultValues[i + 9];
                    entity.MDA_URL_NAME = mdaResultValues[i + 10];
                    entity.MDA_VERSION = mdaResultValues[i + 11];
                    entity.MDA_VERSION_DESC = mdaResultValues[i + 12];
                    entity.MDA_VERSION_NAME = mdaResultValues[i + 13];
                    entityList.Add(entity);
                }
            }
            return entityList;
        }

        private string[] GetWorkplanDataForStation(string stationNo, string workorder)
        {
            List<string> values = new List<string>();
            KeyValue[] workplanFilter = new KeyValue[] { new KeyValue("STATION_BASED_WORKSTEPS", "1"), new KeyValue("WORKORDER_NUMBER", workorder) };
            string[] workplanDataResultKeys = new string[] { "PART_NUMBER", "STATION_NUMBER", "WORKPLAN_VERS", "WORKSTEP_NUMBER" };
            string[] workplanDataResultValues = new string[] { };
            int error = imsapi.mdataGetWorkplanData(sessionContext, stationNo, workplanFilter, workplanDataResultKeys, out workplanDataResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdataGetWorkplanData " + error+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                int loop = workplanDataResultKeys.Length;
                int count = workplanDataResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    if (workplanDataResultValues[i + 1] == stationNo)
                    {
                        values.Add(workplanDataResultValues[i]);
                        values.Add(workplanDataResultValues[i + 2]);
                        values.Add(workplanDataResultValues[i + 3]);
                        break;
                    }
                }

                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdataGetWorkplanData " + error, "");
                return values.ToArray();
            }
        }

        public List<DocumentEntity> GetDocumentDataByAdvice(int adviceID)
        {
            List<DocumentEntity> entityList = new List<DocumentEntity>();
            KeyValue[] attributeFilters = new KeyValue[] { new KeyValue("ADVICE_ID", adviceID.ToString()) };
            KeyValue[] dataTypeFilters = new KeyValue[] { new KeyValue("MDA_ACTIVE", "1"), new KeyValue("MDA_DATA_TYPE", "3") };
            string[] mdaResultKeys = new string[] { "MDA_ACTIVE", "MDA_DATA_TYPE", "MDA_DESC", "MDA_DOC_TYPE", "MDA_DOCUMENT_ID", "MDA_FILE_ID", "MDA_FILE_NAME"
                , "MDA_FILE_PATH", "MDA_NAME", "MDA_STATUS", "MDA_URL_NAME", "MDA_VERSION", "MDA_VERSION_DESC", "MDA_VERSION_NAME" };
            string[] mdaResultValues = new string[] { };
            LogHelper.Info("begin api mdaGetDocuments (Advice ID:" + adviceID + ")");
            int errorCode = imsapi.mdaGetDocuments(sessionContext, init.configHandler.StationNumber, attributeFilters, dataTypeFilters, mdaResultKeys, out mdaResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(errorCode, init, sessionContext);
            LogHelper.Info("end api mdaGetDocuments (result code = " + errorCode + ")");
            if (errorCode != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocuments " + errorCode, "");
            }
            if (errorCode == 0)
            {
                int loop = mdaResultKeys.Length;
                int count = mdaResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    DocumentEntity entity = new DocumentEntity();
                    entity.MDA_ACTIVE = mdaResultValues[i];
                    entity.MDA_DATA_TYPE = mdaResultValues[i + 1];
                    entity.MDA_DESC = mdaResultValues[i + 2];
                    entity.MDA_DOC_TYPE = mdaResultValues[i + 3];
                    entity.MDA_DOCUMENT_ID = mdaResultValues[i + 4];
                    entity.MDA_FILE_ID = mdaResultValues[i + 5];
                    entity.MDA_FILE_NAME = mdaResultValues[i + 6];
                    entity.MDA_FILE_PATH = mdaResultValues[i + 7];
                    entity.MDA_NAME = mdaResultValues[i + 8];
                    entity.MDA_STATUS = mdaResultValues[i + 9];
                    entity.MDA_URL_NAME = mdaResultValues[i + 10];
                    entity.MDA_VERSION = mdaResultValues[i + 11];
                    entity.MDA_VERSION_DESC = mdaResultValues[i + 12];
                    entity.MDA_VERSION_NAME = mdaResultValues[i + 13];
                    entityList.Add(entity);
                }
            }
            return entityList;
        }

        public byte[] GetDocumnetContentByID(long documentID)
        {
            byte[] content = new byte[] { };
            int error = imsapi.mdaGetDocumentContent(sessionContext, init.configHandler.StationNumber, documentID, out content);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocumentContent " + error+"("+errorMsg+")", "");
                return null;
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetDocumentContent " + error, "");
            }
            return content;
        }

        public Advice[] GetAdviceByStationAndWO(string workorder)
        {
            Advice[] adviceArray = null;
            KeyValue[] adviceFilters = new KeyValue[] { new KeyValue("WORKORDER_NUMBER", workorder) };
            int error = imsapi.adviceGetAdvice(sessionContext, init.configHandler.StationNumber, false, false, false, adviceFilters, out adviceArray);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " adviceGetAdvice " + error+"("+errorMsg+")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " adviceGetAdvice " + error, "");
            }
            return adviceArray;
        }

        public Advice[] GetAdviceByStationAndPN(string partNumber)
        {
            Advice[] adviceArray = null;
            KeyValue[] adviceFilters = new KeyValue[] { new KeyValue("PART_NUMBER", partNumber) };
            int error = imsapi.adviceGetAdvice(sessionContext, init.configHandler.StationNumber, false, false, false, adviceFilters, out adviceArray);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            if (error != 0)
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " adviceGetAdvice " + error+"("+errorMsg+")", "");
            }
            else
            {
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " adviceGetAdvice " + error, "");
            }
            return adviceArray;
        }

        public int UploadDocumentForPart(string partNumber, string fileName, string filePath, string mdaName, byte[] content)
        {
            KeyValue[] attributeValues = new KeyValue[] { new KeyValue("PART_NUMBER", partNumber) };
            KeyValue[] mdaValues = new KeyValue[]{new KeyValue("MDA_ACTIVE","1"),new KeyValue("MDA_DATA_TYPE","3"),new KeyValue("MDA_DESC",mdaName)
            ,new KeyValue("MDA_NAME",mdaName),new KeyValue("MDA_FILE_NAME",fileName),new KeyValue("MDA_STATUS","2"),new KeyValue("MDA_URL_NAME",filePath)
            ,new KeyValue("MDA_VERSION_DESC","-1"),new KeyValue("MDA_VERSION_NAME","-1")};
            int error = imsapi.mdaUploadDocument(sessionContext, init.configHandler.StationNumber, 100, attributeValues, mdaValues, content);
            return error;
        }

        public List<RecipeEntity> GetRecipeDataByPN(string stationNumber, string partNumber, string machineGroupNo)
        {
            List<RecipeEntity> entityList = new List<RecipeEntity>();
            KeyValue[] recipeFilters = new KeyValue[] { new KeyValue("MACHINE_GROUP_NUMBER", ""), new KeyValue("PART_NUMBER", ""), new KeyValue("STATION_NUMBER", "") };
            string[] recipeResultKeys = new string[] { "MEASURE_NAME", "NOMINAL", "REMARK" };
            string[] recipeResultValues = new string[] { };
            int error = imsapi.mdaGetRecipeData(sessionContext, stationNumber, -1, "-1", "-1", "-1", -1, "-1", "-1", 1, recipeFilters, recipeResultKeys, out recipeResultValues);
            string errorMsg = UtilityFunction.GetZHSErrorString(error, init, sessionContext);
            LogHelper.Info("api mdaGetRecipeData (station number =" + stationNumber + ",part number =" + partNumber + ",machine group number =" + machineGroupNo + ") error code =" + error);
            if (error == 0)
            {
                int loop = recipeResultKeys.Length;
                int count = recipeResultValues.Length;
                for (int i = 0; i < count; i += loop)
                {
                    RecipeEntity entity = new RecipeEntity();
                    entity.MEASURE_NAME = recipeResultValues[i];
                    entity.NOMINAL = recipeResultValues[i + 1];
                    entity.REMARK = recipeResultValues[i + 2];
                    entityList.Add(entity);
                }
                view.errorHandler(0, init.lang.ERROR_API_CALL_ERROR + " mdaGetRecipeData " + error, "");
            }
            else
            {
                view.errorHandler(2, init.lang.ERROR_API_CALL_ERROR + " mdaGetRecipeData " + error+"("+errorMsg+")", "");
            }
            return entityList;
        }
    }
}