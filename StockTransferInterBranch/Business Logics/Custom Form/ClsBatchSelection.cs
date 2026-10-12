using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using SAPbobsCOM;
using SAPbouiCOM;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using Microsoft.ApplicationBlocks.Data;
using System.Xml;
using System.IO;
using System.Globalization;

namespace StockTransferInterBranch
{
    class ClsBatchSelection : Base
    {


        #region variable declaration
        private SAPbouiCOM.DBDataSource _MasterDBDataSource;
        SAPbouiCOM.Matrix _Matrix;
        SAPbobsCOM.Recordset _RS1 = null;
        SAPbouiCOM.Form transferForm = null;
        public static int EntLine = 0;
        public double _quantity;
        public string _INTEntry;
        public decimal _ReqQty;
        public bool sucess;

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsBatchSelection()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //

        }
        ~ClsBatchSelection()
        {
        }

        #endregion

        #region ITEM EVENT
        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION

            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            double totalQty = 0.0;

                            SAPbouiCOM.Form form = Utilities.Application.SBO_Application.Forms.Item(pVal.FormUID);
                            SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)form.Items.Item("mtx3").Specific;

                            for (int i = 1; i <= matrix.RowCount; i++)
                            {
                                var cell = (SAPbouiCOM.EditText)matrix.Columns.Item("colReqQty").Cells.Item(i).Specific;
                                string qtyStr = cell.Value.Trim();

                                if (double.TryParse(qtyStr, out double qty))
                                {
                                    totalQty += qty;
                                }
                            }

                            if (totalQty != _quantity)
                            {
                                Utilities.Message("Required Quantity must be matched with transfer quantity.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                if (transferForm != null)
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)transferForm.Items.Item("mtx1").Specific;
                                    SAPbouiCOM.CommonSetting commonSetting = oMatrix.CommonSetting;

                                    for (int col = 1; col < oMatrix.Columns.Count; col++)
                                        commonSetting.SetCellEditable(EntLine, col, true);
                                }

                                BubbleEvent = false;
                                return;
                            }

                        }
                        break;
                }
            }
            #endregion

            #region After ACTION

            else if (pVal.Before_Action == false)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && pVal.ActionSuccess && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            _Form.Close();
                            for (int i = 0; i < Utilities.Application.SBO_Application.Forms.Count; i++)
                            {
                                SAPbouiCOM.Form frm = Utilities.Application.SBO_Application.Forms.Item(i);
                                if (frm.TypeEx == "IBT_FORM")
                                {
                                    transferForm = frm;
                                    break;
                                }
                            }

                            if (transferForm != null)
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)transferForm.Items.Item("mtx1").Specific;
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchE").Cells.Item(EntLine).Specific).Value = "Yes";
                                SAPbouiCOM.CommonSetting commonSetting = oMatrix.CommonSetting;
                                for (int col = 1; col < oMatrix.Columns.Count; col++)
                                {
                                    commonSetting.SetCellEditable(EntLine, col, false);
                                }
                                transferForm.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            }
                            else
                            {
                                Utilities.Application.SBO_Application.StatusBar.SetText("form not found.",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            }


                        }

                        if (pVal.ItemUID == "btnPost")
                        {

                            bool sucess = SaveBatchesToUDO(_INTEntry, EntLine, _quantity);
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }
                            if (sucess)
                            {
                                _Form.Close();


                                SAPbouiCOM.Form transferForm = null;
                                for (int i = 0; i < Utilities.Application.SBO_Application.Forms.Count; i++)
                                {
                                    SAPbouiCOM.Form frm = Utilities.Application.SBO_Application.Forms.Item(i);
                                    if (frm.TypeEx == "IBT_FORM")
                                    {
                                        transferForm = frm;
                                        break;
                                    }
                                }
                                if (transferForm != null)
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)transferForm.Items.Item("mtx1").Specific;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchE").Cells.Item(EntLine).Specific).Value = "Yes";
                                    SAPbouiCOM.CommonSetting commonSetting = oMatrix.CommonSetting;
                                    // Lock the row (full batch posted)
                                    for (int col = 1; col < oMatrix.Columns.Count; col++)
                                    {
                                        commonSetting.SetCellEditable(EntLine, col, false);
                                    }
                                    transferForm.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }
                                else
                                {
                                    Utilities.Application.SBO_Application.StatusBar.SetText(
                                        "Form not found.",
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Error
                                    );
                                }

                            }

                            else
                            {
                                for (int i = 0; i < Utilities.Application.SBO_Application.Forms.Count; i++)
                                {
                                    SAPbouiCOM.Form frm = Utilities.Application.SBO_Application.Forms.Item(i);
                                    if (frm.TypeEx == "IBT_FORM")
                                    {
                                        transferForm = frm;
                                        break;
                                    }
                                }
                                if (transferForm !=null)
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)transferForm.Items.Item("mtx1").Specific;
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchE").Cells.Item(EntLine).Specific).Value = " ";
                                    SAPbouiCOM.CommonSetting commonSetting = oMatrix.CommonSetting;
                                    commonSetting.SetCellEditable(EntLine, 3, true);
                                }


                            }

                        }
                        break;

                        #endregion
                }
            }

            #endregion
        }

        #endregion

        #region Form Default
        internal void FormDefault(string itemCode, string whsCode, int line, string IntReqEntry, double quantity, string colid)
        {
            _Form = Utilities.Application.SBO_Application.Forms.Item(FormUID);
            _Form.Freeze(true);
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx3").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BATCHS");
            EntLine = line;
            _quantity = quantity;
            _INTEntry = IntReqEntry;
            if (colid == "colBtchE")
            {
                _Form.Items.Item("btnPost").Visible = false;
            }
            else if (colid == "rowHead")
            {
                _Form.Items.Item("1").Visible = false;
            }
            string query = $@"
            SELECT 
                T0.""ItemCode"", 
                T2.""ItemName"", 
                T0.""WhsCode"", 
                T1.""BatchNum"", 
                T1.""Quantity"" ""BatchQuantity"", 
                T3.""BinCode"", 
                T4.""ExpDate"", 
                T4.""MnfDate"",
                '0' AS ""ReqQty""
            FROM OITW T0
            LEFT JOIN OBIN T3 ON T0.""DftBinAbs"" = T3.""AbsEntry""
            INNER JOIN OIBT T1 ON T0.""WhsCode"" = T1.""WhsCode"" AND T0.""ItemCode"" = T1.""ItemCode""
            INNER JOIN OITM T2 ON T0.""ItemCode"" = T2.""ItemCode""
            INNER JOIN OBTN T4 ON T1.""ItemCode"" = T4.""ItemCode"" AND T1.""BatchNum"" = T4.""DistNumber""
            WHERE T0.""OnHand"" > 0 AND T2.""ItemCode"" = '{itemCode}' AND T1.""Quantity"" > 0 AND T0.""WhsCode"" = '{whsCode}'
            GROUP BY T0.""ItemCode"", T2.""ItemName"", T0.""WhsCode"", T1.""BatchNum"", T1.""Quantity"", T0.""OnHand"", T3.""BinCode"", 
            T4.""ExpDate"", T4.""MnfDate""
                
                union all
                
             SELECT 
                T0.""U_ItemCode"", 
                T0.""U_ItemName"", 
                T0.""U_WhsCode"", 
                T0.""U_BatchNum"", 
                T0.""U_BatchQty"" AS ""BatchQuantity"", 
                T0.""U_BinCode"", 
                T0.""U_ExpDate"", 
                T0.""U_MnfDate"", T0.""U_ReqQty""
            FROM ""@BATCHS"" T0
          
            WHERE T0.""U_ItemCode"" = '{itemCode}' AND ""U_IntDocEntry"" = '{IntReqEntry}' AND ""U_LineNum"" = '{line}' 
             AND (IFNULL(T0.""U_BatchPost"",'') = '' OR T0.""U_BatchPost"" <> 'Y')
            GROUP BY T0.""U_ItemCode"", T0.""U_ItemName"",  T0.""U_WhsCode"", T0.""U_BatchNum"", T0.""U_BatchQty"", T0.""U_BinCode"",  
                T0.""U_ExpDate"", T0.""U_MnfDate"", T0.""U_ReqQty""
            ";

            Utilities.ExecuteSQL(ref _RS1, query);

            int i = 0;
            _MasterDBDataSource.Clear();
            _Matrix.Clear();
            _RS1.MoveFirst();

            int maxCode = 0;
            int.TryParse(Utilities.getMaxColumnValueNum("@BATCHS", "Code").TrimStart('0'), out maxCode);

            // Step 1: Grouping data
            var groupedData = new Dictionary<string, dynamic>();

            while (!_RS1.EoF)
            {
                itemCode = _RS1.Fields.Item("ItemCode").Value.ToString();
                string batchNum = _RS1.Fields.Item("BatchNum").Value.ToString();
                string groupKey = itemCode + "|" + batchNum;

                decimal reqQty = Convert.ToDecimal(_RS1.Fields.Item("ReqQty").Value);
                _ReqQty = reqQty;
                if (!groupedData.ContainsKey(groupKey))
                {
                    groupedData[groupKey] = new
                    {
                        ItemCode = itemCode,
                        //ItemName = _RS1.Fields.Item("ItemName").Value.ToString(),
                        WhsCode = _RS1.Fields.Item("WhsCode").Value.ToString(),
                        BatchNum = batchNum,
                        BatchQty = _RS1.Fields.Item("BatchQuantity").Value.ToString(),
                        ReqQty = reqQty
                    };
                }
                else
                {
                    // Accumulate ReqQty
                    var existing = groupedData[groupKey];
                    groupedData[groupKey] = new
                    {
                        existing.ItemCode,
                        //existing.ItemName,
                        existing.WhsCode,
                        existing.BatchNum,
                        existing.BatchQty,
                        ReqQty = existing.ReqQty + reqQty
                    };
                }

                _RS1.MoveNext();
            }

            foreach (var entry in groupedData.Values)
            {
                int currentIndex = _MasterDBDataSource.Size;
                _MasterDBDataSource.InsertRecord(currentIndex);
                int nextCode = ++maxCode;
                _MasterDBDataSource.SetValue("Code", currentIndex, nextCode.ToString());
                _MasterDBDataSource.SetValue("Name", currentIndex, nextCode.ToString());
                _MasterDBDataSource.SetValue("U_ItemCode", currentIndex, entry.ItemCode);
                //_MasterDBDataSource.SetValue("U_ItemName", currentIndex, entry.ItemName);
                _MasterDBDataSource.SetValue("U_WhsCode", currentIndex, entry.WhsCode);
                _MasterDBDataSource.SetValue("U_BatchNum", currentIndex, entry.BatchNum);
                _MasterDBDataSource.SetValue("U_BatchQty", currentIndex, entry.BatchQty);
                _MasterDBDataSource.SetValue("U_IntDocEntry", currentIndex, IntReqEntry);
                _MasterDBDataSource.SetValue("U_LineNum", currentIndex, EntLine.ToString());
                _MasterDBDataSource.SetValue("U_ReqQty", currentIndex, entry.ReqQty.ToString());
            }
            _Matrix.LoadFromDataSource();
            if (_ReqQty > 0)
            {
                _Form.Items.Item("1").Visible = false;
                _Form.Items.Item("btnPost").Visible = true;
            }

            _Form.Freeze(false);
        }


        #endregion


        #region SetForm
        internal void SetForm(Form oForm)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region SaveBatchesToUDO
        internal bool SaveBatchesToUDO(string intDocEntry, int lineNum, double expectedLineQty)
        {
            _Form = Utilities.Application.SBO_Application.Forms.Item(FormUID);
            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx3").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@BATCHS");

            double totalSelectedQty = 0;
            List<string> matrixBatchNums = new List<string>();

            // 1. Calculate total quantity
            for (int i = 1; i <= _Matrix.RowCount; i++)
            {
                string qtyStr = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colReqQty").Cells.Item(i).Specific).Value;
                if (double.TryParse(qtyStr, out double qty) && qty > 0)
                {
                    totalSelectedQty += qty;
                }
            }

            // 2. Validate
            if (Math.Round(totalSelectedQty, 3) != Math.Round(expectedLineQty, 3))
            {
                Utilities.Application.SBO_Application.StatusBar.SetText(
                    $"Batch quantity mismatch. Required: {expectedLineQty}, Selected: {totalSelectedQty}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //sucess = false;
                return false;

            }

            // 3. Company service + UDO service
            SAPbobsCOM.CompanyService companyService = Utilities.Application.Company.GetCompanyService();
            SAPbobsCOM.GeneralService generalService =
                (SAPbobsCOM.GeneralService)companyService.GetGeneralService("INTBATCHS");  // UDO code

            // 4. Loop matrix rows
            for (int i = 0; i < _MasterDBDataSource.Size; i++)
            {
                string batchNum = _MasterDBDataSource.GetValue("U_BatchNum", i).Trim();
                string qtyStr = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colReqQty").Cells.Item(i + 1).Specific).Value;

                if (!string.IsNullOrEmpty(batchNum) && double.TryParse(qtyStr, out double qty) && qty > 0)
                {
                    matrixBatchNums.Add(batchNum);

                    // check if record exists
                    string fetchKey = $@"SELECT ""Code"" FROM ""@BATCHS"" 
                                  WHERE ""U_IntDocEntry"" = '{intDocEntry}' 
                                  AND ""U_LineNum"" = '{lineNum}' 
                                  AND ""U_BatchNum"" = '{batchNum}' 
                                  AND (IFNULL(""U_BatchPost"",'') = '' OR ""U_BatchPost"" <> 'Y')";
                    string existingCode = Utilities.ExecuteScalar(fetchKey);

                    if (!string.IsNullOrEmpty(existingCode))
                    {
                        // --- UPDATE ---
                        SAPbobsCOM.GeneralDataParams oParams =
                            (SAPbobsCOM.GeneralDataParams)generalService.GetDataInterface(
                                SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                        oParams.SetProperty("Code", existingCode);

                        SAPbobsCOM.GeneralData existingData = generalService.GetByParams(oParams);
                        existingData.SetProperty("U_ReqQty", qtyStr);

                        generalService.Update(existingData);
                    }
                    else
                    {
                        // --- INSERT ---
                        string code = _MasterDBDataSource.GetValue("Code", i).Trim();
                        string name = _MasterDBDataSource.GetValue("Name", i).Trim();
                        string itemCode = _MasterDBDataSource.GetValue("U_ItemCode", i).Trim();
                        string itemName = _MasterDBDataSource.GetValue("U_ItemName", i).Trim();
                        string whsCode = _MasterDBDataSource.GetValue("U_WhsCode", i).Trim();
                        string batchQty = _MasterDBDataSource.GetValue("U_BatchQty", i).Trim();
                        string totalStock = _MasterDBDataSource.GetValue("U_TotalStock", i).Trim();
                        string expDate = _MasterDBDataSource.GetValue("U_ExpDate", i).Trim();
                        string mnfDate = _MasterDBDataSource.GetValue("U_MnfDate", i).Trim();

                        SAPbobsCOM.GeneralData newData =
                            (SAPbobsCOM.GeneralData)generalService.GetDataInterface(
                                SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);

                        newData.SetProperty("Code", code);
                        newData.SetProperty("Name", name);
                        newData.SetProperty("U_ItemCode", itemCode);
                        newData.SetProperty("U_ItemName", itemName);
                        newData.SetProperty("U_WhsCode", whsCode);
                        newData.SetProperty("U_BatchNum", batchNum);
                        newData.SetProperty("U_BatchQty", batchQty);
                        newData.SetProperty("U_TotalStock", totalStock);
                        newData.SetProperty("U_ExpDate", (expDate));
                        newData.SetProperty("U_MnfDate", (mnfDate));
                        newData.SetProperty("U_IntDocEntry", intDocEntry);
                        newData.SetProperty("U_LineNum", lineNum.ToString());
                        newData.SetProperty("U_ReqQty", qtyStr);

                        generalService.Add(newData);
                    }
                }
            }

            // 5. Delete removed batches
            string fetchExisting = $@"SELECT ""Code"", ""U_BatchNum"" FROM ""@BATCHS"" 
                              WHERE ""U_IntDocEntry"" = '{intDocEntry}' AND ""U_LineNum"" = '{lineNum}' AND (IFNULL(""U_BatchPost"",'') = '' OR ""U_BatchPost"" <> 'Y')";
            var dbRecords = GetBatchCodeNumList(fetchExisting); // return List<(string Code, string BatchNum)>

            foreach (var rec in dbRecords)
            {
                if (!matrixBatchNums.Contains(rec.BatchNum))
                {
                    SAPbobsCOM.GeneralDataParams delParams =
                        (SAPbobsCOM.GeneralDataParams)generalService.GetDataInterface(
                            SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    delParams.SetProperty("Code", rec.Code);

                    generalService.Delete(delParams);
                }
            }

            Utilities.Application.SBO_Application.StatusBar.SetText(
                "Batch data saved/updated successfully",
                SAPbouiCOM.BoMessageTime.bmt_Short,
                SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            return true;
        }
        #endregion

        #region BatchRecord
        private class BatchRecord
        {
            public string Code { get; set; }
            public string BatchNum { get; set; }
        }
        #endregion

        #region GetBatchCodeNumList
        private List<BatchRecord> GetBatchCodeNumList(string sql)
        {
            var list = new List<BatchRecord>();
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            while (!rs.EoF)
            {
                list.Add(new BatchRecord
                {
                    Code = rs.Fields.Item("Code").Value.ToString(),
                    BatchNum = rs.Fields.Item("U_BatchNum").Value.ToString()
                });
                rs.MoveNext();
            }

            return list;
        }
        #endregion

        #region GetPosted Batch Qty
        private double GetPostedBatchQty(string intEntry, int lineNum)
        {
            string query = $@"
        SELECT IFNULL(SUM(CAST(""U_BatchQty"" AS NUMERIC)), 0) AS TotalQty
        FROM ""@BATCHS""
        WHERE ""U_IntDocEntry"" = '{intEntry}'
          AND ""U_LineNum"" = '{lineNum}'";

            SAPbobsCOM.Recordset rs =
                (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(query);

            return Convert.ToDouble(rs.Fields.Item("TotalQty").Value);
        }
        #endregion  





    }
}
