using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;
using SAPbobsCOM;
//using SAPbouiCOM;
using System.Data;
using System.Data.SqlClient;
using System.Data.Sql;
using Microsoft.ApplicationBlocks.Data;
using System.Xml;
using System.IO;
using SAPbouiCOM;

namespace StockTransferInterBranch
{
    class ClsInterBranchTransfer : Base
    {
        #region VARIABLE DECLARATION
        private SAPbobsCOM.Recordset _RS = null, _RSGR = null, _RSGRU = null,_RSSeries = null,_RSGLCode = null,_RSBPL=null;
        SAPbouiCOM.Item _ExistItem, _NewItem;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "", _DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        private bool flag;
        private Hashtable _lines = new Hashtable();
        SAPbobsCOM.Recordset _RS2 = null, _RS1 = null, _RSBatch =null ;
        public string oSql2 = string.Empty;
        public string GRPODelNo = string.Empty, BranchCd = string.Empty;
        SAPbobsCOM.Recordset _Rs1 = null, _RSGRPO = null;
        SAPbouiCOM.Matrix _Matrix;
        public string _ITREntry = string.Empty;
        public string _ItemCode = string.Empty;
        public string  _LineNo = string.Empty;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource, _ChildDBDataSource;
        private SAPbouiCOM.Application _Application;
        private bool errorFlag;
        public string checkDocEntry;
        private int _rowToDelete = -1;
        public string IssueEntry = string.Empty;
        public string ReceiptEntry = string.Empty;
        public string _Error = string.Empty;
        List<LineInfo> lineList = new List<LineInfo>();




        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsInterBranchTransfer()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
           
        }
        ~ClsInterBranchTransfer()
        {
        }

        #endregion

        #region FormDataEvent
        public override void FormData_Event(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        Form.Freeze(true);
                        Disable_Form_Item();
                        Form.Freeze(false);
                        _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        break;


                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }


        }
        #endregion


        #region ITEM EVENT

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {

            #region BEFORE ACTION ACTION

            if (pVal.Before_Action == true)
            {
                switch (pVal.EventType)
                {
                   
                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                       if (pVal.ItemUID == "1"&& _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value == "")
                            {
                                Utilities.Message("Inter Branch Transfer document cannot be created without a corresponding Inter Branch Transfer Request.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }

                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtGoodRN").Specific).Value == "" && ((SAPbouiCOM.EditText)_Form.Items.Item("txtGoodIN").Specific).Value == "")
                            {
   
                                 Utilities.Message("Please wait operation is in progress....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                var company = Utilities.Application.Company;
                       
                                try
                                {
                                    Utilities.Message("Please wait, operation is in progress...", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                                    if (!company.InTransaction)
                                        company.StartTransaction();

                                    errorFlag = false;

                                    // Step 1: Goods Issue
                                    InventoryGoodsIssue();
                                    if (errorFlag)
                                        throw new Exception("Goods Issue failed. Rolling back transaction.");

                                    // Step 2: Goods Receipt
                                    InventoryGoodsReceipt();
                                    if (errorFlag)
                                        throw new Exception("Goods Receipt failed. Rolling back transaction.");

                                    company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                                    // Update batch info if available
                                    foreach (var line in lineList)
                                    {
                                        string updateQry = $@"
                                        UPDATE ""@BATCHS"" 
                                        SET ""U_BatchPost"" = 'Y' 
                                        WHERE ""U_IntDocEntry"" = '{line.DocEntry}' 
                                            AND ""U_ItemCode"" = '{line.ItemCode}' 
                                            AND ""U_LineNum"" = '{line.LineId}'";
                                        var rsUpdate = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rsUpdate.DoQuery(updateQry);
                                    }
                                    // Save Good Issue/Receipt entries
                                    var dbSource = _Form.DataSources.DBDataSources.Item("@IBTFRH");
                                    dbSource.SetValue("U_GoodIssueEntry", 0, IssueEntry.ToString());
                                    dbSource.SetValue("U_GoodReceiptEntry", 0, ReceiptEntry.ToString());

                                    // Assign DocNums
                                    dbSource.SetValue("U_GoodIssueNo", 0, GetDocNum("OIGE", Convert.ToInt32(IssueEntry)));
                                    dbSource.SetValue("U_GoodReceiptNo", 0, GetDocNum("OIGN", Convert.ToInt32(ReceiptEntry)));

                                    // Function to fetch DocNum
                                    string GetDocNum(string table, int docEntry)
                                    {
                                        string query = $"SELECT \"DocNum\" FROM \"{table}\" WHERE \"DocEntry\" = '{docEntry}'";
                                        var rs = (SAPbobsCOM.Recordset)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        rs.DoQuery(query);
                                        return rs.EoF ? "" : rs.Fields.Item("DocNum").Value.ToString();
                                    }

                                    Utilities.Message("InterBranch Transfer created successfully.", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                }
                                catch (Exception ex)
                                {
                                    if (company.InTransaction)
                                        company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                                    Utilities.Message("Error during InterBranch Transfer: " + ex.Message +_Error, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    oBubbleEvent = false;
                                }

                            }
                            else
                            {
                                Utilities.Message("InterBranch Transfer already created for this issue", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false;
                                return;
                            }
                        }


                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE && pVal.BeforeAction)
                        {
                            string goodRN = ((SAPbouiCOM.EditText)_Form.Items.Item("txtGoodRN").Specific).Value;
                            string goodIN = ((SAPbouiCOM.EditText)_Form.Items.Item("txtGoodIN").Specific).Value;

                            if (!string.IsNullOrEmpty(goodRN) && !string.IsNullOrEmpty(goodIN))
                            {
                                Utilities.Message("Cannot change the date of already created document.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                oBubbleEvent = false; 
                                return;
                            }
                        }


                        break;
                  
                        #endregion
                }
            }

            #endregion

            #region After ACTION ACTION

            else if (pVal.Before_Action == false)
            {
                switch (pVal.EventType)
                {
                    #region ITEM PRESS EVENT
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
               
                        break;

                   
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        break;

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                     
                        break;

            
                    #region Double Click
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ColUID == "rowHead")
                        {

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific;
                            string currentbatch = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchE").Cells.Item(pVal.Row).Specific).Value;
                            string batchitem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchSts").Cells.Item(pVal.Row).Specific).Value;
                            
                                string _itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colItmCd").Cells.Item(pVal.Row).Specific).Value;
                                int _Line = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("rowHead").Cells.Item(pVal.Row).Specific).Value);
                                string _whsCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value;
                                string _IntReqEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value;
                                double _Qty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("colQty").Cells.Item(pVal.Row).Specific).Value);


                                this._Object = new ClsBatchSelection();
                                this.OpenChildForm(Constants.Forms.BatchSelection, ref pVal, true);
                                //Utilities.LoadForm(ref this._Object, Constants.Forms.BatchSelection);
                                ((ClsBatchSelection)_Object).FormDefault(_itemCode, _whsCode, _Line, _IntReqEntry, _Qty,pVal.ColUID);
         
                        }

                        break;
                        #endregion
                        #endregion

                    #region et_VALIDATE

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (!pVal.BeforeAction && pVal.ItemUID == "mtx1" && pVal.ColUID == "colBtchE")
                        {
                            

                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific;
                                string currentbatch = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchE").Cells.Item(pVal.Row).Specific).Value;
                                string batchitem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colBtchSts").Cells.Item(pVal.Row).Specific).Value;
                                if (currentbatch != "Yes" || batchitem == "N")
                                    {
                                        string _itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colItmCd").Cells.Item(pVal.Row).Specific).Value;
                                        int _Line = Convert.ToInt32(((SAPbouiCOM.EditText)oMatrix.Columns.Item("rowHead").Cells.Item(pVal.Row).Specific).Value);
                                        string _whsCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value;
                                        string _IntReqEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value;
                                        double _Qty = Convert.ToDouble(((SAPbouiCOM.EditText)oMatrix.Columns.Item("colQty").Cells.Item(pVal.Row).Specific).Value);


                                this._Object = new ClsBatchSelection();
                                        this.OpenChildForm(Constants.Forms.BatchSelection, ref pVal, true);
                                        //Utilities.LoadForm(ref this._Object, Constants.Forms.BatchSelection);
                                        ((ClsBatchSelection)_Object).FormDefault(_itemCode, _whsCode, _Line, _IntReqEntry,_Qty ,pVal.ColUID);
                                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                //batchForm.FormDefault(itemCode, whsCode);
                            }
                            
                        }
                        break;

                        #endregion
                }
            }

            #endregion

        }
        #endregion

        #region Menu_Event
        public override void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
        {
            try
            {

                if (pVal.BeforeAction)
                {

                    switch (pVal.MenuUID)
                    {
                        
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        case Constants.System_Menus.mnu_FIND:
                          
                            break;

                        case Constants.System_Menus.mnu_FIRST:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }

                            break;

                        case Constants.System_Menus.mnu_LAST:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }

                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }
                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            if (_Form.Mode != BoFormMode.fm_OK_MODE)
                            {
                                // Cancel the changes (like hitting "No")
                                _Form.Mode = BoFormMode.fm_OK_MODE; // Reset mode to OK
                            }

                            break;
                          
                    }
                    if (pVal.MenuUID == "1293")
                    {
                        //RefreshRowIndex((SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific);
                        RenumberMatrixAndDataSource(_Matrix, _ChildDBDataSource);
                    }

                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                Form.Freeze(false);
            }
          
        }
        #endregion

        #region Form Default
        internal void FormDefault(string DocEntry)
        {

            _Form.Freeze(true);
            _Form.DataBrowser.BrowseBy = "txtDocNum";
            if (DocEntry != "")
            {
                //_Form.Settings.Enabled = true;
                _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtx1").Specific;
                _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@IBTFRH");
                _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@IBTFRC");
                //string _SQL = "select (select \"BPLName\" from \"OBPL\" where \"BPLId\" = T0.\"U_Branch\")\"U_BranchName\",\"U_Branch\",\"U_DocDate\",\"U_DocNum\",\"U_FrmWhse\",\"U_ToWhse\",\"U_Status\",\"U_CreatedBy\",\"U_Remarks\",\"U_CRemarks\",\"U_FBranchN\", \"U_TBranchN\", \"DocEntry\" from \"@IBTFRRH\" where \"DocEntry\" = '" + DocEntry + "'";
                string _SQL = @"
                SELECT 
                    (SELECT ""BPLName"" FROM ""OBPL"" WHERE ""BPLId"" = T0.""U_Branch"") AS ""U_BranchName"",
                    T0.""U_Branch"",
                    T0.""U_DocDate"",
                    T0.""U_DocNum"",
                    T0.""U_FrmWhse"",
                    T0.""U_ToWhse"",
                    T0.""U_Status"",
                    T0.""U_CreatedBy"",
                    T0.""U_Remarks"",
                    T0.""U_CRemarks"",
                    T0.""U_FBranchN"",
                    T0.""U_TBranchN"",
                    T0.""DocEntry""
                FROM ""@IBTFRRH"" T0
                WHERE T0.""DocEntry"" = '" + DocEntry + "'";

                Utilities.ExecuteSQL(ref _RS, _SQL);

                ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbBranch").Specific).Select(_RS.Fields.Item("U_BranchName").Value.ToString(), SAPbouiCOM.BoSearchKey.psk_ByValue);
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtDate").Specific).Value = Convert.ToDateTime(_RS.Fields.Item("U_DocDate").Value.ToString()).ToString("yyyyMMdd");
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNum").Specific).Value = Utilities.getMaxColumnValueNum("@IBTFRH", "U_DocNum");
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value = _RS.Fields.Item("U_FrmWhse").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value = _RS.Fields.Item("U_ToWhse").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = _RS.Fields.Item("U_Status").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtBy").Specific).Value = _RS.Fields.Item("U_CreatedBy").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value = _RS.Fields.Item("U_FBranchN").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value = _RS.Fields.Item("U_TBranchN").Value.ToString();
                ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value = _RS.Fields.Item("DocEntry").Value.ToString();
                
                string _SQL1 = "SELECT T1.\"U_ItemCode\", T1.\"U_ItemName\", T1.\"U_Qty\", T1.\"U_OpnQty\", T1.\"U_InterComCode\", T1.\"U_Remarks\" AS \"T1_U_Remarks\", T1.\"U_ToWhseStatus\", T1.\"U_LineId\", " +
                   "T0.\"U_Branch\", T0.\"U_DocDate\", T0.\"U_DocNum\", T0.\"U_FrmWhse\", T0.\"U_ToWhse\", T0.\"U_Status\", T0.\"U_CreatedBy\", T0.\"U_Remarks\", T0.\"U_CRemarks\",  T2.\"ManBtchNum\", " +
                   "T3.\"OnHand\" " +
                   "FROM \"@IBTFRRH\" T0 " +
                   "INNER JOIN \"@IBTFRRC\" T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                   "INNER JOIN \"OITM\" T2 ON T2.\"ItemCode\" = T1.\"U_ItemCode\" " +
                   "LEFT JOIN OITW  T3 ON T3.\"ItemCode\" = T1.\"U_ItemCode\" AND T3.\"WhsCode\" = T0.\"U_FrmWhse\" " +
                   "WHERE T0.\"DocEntry\" = '" + DocEntry + "' and T1.\"U_OpnQty\" >0 " +
                   "ORDER BY T1.\"U_LineId\"";


                Utilities.ExecuteSQL(ref _RS1, _SQL1);


                

                _ChildDBDataSource.Clear();    // UDO child table datasource
                _Matrix.Clear();               // Matrix itself
                _RS1.MoveFirst();
                int i = 0;
                while (!_RS1.EoF)
                {
                    _ChildDBDataSource.InsertRecord(i);
                    string seq = (i + 1).ToString();
                    _ChildDBDataSource.SetValue("U_LineId", i, seq);
                    //_ChildDBDataSource.SetValue("U_LineId", i, _RS1.Fields.Item("U_LineId").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemCode", i, _RS1.Fields.Item("U_ItemCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ItemName", i, _RS1.Fields.Item("U_ItemName").Value.ToString());
                    //string _SqlBatch = "SELECT T1.\"U_ItemCode\", T1.\"U_ReqQty\", "+
                    //                "FROM \"@BATCHS\" T0 " +
                    //                "WHERE T0.\"U_IntDocEntry\" ='"+_RS.Fields.Item("DocEntry").Value.ToString()+"' and  T0.\"U_ItemCode \" ='"+ _RS1.Fields.Item("U_ItemCode").Value.ToString()+"'"+
                    //                "ORDER BY T1.\"U_LineId\"";
                    //Utilities.ExecuteSQL(ref _RSBatch, _SqlBatch);


                    _ChildDBDataSource.SetValue("U_Qty", i, _RS1.Fields.Item("U_OpnQty").Value.ToString());
                    _ChildDBDataSource.SetValue("U_InterComCode", i, _RS1.Fields.Item("U_InterComCode").Value.ToString());
                    _ChildDBDataSource.SetValue("U_Remarks", i, _RS1.Fields.Item("T1_U_Remarks").Value.ToString());
                    _ChildDBDataSource.SetValue("U_ToWhseStatus", i, _RS1.Fields.Item("OnHand").Value.ToString());
                    _ChildDBDataSource.SetValue("U_BatchStatus", i, _RS1.Fields.Item("ManBtchNum").Value.ToString());
                    _ChildDBDataSource.SetValue("U_LineNum", i, _RS1.Fields.Item("U_LineId").Value.ToString());
                    i++;
                    _RS1.MoveNext();
                }
                _Matrix.LoadFromDataSource();
            }

            _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

            _Form.EnableMenu("1293", true);
            //EnableFormMenu();
            _Form.Freeze(false);
            _Form.PaneLevel = 1;

           
        }
        #endregion

        #region InventoryGoodsIssue
        private void InventoryGoodsIssue()
        {

            int _lineCount = 0;
            SAPbobsCOM.Documents _GI = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);


            DateTime parsedPostingDate = DateTime.Now;


            string branchName = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value;
            string bpl = "SELECT \"BPLId\" FROM OBPL WHERE  \"BPLName\" = '" + branchName + "'";
            Utilities.ExecuteSQL(ref _RSBPL, bpl);
            int branchCode = Convert.ToInt32(_RSBPL.Fields.Item("BPLId").Value);
            //int branchCode = _RSBPL.Fields.Item("BPLId").Value.ToString();
            _GI.DocDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
            _GI.DocDueDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
            _GI.Comments = "Inter Branch Transfer";
            _GI.JournalMemo = "Inter Branch Transfer";
            _GI.BPL_IDAssignedToInvoice = branchCode;
            string docEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value;
            string seriesQry = "SELECT \"Series\" FROM NNM1 WHERE \"ObjectCode\" = '60' AND \"BPLId\" = '" + branchCode + "' AND \"Indicator\" = " +
                               "(SELECT DISTINCT \"Indicator\" FROM OFPR WHERE '" + parsedPostingDate.ToString("yyyyMMdd") + "' BETWEEN \"F_TaxDate\" AND \"T_TaxDate\")";
            Utilities.ExecuteSQL(ref _RSSeries, seriesQry);

            string _QryGLCode = "select \"U_InterBranchGLCode\" from OWHS where \"WhsCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value + "'";
            Utilities.ExecuteSQL(ref _RSGLCode, _QryGLCode);
            if (_RSSeries.RecordCount > 0)
            {
                _GI.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value);
            }
            
            string _InterBranchReq = "select \"U_DocNum\",\"DocEntry\" from \"@IBTFRRH\" where \"DocEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value + "'";
            SAPbobsCOM.Recordset rsIBT = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rsIBT.DoQuery(_InterBranchReq);
            string TransDocNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNum").Specific).Value.ToString();
            if (!rsIBT.EoF)
            {
                string ibtDocNum = rsIBT.Fields.Item("U_DocNum").Value.ToString();
                string ibtDocEntry = rsIBT.Fields.Item("DocEntry").Value.ToString();

                _GI.UserFields.Fields.Item("U_IBTRNo").Value = ibtDocNum;
                _GI.UserFields.Fields.Item("U_IBTREntry").Value = ibtDocEntry;
                _GI.UserFields.Fields.Item("U_IBTNo").Value = TransDocNum;
            }



            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific;
            int rowCount = oMatrix.RowCount;
            //List<LineInfo> lineList = new List<LineInfo>();
            for (int row = 1; row <= rowCount; row++)
            {
                if (_lineCount > 0)
                    _GI.Lines.Add();

                string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colItmCd").Cells.Item(row).Specific).Value;
                string IntitemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colICom").Cells.Item(row).Specific).Value;
                string qtyStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colQty").Cells.Item(row).Specific).Value;
                string lineId = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("rowHead").Cells.Item(row).Specific).Value;
                double quantity = string.IsNullOrEmpty(qtyStr) ? 0 : Convert.ToDouble(qtyStr);
                _ItemCode = itemCode;
                _LineNo = lineId;
                _GI.Lines.ItemCode = itemCode;
                _GI.Lines.Quantity = quantity;
                _GI.Lines.WarehouseCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value;
                if (_RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString() == "")
                {
                   
                    throw new Exception("Inter Branch GL Code is missing for the selected warehouse.");

                    errorFlag = true;
                    return;
                }
                else
                {
                    _GI.Lines.AccountCode = _RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString();
                }
                
                lineList.Add(new LineInfo
                {
                    ItemCode = itemCode,
                    LineId = lineId,
                    DocEntry = docEntry
                });


                // Get batches from @BATCHS
                string batchQry = $@"
                 SELECT ""U_BatchNum"", ""U_ReqQty"", ""U_ExpDate"" FROM ""@BATCHS"" WHERE ""U_IntDocEntry"" = '{docEntry}'
                 AND ""U_ItemCode"" = '{itemCode}'AND ""U_LineNum"" = '{lineId}'   
                 AND (IFNULL(""U_BatchPost"",'') = '' OR ""U_BatchPost"" <> 'Y')";
                SAPbobsCOM.Recordset rsBatch = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsBatch.DoQuery(batchQry);

                int batchLine = 0;
                while (!rsBatch.EoF)
                {
                    _GI.Lines.BatchNumbers.SetCurrentLine(batchLine);
                    _GI.Lines.BatchNumbers.BatchNumber = rsBatch.Fields.Item("U_BatchNum").Value.ToString();
                    _GI.Lines.BatchNumbers.Quantity = Convert.ToDouble(rsBatch.Fields.Item("U_ReqQty").Value);
                    string expDateStr = rsBatch.Fields.Item("U_ExpDate").Value.ToString();
                    if (!string.IsNullOrEmpty(expDateStr))
                    {
                        DateTime Expdate = DateTime.ParseExact(expDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        _GI.Lines.BatchNumbers.ExpiryDate = Utilities.ToDate(Expdate.ToString("yyyyMMdd"));
                    }
                    _GI.Lines.BatchNumbers.Add();
                    batchLine++;
                    rsBatch.MoveNext();
                }

                _lineCount++;
            }

            int retVal = _GI.Add();
            if (retVal == 0)
            {
                string newDocEntry;
                Utilities.Application.Company.GetNewObjectCode(out newDocEntry);
                IssueEntry = newDocEntry;
                errorFlag = false;
            }
            else
            {
                _Error = Utilities.Application.Company.GetLastErrorDescription();
                errorFlag = true;
            }
        }

        #endregion

        #region InventoryGoodsReceipt
        private void InventoryGoodsReceipt()
        {
            int _lineCount = 0;
            SAPbobsCOM.Documents _GR = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);

            // Get document header data 
            DateTime parsedPostingDate = DateTime.Now;
            //_GRPO.DocDate = 
            string branchName = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value;
            string toWhseName = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value;
            string bpl = "SELECT \"BPLId\" FROM OBPL WHERE  \"BPLName\" = '" + branchName + "'";
            Utilities.ExecuteSQL(ref _RSBPL, bpl);
            //int branchCode =_RSBPL.Fields.Item("BPLId").Value.ToString();
            int branchCode = Convert.ToInt32(_RSBPL.Fields.Item("BPLId").Value);
            _GR.DocDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
            _GR.DocDueDate = Utilities.ToDate(parsedPostingDate.ToString("yyyyMMdd"));
            _GR.Comments = "Inter Branch Transfer";
            _GR.JournalMemo = "Inter Branch Transfer";
            _GR.BPL_IDAssignedToInvoice = branchCode;
            //_GR.Lines.WarehouseCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value;
            // Get the correct Series from DB
            string docEntry = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value;

            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific;

            // Check the first row of colICom to add warehouse condition
            string checkIntItem = "";
            if (oMatrix.RowCount > 0)
            {
                checkIntItem = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colICom").Cells.Item(1).Specific).Value;
            }

            string seriesQry = "SELECT \"Series\" FROM NNM1 WHERE \"ObjectCode\" = '59' AND \"BPLId\" = '" + branchCode+ "' AND \"Indicator\" = " +
                               "(SELECT DISTINCT \"Indicator\" FROM OFPR WHERE '" + parsedPostingDate.ToString("yyyyMMdd") + "' BETWEEN \"F_TaxDate\" AND \"T_TaxDate\")";
            if (string.IsNullOrEmpty(checkIntItem))
            {
                seriesQry += " AND \"U_whse\" = '" + toWhseName + "' ";
            }

            Utilities.ExecuteSQL(ref _RSSeries, seriesQry);


            string _QryGLCode = "select \"U_InterBranchGLCode\" from OWHS where \"WhsCode\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value + "'";
            Utilities.ExecuteSQL(ref _RSGLCode, _QryGLCode);


            string _InterBranchReq = "select \"U_DocNum\",\"DocEntry\" from \"@IBTFRRH\" where \"DocEntry\" = '" + ((SAPbouiCOM.EditText)_Form.Items.Item("txtTREnt").Specific).Value + "'";
            SAPbobsCOM.Recordset rsIBT = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rsIBT.DoQuery(_InterBranchReq);
            string TransDocNum = ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNum").Specific).Value.ToString();
            if (!rsIBT.EoF)
            {
                string ibtDocNum = rsIBT.Fields.Item("U_DocNum").Value.ToString();
                string ibtDocEntry = rsIBT.Fields.Item("DocEntry").Value.ToString();

                _GR.UserFields.Fields.Item("U_IBTRNo").Value = ibtDocNum;
                _GR.UserFields.Fields.Item("U_IBTREntry").Value = ibtDocEntry;
                _GR.UserFields.Fields.Item("U_IBTNo").Value = TransDocNum;
            }

            if (_RSSeries.RecordCount > 0)
            {
                _GR.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value);
            }

            //SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)_Form.Items.Item("mtx1").Specific;
            int rowCount = oMatrix.RowCount;

            for (int row = 1; row <= rowCount; row++)
            {
                if (_lineCount > 0)
                    _GR.Lines.Add();

                string itemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colItmCd").Cells.Item(row).Specific).Value;
                string IntitemCode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colICom").Cells.Item(row).Specific).Value;
                string qtyStr = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("colQty").Cells.Item(row).Specific).Value;
                string lineId = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("rowHead").Cells.Item(row).Specific).Value;
                double quantity = string.IsNullOrEmpty(qtyStr) ? 0 : Convert.ToDouble(qtyStr);
                double itemCost = 0;

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"SELECT \"StockPrice\" FROM IGE1 WHERE \"ItemCode\" = '{itemCode}' and \"DocEntry\" ='{IssueEntry}'");

                if (!rs.EoF)
                {
                    itemCost = Convert.ToDouble(rs.Fields.Item(0).Value);
                }
                _GR.Lines.UnitPrice = itemCost; 
                //_GR.Lines.Price = itemCost;
                if(IntitemCode == "")
                {
                    _GR.Lines.ItemCode = itemCode;
                }
                else
                {
                    _GR.Lines.ItemCode = IntitemCode;
                }
                _GR.Lines.Quantity = quantity;
                _GR.Lines.WarehouseCode = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value;
                if (_RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString() == "")
                {
                    //Utilities.Message("Inter Branch GL Code not found for selected warehouse.",
                    //      SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception("Inter Branch GL Code is missing for the selected warehouse.");

                    errorFlag = true;
                    return;
                }
                else 
                { 
                    _GR.Lines.AccountCode = _RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString();
                }

                // Get batches from @BATCHS
                string batchQry = $@"
                 SELECT ""U_BatchNum"", ""U_ReqQty"", ""U_ExpDate"" FROM ""@BATCHS"" WHERE ""U_IntDocEntry"" = '{docEntry}'
                 AND ""U_ItemCode"" = '{itemCode}'AND ""U_LineNum"" = '{lineId}'   
                 AND (IFNULL(""U_BatchPost"",'') = '' OR ""U_BatchPost"" <> 'Y')";
                SAPbobsCOM.Recordset rsBatch = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rsBatch.DoQuery(batchQry);

                int batchLine = 0;
                while (!rsBatch.EoF)
                {
                    _GR.Lines.BatchNumbers.SetCurrentLine(batchLine);
                    _GR.Lines.BatchNumbers.BatchNumber = rsBatch.Fields.Item("U_BatchNum").Value.ToString();
                    _GR.Lines.BatchNumbers.Quantity = Convert.ToDouble(rsBatch.Fields.Item("U_ReqQty").Value);
                    string expDateStr = rsBatch.Fields.Item("U_ExpDate").Value.ToString();
                    if (!string.IsNullOrEmpty(expDateStr))
                    {
                        DateTime Expdate = DateTime.ParseExact(expDateStr, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture);
                        _GR.Lines.BatchNumbers.ExpiryDate = Utilities.ToDate(Expdate.ToString("yyyyMMdd"));
                    }
                    _GR.Lines.BatchNumbers.Add();
                    batchLine++;
                    rsBatch.MoveNext();
                }

                _lineCount++;
            }

            int retVal = _GR.Add();
            if (retVal == 0)
            {
                string newDocEntry;
                Utilities.Application.Company.GetNewObjectCode(out newDocEntry);
                ReceiptEntry = newDocEntry;
                errorFlag = false;
            }
            else
            {
                _Error = Utilities.Application.Company.GetLastErrorDescription();
                //Utilities.Message(err, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                errorFlag = true;
            }
        }
        #endregion

        #region RefreshRowIndex
        private void RefreshRowIndex(SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                int rowCount = oMatrix.RowCount-1;
                for (int i = 1; i <= rowCount; i++)
                {
                    // Assuming first column UID is "LineId" or "V_-1" for default RowNum
                    oMatrix.SetCellWithoutValidation(i, "rowHead", i.ToString());
                }

                // Load data back into the matrix
                //oMatrix.LoadFromDataSource();
                int row = oMatrix.RowCount;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region RenumberMatrix
        private void RenumberMatrixAndDataSource(SAPbouiCOM.Matrix oMatrix, SAPbouiCOM.DBDataSource dbDataSource)
        {

            oMatrix.FlushToDataSource();

            
            for (int i = dbDataSource.Size - 1; i >= 0; i--)
            {
                string itemCode = dbDataSource.GetValue("U_ItemCode", i).Trim();
                if (string.IsNullOrEmpty(itemCode))
                {
                    dbDataSource.RemoveRecord(i);
                }
            }

            if (dbDataSource.Size > 0)
            {
                for (int i = 0; i < dbDataSource.Size; i++)
                {
                    dbDataSource.SetValue("U_LineId", i, (i + 1).ToString());
                }
            }

            oMatrix.LoadFromDataSource();

           
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oMatrix.SetCellWithoutValidation(i, "rowHead", i.ToString());
            }

        }
        #endregion

        #region EnableFormmenu
        private void EnableFormMenu()
        {
            _Form.EnableMenu("1287", true);
            _Form.EnableMenu("1288", true);
            _Form.EnableMenu("1289", true);
            _Form.EnableMenu("1290", true);
            _Form.EnableMenu("1291", true);
            _Form.EnableMenu("1292", true);
            _Form.EnableMenu("1293", true);
            _Form.EnableMenu("779", true);
        }
        #endregion
        

        #region LineInfo
        public class LineInfo
        {
            public string ItemCode { get; set; }
            public string LineId { get; set; }
            public string DocEntry { get; set; }
        }
        #endregion

        #region Disable Form item on Load
        private void Disable_Form_Item()
        {
            string issueNo = ((SAPbouiCOM.EditText)Form.Items.Item("txtGoodIN").Specific).Value;
            string receiptNo = ((SAPbouiCOM.EditText)Form.Items.Item("txtGoodRN").Specific).Value;

            if (issueNo != "" && receiptNo != "") 
            { 
            //_Form.Items.Item("txtDate").Enabled = false;
            _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            _Form.Items.Item("mtx1").Enabled = false;

            }
       
        }
        #endregion

    }
}
