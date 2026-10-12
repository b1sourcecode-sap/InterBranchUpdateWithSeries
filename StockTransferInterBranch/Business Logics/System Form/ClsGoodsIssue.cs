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


namespace StockTransferInterBranch
{
    class ClsGoodsIssue : Base
    {
        #region VARIABLE DECLARATION
        private SAPbobsCOM.Recordset _RS = null, _RSGR = null, _RSGRU = null,_RSSeries = null,_RSGLCode = null;
        SAPbouiCOM.Item _ExistItem, _NewItem;
        SAPbouiCOM.Button _Button;
        string _DocEntry = "", _DocNum = "";
        CrystalDecisions.CrystalReports.Engine.ReportDocument objCryRep;
        string ReportPath = string.Empty;
        string _ImportFileName = "";
        private bool flag;
        private Hashtable _lines = new Hashtable();
        SAPbobsCOM.Recordset _RS2 = null;
        public string oSql2 = string.Empty;
        public string GRPODelNo = string.Empty, BranchCd = string.Empty;
        SAPbobsCOM.Recordset _Rs1 = null, _RSGRPO = null;
        SAPbouiCOM.Matrix _Matrix;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsGoodsIssue()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsGoodsIssue()
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
                   
                    #region ITEM PRESS EVENT

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "GR")
                        {
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_RecvWhse", 0).ToString() == "")
                            {
                                Utilities.Message("Please select receiving warehouse", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer")
                            {
                                if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() == "" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsRecEntry", 0).ToString() == "")
                                {
                                    Utilities.Message("Please wait operation is in progress....", SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    InventoryGoodsReceipt();
                                }
                                else
                                {
                                    Utilities.Message("InterBranch Transfer Already Created for this issue", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                Utilities.Message("Can not make interbranch transfer document", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
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
                        if (pVal.ItemUID == "U_IssueType" && (_Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || _Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE))
                        {
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() != "InterBranchTransfer")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;
                        }
                        if (pVal.ItemUID == "U_IssueType" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer")
                            {
                                _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("13").Specific;

                                for (int i = 0; i < _Matrix.VisualRowCount - 1; i++)
                                {
                                    string _QryGLCode = "select \"U_InterBranchGLCode\" from OWHS where \"WhsCode\" = '" + ((SAPbouiCOM.EditText)_Matrix.Columns.Item("15").Cells.Item(i+1).Specific).Value + "'";
                                    Utilities.ExecuteSQL(ref _RSGLCode, _QryGLCode);
                                    if (_RSGLCode.RecordCount > 0)
                                        ((SAPbouiCOM.EditText)_Matrix.Columns.Item("59").Cells.Item(i+1).Specific).Value = _RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString();
                                }
                            }
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if(pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;
                        }
                        break;

                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.ItemUID == "13" && (pVal.ColUID == "15" || pVal.ColUID == "1") && pVal.InnerEvent == true)
                        {
                            _Matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("13").Specific;
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer")
                            {
                                string _QryGLCode = "select \"U_InterBranchGLCode\" from OWHS where \"WhsCode\" = '" + ((SAPbouiCOM.EditText)_Matrix.Columns.Item("15").Cells.Item(pVal.Row).Specific).Value + "'";
                                Utilities.ExecuteSQL(ref _RSGLCode, _QryGLCode);
                                if(_RSGLCode.RecordCount > 0)
                                ((SAPbouiCOM.EditText)_Matrix.Columns.Item("59").Cells.Item(pVal.Row).Specific).Value = _RSGLCode.Fields.Item("U_InterBranchGLCode").Value.ToString();
                            }
                        }
                        break;

                    #endregion

                    #region et_VALIDATE

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.ItemUID == "4")
                        {
                            //SAPbobsCOM.Recordset _RS1 = null;
                            //if (((SAPbouiCOM.EditText)_Form.Items.Item("4").Specific).Value != "")
                            //{
                            //    GRPODelNo = _Form.DataSources.DBDataSources.Item("ODLN").GetValue("U_GRPONum", 0).ToString();
                            //    if (GRPODelNo != "")
                            //        _Form.Items.Item("GRPO").Enabled = false;
                            //    else
                            //        _Form.Items.Item("GRPO").Enabled = true;
                            //}
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
                        case Constants.System_Menus.mnu_ADD:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {

                        case Constants.System_Menus.mnu_ADD:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_FIND:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_FIRST:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_LAST:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_NEXT:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_PREVIOUS:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;

                        case Constants.System_Menus.mnu_REFRESH:
                            if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                                _Form.Items.Item("GR").Visible = false;
                            else
                                _Form.Items.Item("GR").Visible = true;

                            break;
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
        internal void FormDefault()
        {
            try
            {
                    _ExistItem = _Form.Items.Item("2");
                    _NewItem = _Form.Items.Add("GR", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    _NewItem.Left = _ExistItem.Left + _ExistItem.Width + 10;
                    _NewItem.Width = _ExistItem.Width + 20;
                    _NewItem.Top = _ExistItem.Top;
                    _Button = (SAPbouiCOM.Button)_NewItem.Specific;
                    _Button.Caption = "Create G.Receipt";
                _Form.Items.Item("GR").Visible = false;
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
        #endregion

        #region InventoryGoodsReceipt

        private void InventoryGoodsReceipt()
        {

            int _lineCount = 0;
            SAPbobsCOM.BoObjectTypes objType = SAPbobsCOM.BoObjectTypes.oInventoryGenEntry;
            SAPbobsCOM.Documents document;
            SAPbobsCOM.Documents _GR = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
            document = (SAPbobsCOM.Documents)Utilities.Application.Company.GetBusinessObject(objType);
           

            string _QryGRData = "select T0.\"DocNum\",T0.\"DocEntry\",T0.\"DocDate\",T0.\"TaxDate\",T3.\"U_InterComItemCode\" \"ItemCode\", " +
                                "T1.\"Dscription\",T1.\"Quantity\", " +
                                " T0.\"U_RecvWhse\",T2.\"BPLid\",T3.\"ManBtchNum\", T0.\"Comments\",T1.\"LineNum\",T1.\"Price\",T2.\"U_InterBranchGLCode\",T1.\"WhsCode\" " +
                                " from OIGE T0 inner join IGE1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                " left outer join OWHS T2 ON T2.\"WhsCode\" = T0.\"U_RecvWhse\" " +
                                " inner join OITM T3 ON T1.\"ItemCode\" = T3.\"ItemCode\" where T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).ToString() + "' ";
            Utilities.ExecuteSQL(ref _RS, _QryGRData);

            string _QrySeries = "select \"Series\" from NNM1 where \"ObjectCode\" = '59' AND \"BPLId\" = '"+ _RS.Fields.Item("BPLid").Value.ToString() + "' and \"Indicator\" = " +
                               " (select distinct \"Indicator\" from OFPR where '"+ Convert.ToDateTime(_RS.Fields.Item("DocDate").Value.ToString()).ToString("yyyyMMdd") + "' between \"F_TaxDate\" and \"T_TaxDate\")";
            Utilities.ExecuteSQL(ref _RSSeries, _QrySeries);
            if (_RS.RecordCount > 0)
            {
                for (int i = 0; i < _RS.RecordCount; i++)
                {
                    if (_lineCount == 0)
                    {
                        DateTime parsedPostingDate = DateTime.Parse(_RS.Fields.Item("DocDate").Value.ToString());
                        DateTime parsedPDeliveryDate = DateTime.Parse(_RS.Fields.Item("TaxDate").Value.ToString());
                        int BranchCode = Convert.ToInt32(_RS.Fields.Item("BPLid").Value.ToString());
                        _GR.DocDate = parsedPostingDate;
                        _GR.DocDueDate = parsedPDeliveryDate;
                        _GR.JournalMemo = "Inter Branch Transfer";
                        _GR.Comments = "Inter Branch Transfer";
                        _GR.BPL_IDAssignedToInvoice = Convert.ToInt16(BranchCode.ToString());
                        _GR.UserFields.Fields.Item("U_GoodsIssueNo").Value = _RS.Fields.Item("DocNum").Value.ToString();
                        _GR.UserFields.Fields.Item("U_GoodsIssueEntry").Value = _RS.Fields.Item("DocEntry").Value.ToString();
                        _GR.UserFields.Fields.Item("U_IssueWhse").Value = _RS.Fields.Item("WhsCode").Value.ToString();
                        _GR.Series = Convert.ToInt16(_RSSeries.Fields.Item("Series").Value.ToString());
                        _lineCount++;
                    }
                    else
                        _GR.Lines.Add();

                    //if (_RS.Fields.Item("ManSerNum").Value.ToString() == "Y")
                    //{
                    //    string Qry = "SELECT Distinct T3.\"ItemCode\",T3.\"DistNumber\",T3.\"MnfSerial\" FROM OIGE T0 INNER JOIN IGE1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                    //                 " INNER JOIN SRI1 T2 on T1.\"ItemCode\"=T2.\"ItemCode\" and T1.\"LineNum\" = T2.\"BaseLinNum\" and (T1.\"DocEntry\"=T2.\"BaseEntry\" and T1.\"ObjType\"=T2.\"BaseType\") " +
                    //                 " INNER JOIN OSRN T3 on T3.\"ItemCode\"=T2.\"ItemCode\" and T2.\"SysSerial\"=T3.\"SysNumber\" where  T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).ToString() + "' and T3.\"ItemCode\" = '" + _RS.Fields.Item("ItemCode").Value.ToString() + "' AND T1.\"LineNum\" = " + _RS.Fields.Item("LineNum").Value.ToString() + " ";
                    //    Utilities.ExecuteSQL(ref _Rs1, Qry);

                    //    if (_Rs1.RecordCount > 0)
                    //    {
                    //        for (int j = 0; j <= _Rs1.RecordCount - 1; j++)
                    //        {
                    //            _GR.Lines.SerialNumbers.SetCurrentLine(j);
                    //            _GR.Lines.SerialNumbers.ManufacturerSerialNumber = _Rs1.Fields.Item("MnfSerial").Value.ToString();
                    //            _GR.Lines.SerialNumbers.InternalSerialNumber = _Rs1.Fields.Item("DistNumber").Value.ToString();
                    //            _GR.Lines.SerialNumbers.Add();
                    //            _Rs1.MoveNext();
                    //        }
                    //    }
                    //}

                    if (_RS.Fields.Item("ManBtchNum").Value.ToString() == "Y")
                    {
                        string Qry = "SELECT DISTINCT T3.\"ItemCode\", T3.\"DistNumber\", T3.\"ExpDate\" " +
                                     "FROM OIGE T0 " +
                                     "INNER JOIN IGE1 T1 ON T0.\"DocEntry\" = T1.\"DocEntry\" " +
                                     "INNER JOIN IBT1 T2 ON T1.\"ItemCode\" = T2.\"ItemCode\" " +
                                     "AND T1.\"LineNum\" = T2.\"BaseLinNum\" " +
                                     "AND T1.\"DocEntry\" = T2.\"BaseEntry\" " +
                                     "AND T1.\"ObjType\" = T2.\"BaseType\" " +
                                     "INNER JOIN OBTN T3 ON T3.\"ItemCode\" = T2.\"ItemCode\" " +
                                     "AND T2.\"BatchNum\" = T3.\"DistNumber\" " +
                                     "WHERE T0.\"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).ToString() + "' " +
                                     "AND T3.\"ItemCode\" = '" + _RS.Fields.Item("ItemCode").Value.ToString() + "' " +
                                     "AND T1.\"LineNum\" = " + _RS.Fields.Item("LineNum").Value.ToString();

                        Utilities.ExecuteSQL(ref _Rs1, Qry);

                        if (_Rs1.RecordCount > 0)
                        {
                            for (int j = 0; j <= _Rs1.RecordCount - 1; j++)
                            {
                                _GR.Lines.BatchNumbers.SetCurrentLine(j);
                                _GR.Lines.BatchNumbers.BatchNumber = _Rs1.Fields.Item("DistNumber").Value.ToString();

                                if (_Rs1.Fields.Item("ExpDate").Value != null && _Rs1.Fields.Item("ExpDate").Value.ToString() != "")
                                {
                                    _GR.Lines.BatchNumbers.ExpiryDate = Convert.ToDateTime(_Rs1.Fields.Item("ExpDate").Value);
                                }
                                _GR.Lines.BatchNumbers.Add();
                                _Rs1.MoveNext();
                            }
                        }
                    }

                    _GR.Lines.ItemCode = _RS.Fields.Item("ItemCode").Value.ToString();
                    _GR.Lines.WarehouseCode = _RS.Fields.Item("U_RecvWhse").Value.ToString();
                    _GR.Lines.Quantity = double.Parse(_RS.Fields.Item("Quantity").Value.ToString());
                    _GR.Lines.Price = double.Parse(_RS.Fields.Item("Price").Value.ToString());
                    _GR.Lines.AccountCode = _RS.Fields.Item("U_InterBranchGLCode").Value.ToString();
                    _RS.MoveNext();
                }
            }
            int RetVal = _GR.Add();
            if (RetVal == 0)
            {
                string objCode = "";
                Utilities.Application.Company.GetNewObjectCode(out objCode);
                string StrQry1 = "select \"DocNum\" from OIGN where \"DocEntry\"  = '" + objCode + "'";
                Utilities.ExecuteSQL(ref _RSGR, StrQry1);

                string StrQry = "update OIGE set \"U_GoodsReceiptNo\" = '" + _RSGR.Fields.Item("DocNum").Value.ToString() + "',\"U_GoodsRecEntry\" = '" + objCode + "' where \"DocEntry\" = '" + _Form.DataSources.DBDataSources.Item("OIGE").GetValue("DocEntry", 0).ToString() + "'";
                Utilities.ExecuteSQL(ref _RSGRU, StrQry);
                Utilities.Message("Operation completed sucessfully", SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            else
            {
                string GetLastErrorDescription = Utilities.Application.Company.GetLastErrorDescription();
                Utilities.Message(GetLastErrorDescription, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

    }
}
