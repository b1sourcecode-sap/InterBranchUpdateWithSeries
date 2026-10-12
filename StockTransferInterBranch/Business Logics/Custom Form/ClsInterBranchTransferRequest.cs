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
    class ClsInterBranchTransferRequest : Base
    {
        #region VARIABLE DECLARATION
        private SAPbobsCOM.Recordset _RS = null, _RSGR = null, _RSGRU = null, _RSSeries = null, _RSGLCode = null,_RSBPL =null ;
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
        public SAPbouiCOM.Matrix _Matrix;
        private SAPbouiCOM.DBDataSource _MasterDBDataSource = null, _ChildDBDataSource = null;
        SAPbouiCOM.ChooseFromListEvent _SysCFLEvent = null;
        private SAPbouiCOM.ComboBox _ComboBoxC, _ComboBoxCRow;
        private SAPbouiCOM.ComboBox _ComboBoxB;
        public string ErrMsg = string.Empty;
        public int ErrCode = 0;
        private SAPbouiCOM.Form _UDFForm;


        #endregion

        #region CONSTRUCTOR & DISTRUCTOR

        public ClsInterBranchTransferRequest()
            : base()
        {
            //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
            //
            // TODO: Add constructor logic here
            //
        }
        ~ClsInterBranchTransferRequest()
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
                        string status = ((SAPbouiCOM.EditText)Form.Items.Item("txtStatus").Specific).Value;
                        Utilities.SerializedMartix(ref _Form, ref _Matrix);
                        AddSequenceNumbersToMatrix(_Matrix);
                        Disable_Form_Item(status);
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

        public override void Item_Event(string oFormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {

            #region BEFORE ACTION ACTION

            if (pVal.Before_Action == true)
            {

                switch (pVal.EventType)
                {
                    #region Choose From List
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                    if (pVal.ItemUID == "mtxIBTRR" && pVal.ColUID == "colItmCd")
                        {
                            string fromBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value.Trim();
                            string toBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value.Trim();
                            string HeadRemarks = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRmrks").Specific).Value.Trim();
                            if ((string.IsNullOrWhiteSpace(fromBranch) || string.IsNullOrWhiteSpace(toBranch)))
                            {
                                Utilities.Message("Both from and to warehouse must be selected at first", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            if (HeadRemarks == "")
                            {
                                Utilities.Message("Header Remarks can not be blank before choosing items", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                        }
                   break;
                    #endregion

                    #region item Pressed Event
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        #region validation

                        if (pVal.ItemUID == "1" && Form.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {

                            string fromBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value;
                            string toBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value;

                            if (((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbBranch").Specific).Value == "")
                            {
                                Utilities.Message("Branch must be selected.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtRmrks").Specific).Value == "")
                            {
                                Utilities.Message("Remarks cannot left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }

                            else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtFWhse").Specific).Value == "")
                            {
                                Utilities.Message("From warehouse cannot left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            else if (((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value == "")
                            {
                                Utilities.Message("To warehouse cannot left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            else if (fromBranch == toBranch)
                            {
                                Utilities.Message("Inter-branch transfer cannot be created when the source and destination branches are the same. Please choose different branches assigned to their respective warehouses.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            int _rowCount = _Matrix.RowCount;
                            if (_rowCount == 1)
                            {
                                if (!Validation(_Matrix, 1))
                                {
                                    Utilities.Message("Matrix row cannot be blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                            else
                            {
                                for (int i = 1; i <= _rowCount - 1; i++)  // Skip the last row
                                {
                                    if (!Validation(_Matrix, i))
                                    {
                                        BubbleEvent = false;
                                        return;
                                    }
                                }

                            }

                        }

                        #endregion
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
                        if (pVal.ItemUID == "cmbBranch" && _Matrix.RowCount == 0)
                        {
                            _Matrix.AddRow();
                        }


                        break;

                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1" && _Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            //if (_Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_IssueType", 0).ToString() == "InterBranchTransfer" && _Form.DataSources.DBDataSources.Item("OIGE").GetValue("U_GoodsReceiptNo", 0).ToString() != "")
                            //    _Form.Items.Item("GR").Visible = false;
                           
                        }
                        if (pVal.ItemUID == "btnCpyTo")
                        {

                            this._Object = new ClsInterBranchTransfer();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.InterBranchTransfer);
                            ((ClsInterBranchTransfer)_Object).FormDefault(_Form.DataSources.DBDataSources.Item("@IBTFRRH").GetValue("DocEntry", 0).ToString());
                            _Form.Close();
                        }

                        #region Close Button
                        if (pVal.ItemUID == "btnClose" && ((_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) || (_Form.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)))
                        {
                            if (((SAPbouiCOM.EditText)_Form.Items.Item("txtCRmrks").Specific).Value == "")
                            {
                                Utilities.Message("Closing Remarks cannot left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                                return;
                            }
                            else
                            {
                                _Form.DataSources.DBDataSources.Item(0).SetValue("U_Status", 0, "Closed");
                                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                _Form.Items.Item("1").Click();
                            }
                            _Form.Close();

                        }


                        #endregion


                        break;


                    #region CFL values
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        _SysCFLEvent = (SAPbouiCOM.ChooseFromListEvent)pVal;

                        SetWarehouseAndBranchName("txtFWhse", "U_FrmWhse", "txtFBrnch", pVal);
                        SetWarehouseAndBranchName("txtTWhse", "U_ToWhse", "txtTBrnch", pVal);

                        if (pVal.ItemUID == "mtxIBTRR" && pVal.ColUID == "colItmCd" && _SysCFLEvent.SelectedObjects != null)
                        {

                            try
                            {
                                
                                string itemCode = _SysCFLEvent.SelectedObjects.GetValue("ItemCode", 0).ToString().Trim();
                                string itemName = _SysCFLEvent.SelectedObjects.GetValue("ItemName", 0).ToString().Trim();
                                string fromBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value.Trim();
                                string toBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value.Trim();
                                string interComCode = string.Empty;
                                if ((fromBranch == "UNIT A" && toBranch == "UNIT B") ||
                                    (fromBranch == "UNIT B" && toBranch == "UNIT A"))
                                {
                                    interComCode = _SysCFLEvent.SelectedObjects.GetValue("U_InterComItemCode", 0).ToString().Trim();
                                }
                                else
                                {
                                    interComCode = string.Empty;
                                }
                                string remarks = ((SAPbouiCOM.EditText)_Form.Items.Item("txtRmrks").Specific).Value;

                                _Matrix.FlushToDataSource();
                                _ChildDBDataSource.Offset = pVal.Row - 1;
                                _ChildDBDataSource.SetValue("U_ItemCode", _ChildDBDataSource.Offset, itemCode);
                                _ChildDBDataSource.SetValue("U_ItemName", _ChildDBDataSource.Offset, itemName);
                                _ChildDBDataSource.SetValue("U_InterComCode", _ChildDBDataSource.Offset, interComCode);
                                _ChildDBDataSource.SetValue("U_Qty", _ChildDBDataSource.Offset, "1");
                                _ChildDBDataSource.SetValue("U_OpnQty", _ChildDBDataSource.Offset, "1");
                                _ChildDBDataSource.SetValue("U_Remarks", _ChildDBDataSource.Offset, remarks);
                                string toWarehouse = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTWhse").Specific).Value.Trim();
                                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                string query = $@" SELECT T0.""Locked"" FROM OITW T0 JOIN OWHS T1 ON T0.""WhsCode"" = T1.""WhsCode"" WHERE T0.""WhsCode"" = '{toWarehouse}' AND T0.""ItemCode"" = '{itemCode}'";
                                oRec.DoQuery(query);
                                string lockedStatus = "Unlocked";
                                if (!oRec.EoF)
                                {
                                    string dbValue = oRec.Fields.Item("Locked").Value.ToString().Trim();
                                    lockedStatus = dbValue == "Y" ? "Locked" : "Unlocked";
                                }

                                _ChildDBDataSource.SetValue("U_ToWhseStatus", _ChildDBDataSource.Offset, lockedStatus);
                                _Matrix.LoadFromDataSource();
                                if (pVal.Row == _Matrix.RowCount)
                                {
                                    AddNewRow();
                                    _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                }

                            }
                            catch
                            {
                                Utilities.Application.Company.GetLastError(out ErrCode, out ErrMsg);
                            }

                        }
                        break;
                        #endregion

                        #endregion
                        break;


                    #region et_VALIDATE

                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        if (pVal.BeforeAction == false && pVal.ItemUID == "mtxIBTRR" && pVal.ColUID == "colQty")
                        {
                            var colQty = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colQty").Cells.Item(pVal.Row).Specific);
                            var colOQty = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colOQty").Cells.Item(pVal.Row).Specific);
                            string qtyValue = colQty.Value.Trim();
                            if (!string.IsNullOrEmpty(qtyValue))
                            {
                                colOQty.Value = qtyValue;  // Auto-fill open quantity
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
                            Disable_Form_Item("");

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

        #region Add New Row
        private void AddNewRow()
        {
            SAPbouiCOM.DBDataSource dbData = default(SAPbouiCOM.DBDataSource);
            dbData = (SAPbouiCOM.DBDataSource)_Form.DataSources.DBDataSources.Item(1);
            int selRow = _Matrix.GetNextSelectedRow();
            dbData.Clear();
            _Matrix.AddRow(1, selRow - 1);
            AddSequenceNumbersToMatrix(_Matrix);
        }
        #endregion

        #region Add Sequence No
        private void AddSequenceNumbersToMatrix(SAPbouiCOM.Matrix oMatrix)
        {
            oMatrix.Columns.Item("rowHead").Visible = true;
            for (int i = 1; i <= oMatrix.RowCount; i++)
            {
                oMatrix.Columns.Item("rowHead").Cells.Item(i).Specific.Value = i.ToString();

            }
        }
        #endregion

        #region SetWarehouseAndBranchName
        private void SetWarehouseAndBranchName(string itemUID, string targetUDF, string branchFieldItemID, SAPbouiCOM.ItemEvent pVal)
        {
            if (pVal.ItemUID == itemUID && _SysCFLEvent.SelectedObjects != null)
            {
                string selectedWhsCode = _SysCFLEvent.SelectedObjects.GetValue("WhsCode", 0).ToString().Trim();

                _MasterDBDataSource.SetValue(targetUDF, 0, selectedWhsCode);

                SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string query = $@"SELECT T0.""BPLName""FROM OBPL T0 JOIN OWHS T1 ON T0.""BPLId"" = T1.""BPLid"" WHERE T1.""WhsCode"" = '{selectedWhsCode}'";
                oRec.DoQuery(query);
                if (!oRec.EoF)
                {
                    ((SAPbouiCOM.EditText)_Form.Items.Item(branchFieldItemID).Specific).Value = oRec.Fields.Item("BPLName").Value.ToString().Trim();
                }
            }
        }

        #endregion

        #region matrix column validation

        private bool Validation(SAPbouiCOM.Matrix _Matrix, int i)
        {
            try
            {
                if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("colItmCd").Cells.Item(i).Specific).Value == "")
                {
                    Utilities.Message("Item Code cannot be left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("colItmDsp").Cells.Item(i).Specific).Value == "")
                {
                    Utilities.Message("Item Name cannot be left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                string qtyVal = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colQty").Cells.Item(i).Specific).Value;
                if (string.IsNullOrWhiteSpace(qtyVal) || Convert.ToDouble(qtyVal) <= 0)
                {
                    Utilities.Message("Quantity must be greater than zero.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                
                if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("colRmrks").Cells.Item(i).Specific).Value == "")
                {
                    Utilities.Message("Remarks in row cannot be left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                string fromBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtFBrnch").Specific).Value.Trim();
                string toBranch = ((SAPbouiCOM.EditText)_Form.Items.Item("txtTBrnch").Specific).Value.Trim();

                bool isSpecialCase =
                    (!fromBranch.Equals("UNIT A", StringComparison.OrdinalIgnoreCase) &&
                     !fromBranch.Equals("UNIT B", StringComparison.OrdinalIgnoreCase))
                    ||
                    (!toBranch.Equals("UNIT A", StringComparison.OrdinalIgnoreCase) &&
                     !toBranch.Equals("UNIT B", StringComparison.OrdinalIgnoreCase));

                if (!isSpecialCase)
                {
                    if (((SAPbouiCOM.EditText)_Matrix.Columns.Item("colICom").Cells.Item(i).Specific).Value == "")
                    {
                        Utilities.Message("Inter Company Code cannot be left blank.", SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
                return true;
            }
            catch
            {
                throw;
            }
        }

        #endregion

        #region Disable Form item on Load
        private void Disable_Form_Item(string status)
        {
            double totalQty = 0;
            for (int i = 1; i <= _Matrix.RowCount; i++)
            {
                string qtyStr = ((SAPbouiCOM.EditText)_Matrix.Columns.Item("colOQty").Cells.Item(i).Specific).Value;

                if (double.TryParse(qtyStr, out double qty))
                {
                    totalQty += qty;
                }
            }
            if (status == "Approved" && totalQty != 0)
            {
                _Form.Items.Item("btnCpyTo").Enabled = true;
                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("cmbBranch").Enabled = false;
                _Form.Items.Item("txtFWhse").Enabled = false;
                _Form.Items.Item("txtTWhse").Enabled = false;
                _Form.Items.Item("txtFBrnch").Enabled = false;
                _Form.Items.Item("txtTBrnch").Enabled = false;
                _Form.Items.Item("txtRmrks").Enabled = false;
                _Form.Items.Item("txtCRmrks").Enabled = true;
                _Form.Items.Item("txtDate").Enabled = false;
                _Form.Items.Item("txtDocNum").Enabled = false;
                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Matrix.Columns.Item("colItmCd").Editable = false;
                _Matrix.Columns.Item("colItmDsp").Editable = false;
                _Matrix.Columns.Item("colQty").Editable = false;
                _Matrix.Columns.Item("colICom").Editable = false;
                _Matrix.Columns.Item("colRmrks").Editable = false;

            }
           
            if (status == "Approved" && totalQty == 0)
            {
                if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                {
                    _Form.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Closed";
                    _Form.Items.Item("1").Click();
                }

                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("btnCpyTo").Enabled = false;
                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("cmbBranch").Enabled = false;
                _Form.Items.Item("txtFWhse").Enabled = false;
                _Form.Items.Item("txtTWhse").Enabled = false;
                _Form.Items.Item("txtFBrnch").Enabled = false;
                _Form.Items.Item("txtTBrnch").Enabled = false;
                _Form.Items.Item("txtRmrks").Enabled = false;
                _Form.Items.Item("txtCRmrks").Enabled = false;
                _Form.Items.Item("txtCrtBy").Enabled = false;
                _Form.Items.Item("txtDocNum").Enabled = false;
                //_Form.Items.Item("txtTEnt").Enabled = false;
                _Form.Items.Item("txtDate").Enabled = false;
                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("mtxIBTRR").Enabled = false;
                //_Form.Items.Item("1").Click();
            }

            if (status == "Closed")
            {
                _Form.Items.Item("btnCpyTo").Enabled = false;
                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("cmbBranch").Enabled = false;
                _Form.Items.Item("txtFWhse").Enabled = false;
                _Form.Items.Item("txtTWhse").Enabled = false;
                _Form.Items.Item("txtFBrnch").Enabled = false;
                _Form.Items.Item("txtTBrnch").Enabled = false;
                _Form.Items.Item("txtRmrks").Enabled = false;
                _Form.Items.Item("txtCRmrks").Enabled = false;
                _Form.Items.Item("txtCrtBy").Enabled = false;
                _Form.Items.Item("txtDocNum").Enabled = false;
                //_Form.Items.Item("txtTEnt").Enabled = false;
                _Form.Items.Item("txtDate").Enabled = false;
                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("mtxIBTRR").Enabled = false;
            }
            if (status == "Rejected" || status == "Pending")
            {
                _Form.Items.Item("btnCpyTo").Enabled = false;
                _Form.Items.Item("txtStatus").Enabled = false;
                _Form.Items.Item("cmbBranch").Enabled = false;
                _Form.Items.Item("txtFWhse").Enabled = false;
                _Form.Items.Item("txtTWhse").Enabled = false;
                _Form.Items.Item("txtFBrnch").Enabled = false;
                _Form.Items.Item("txtTBrnch").Enabled = false;
                _Form.Items.Item("txtRmrks").Enabled = false;
                _Form.Items.Item("txtCRmrks").Enabled = false;
                _Form.Items.Item("txtCrtBy").Enabled = false;
                _Form.Items.Item("txtDocNum").Enabled = false;
                _Form.Items.Item("txtDate").Enabled = false;
                _Form.Items.Item("txtFocus").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                _Form.Items.Item("mtxIBTRR").Enabled = false;
            }
            if (_Form.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
            {
                _Form.Items.Item("txtStatus").Enabled = true;
                _Form.Items.Item("txtFWhse").Enabled = true;
                _Form.Items.Item("txtTWhse").Enabled = true;
                _Form.Items.Item("txtDate").Enabled = true;
                _Form.Items.Item("txtDocNum").Enabled = true;
                _Form.Items.Item("txtCRmrks").Enabled = false;
                _Form.Items.Item("btnCpyTo").Enabled = false;

            }

        }
        #endregion

        #region Form Default
        internal void FormDefault()
        {
            _Form.Freeze(true);
            _Form.DataBrowser.BrowseBy = "txtDocNum";
            _Form.Settings.Enabled = true;
            UserObjectsMD oUserObjectMD = (UserObjectsMD)Utilities.Application.Company.GetBusinessObject(BoObjectTypes.oUserObjectsMD);
            _Matrix = (SAPbouiCOM.Matrix)Form.Items.Item("mtxIBTRR").Specific;
            _MasterDBDataSource = _Form.DataSources.DBDataSources.Item("@IBTFRRH");
            _ChildDBDataSource = _Form.DataSources.DBDataSources.Item("@IBTFRRC");
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDocNum").Specific).Value = (Utilities.getMaxColumnValueNum("@IBTFRRH", "U_DocNum"));
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtCrtBy").Specific).Value = Utilities.Application.SBO_Application.Company.UserName;
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtStatus").Specific).Value = "Draft";
            ((SAPbouiCOM.EditText)_Form.Items.Item("txtDate").Specific).Value = DateTime.Now.ToString("yyyyMMdd");
            string bpl = "SELECT \"BPLId\",\"BPLName\" FROM OBPL";
            Utilities.ExecuteSQL(ref _RSBPL, bpl);
            while (!_RSBPL.EoF)
            {
                string branchId = _RSBPL.Fields.Item("BPLId").Value.ToString();
                string branchName = _RSBPL.Fields.Item("BPLName").Value.ToString();
                ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbBranch").Specific).ValidValues.Add(branchId,branchName);
                ((SAPbouiCOM.ComboBox)_Form.Items.Item("cmbBranch").Specific).Item.DisplayDesc = true;
                _RSBPL.MoveNext();
            }
            _Matrix.Columns.Item("colItmCd").Editable = true;
            _Matrix.Columns.Item("colQty").Editable = true;
            _Matrix.Columns.Item("colRmrks").Editable = true;
             EnableFormMenu();
            _Form.Freeze(false);
            _Form.PaneLevel = 1;
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
            _Form.EnableMenu("779", true);
        }
        #endregion

    }
}
