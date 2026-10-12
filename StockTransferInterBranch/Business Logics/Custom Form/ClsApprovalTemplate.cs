using SAPbouiCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StockTransferInterBranch
{ 
    class ClsApprovalTemplate : Base
    {


        #region VARIABLE DECLRATION
        private SAPbobsCOM.Recordset _Rs, _RS, _rs, _RSD, _RSBom, _RSProd, _RSRoute, _RSUsr, _RSMgItm, _RSBomP;

    #endregion


    #region CONSTRUCTOR & DISTRUCTOR

    public ClsApprovalTemplate()
        : base()
    {
        //_Form  = Utilities.Application.SBO_Application.Forms.ActiveForm;
        //
        // TODO: Add constructor logic here
        //

    }
    ~ClsApprovalTemplate()
    {
    }

    #endregion


    public override void Item_Event(string FormUID, ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
        {
            #region Before Action
            if (pVal.BeforeAction)
            {
                switch (pVal.EventType)
                {

                }
            }
            #endregion

            #region After Action
            else
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:


                        if (pVal.ItemUID == "btnApprove")
                        {
                            ProcessSelections("Approved");
                            
                        }
                        if (pVal.ItemUID == "btnReject")
                        {
                            ProcessSelections("Rejected");
                            
                        }
                        if (pVal.ItemUID == "btnPending")
                        {
                            ProcessSelections("Pending");
                            
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:

                        if (pVal.EventType == BoEventTypes.et_MATRIX_LINK_PRESSED)
                        {
                            if (pVal.ColUID == "colDoc")
                            {
                                SAPbouiCOM.Form form = Utilities.Application.SBO_Application.Forms.Item(pVal.FormUID);
                                SAPbouiCOM.Matrix matrix = (SAPbouiCOM.Matrix)form.Items.Item("matrix").Specific;

                                string docEntry = ((SAPbouiCOM.EditText)matrix.Columns.Item("colDoc").Cells.Item(pVal.Row).Specific).Value;

                                if (!string.IsNullOrEmpty(docEntry))
                                {
                                    try
                                    {
                                        Utilities.Application.SBO_Application.OpenForm(BoFormObjectEnum.fo_UserDefinedObject, "INTBRNCHTRANSREQH", docEntry);
                                        form.Close();
                                    }
                                    catch (Exception ex)
                                    {
                                        Utilities.Application.SBO_Application.StatusBar.SetText($"Error opening UDO form: {ex.Message}", BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error);
                                    }
                                }
                            }
                        }


                        break;
                }
            }
            #endregion

        }

        #region Form Default
        internal void FormDefault()
        {
            string userCode = Utilities.Application.SBO_Application.Company.UserName;
            Utilities.ExecuteSQL(ref _RSUsr, $"SELECT \"U_TransAuth\" FROM \"OUSR\" WHERE \"USER_CODE\" = '{userCode.Replace("'", "''")}'");
            string check = _RSUsr.Fields.Item("U_TransAuth").Value.ToString();


            // Step 4: Check authorization
            if (!_RSUsr.EoF && _RSUsr.Fields.Item("U_TransAuth").Value.ToString() == "Yes")
            {
                LoadApprovalData(userCode);
            }
            else
            {
                // Show error message
                Utilities.Application.SBO_Application.MessageBox("You are not an approver.", 1, "OK", "", "");
            }

            // Cleanup
            System.Runtime.InteropServices.Marshal.ReleaseComObject(_RSUsr);
            _RSUsr = null;

        }
        #endregion


        #region Load Approval Data
        private void LoadApprovalData(string userCode)
        {
            var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;

            // Step 1: Prepare DataTable (with safe check)
            SAPbouiCOM.DataTable dt;

            try
            {
                dt = _Form.DataSources.DataTables.Item("dtDisplay");
                dt.Clear();  // reuse if already exists
            }
            catch
            {
                // create if doesn't exist
                dt = _Form.DataSources.DataTables.Add("dtDisplay");
                dt.Columns.Add("U_Selected", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
                dt.Columns.Add("U_DocEntry", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 20);
                dt.Columns.Add("U_Branch", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Status", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_Remarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
                dt.Columns.Add("U_CreatedBy", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 100);
                dt.Columns.Add("U_ARemarks", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 254);
            }

            // Step 2: Bind Matrix Columns to DataTable
            matrix.Columns.Item("colDoc").DataBind.Bind("dtDisplay", "U_DocEntry");
            matrix.Columns.Item("colCode").DataBind.Bind("dtDisplay", "U_Branch");
            matrix.Columns.Item("colStatus").DataBind.Bind("dtDisplay", "U_Status");
            matrix.Columns.Item("colRmrks").DataBind.Bind("dtDisplay", "U_Remarks");
            matrix.Columns.Item("colCretBy").DataBind.Bind("dtDisplay", "U_CreatedBy");
            matrix.Columns.Item("colAPRmrks").DataBind.Bind("dtDisplay", "U_ARemarks");
            string sql = $@"
                            SELECT DISTINCT T0.""DocEntry"",(select ""BPLName"" from ""OBPL"" where ""BPLId"" = T0.""U_Branch"")""U_Branch"", T0.""U_Status"", T0.""U_Remarks"", T0.""U_CreatedBy"",T0.""U_CRemarks""
                            FROM ""@IBTFRRH"" T0
                            INNER JOIN ""USR6"" T1 ON T0.""U_CreatedBy"" = T1.""UserCode""
                            WHERE T0.""U_Status"" = 'Draft' AND T1.""BPLId"" NOT IN (SELECT ""BPLId"" FROM ""USR6"" WHERE ""UserCode"" = '{userCode}')";
            var rs = (SAPbobsCOM.Recordset)Utilities.Application.Company
                         .GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);

            // Step 4: Load Data into DataTable
            while (!rs.EoF)
            {
                int row = dt.Rows.Count;
                dt.Rows.Add();
                dt.SetValue("U_DocEntry", row, rs.Fields.Item("DocEntry").Value.ToString());
                dt.SetValue("U_Branch", row, rs.Fields.Item("U_Branch").Value.ToString());
                dt.SetValue("U_Status", row, rs.Fields.Item("U_Status").Value.ToString());
                dt.SetValue("U_Remarks", row, rs.Fields.Item("U_Remarks").Value.ToString());
                dt.SetValue("U_CreatedBy", row, rs.Fields.Item("U_CreatedBy").Value.ToString());
                dt.SetValue("U_ARemarks", row, rs.Fields.Item("U_CRemarks").Value.ToString());

                rs.MoveNext();
            }

            // Step 5: Load Matrix from DataTable
            matrix.LoadFromDataSource();
        }

        #endregion


        #region Process Selections
        private void ProcessSelections(string newStatus)
        {
            var matrix = (SAPbouiCOM.Matrix)_Form.Items.Item("matrix").Specific;
            var dt = _Form.DataSources.DataTables.Item("dtDisplay");

            // 1) Get the CompanyService and GeneralService for your UDO
            var companyService = Utilities.Application.Company.GetCompanyService();
            var svc = (SAPbobsCOM.GeneralService)companyService.GetGeneralService("INTBRNCHTRANSREQH");

            // 2) Retrieve the correct GeneralDataParams interface
            SAPbobsCOM.GeneralDataParams generalParams;

            try
            {
                generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(
                    SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
            }
            catch
            {
                // Fallback if enum member differs by DI API version
                generalParams = (SAPbobsCOM.GeneralDataParams)svc.GetDataInterface(
                    SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralCollectionParams);
            }

            // 3) Iterate selected rows and update UDO
            for (int i = 1; i <= matrix.RowCount; i++)
            {
                var chkColumn = matrix.Columns.Item("chkSelect");
                var cell = chkColumn.Cells.Item(i).Specific;

                bool isSelected = false;
                if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                {
                    var chk = (SAPbouiCOM.CheckBox)cell;
                    isSelected = chk.Checked;
                }
                else if (chkColumn.Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                {
                    var combo = (SAPbouiCOM.ComboBox)cell;
                    isSelected = combo.Selected.Value == "Y";
                }

                if (isSelected)
                {

                    var matrixCol = matrix.Columns.Item("colDoc");
                    var cellSpec = matrixCol.Cells.Item(i).Specific;


                    var txt = (SAPbouiCOM.EditText)cellSpec;
                    string docEntryStr = txt.Value?.Trim();


                    if (!int.TryParse(docEntryStr, out int docEntry))
                    {
                        Utilities.Application.SBO_Application.MessageBox($"Row {i}: Invalid DocEntry value '{docEntryStr}'");
                        continue;
                    }
                    generalParams.SetProperty("DocEntry", docEntry);
                    var data = svc.GetByParams(generalParams);
                    data.SetProperty("U_Status", newStatus);
                    var remarksCell = (SAPbouiCOM.EditText)matrix.Columns.Item("colAPRmrks").Cells.Item(i).Specific;
                    string remarksValue = remarksCell.Value?.Trim() ?? "";
                    data.SetProperty("U_CRemarks", remarksValue);
                    svc.Update(data);
                }

            }
            Utilities.Application.SBO_Application.MessageBox($"{newStatus} successfully applied.");

            _Form.Close();


        }
        #endregion

    }
}
