using System;
using System.Data;
using System.Data.SqlClient;
using Microsoft.ApplicationBlocks.Data;

namespace StockTransferInterBranch
{
	/// <summary>
	/// Summary description for Database.
	/// </summary>
	public sealed class Database
	{
		private Database() 
		{ 
		} 

		#region INITIALISE DATABASE

		public static void InitializeDatabase()
		{
			SAPbouiCOM.ProgressBar oProgressBar = null;
			string[] oChildTables = new string[3];
			string[] oFindColumns = new string[3];
			string[] oColumnName  = new string[2];
			try
			{
				oProgressBar = Utilities.Application.SBO_Application.StatusBar.CreateProgressBar("Initializing database...Please wait.", 50, false);
				oProgressBar.Value = 1;

				if( !Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.StartTransaction();

                #region Delivery Challan Header 

                CreateTable("DELCH", "Delivery Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("DELCH", "DocNum", "Document Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DocEntry", "Document Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "PstDt", "Posting Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DelDt", "Delivery Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "CardCode", "BP Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "CardName", "BP Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "BPRefNo", "BP Reference No", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "Series", "Series", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCH", "NepaliDt", "Nepali Date", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
               // AddColumn("DELCH", "TotBDisc", "Total Before Discount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCH", "DiscAmt", "Discount Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "TaxAmt", "Tax Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DocTotal", "Document Total", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DelNum", "Delivery Number", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "DelEntry", "Delivery Entry", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCH", "Status", "Status", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Delivery Challan Line 

                CreateTable("DELCL", "Delivery Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("DELCL", "Chk", "Check", 1, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "DocType", "Document Type", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "ItemCode", "Item Code", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "Price", "Price", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Price, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "LineTot", "Line Total", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCL", "Disc", "Discount", 11, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "LineNo", "Line No", 5, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "DocNum", "Document No", 11, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "DocEntry", "DocEntry", 11, SAPbobsCOM.BoFieldTypes.db_Numeric, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCL", "Curr", "Currency", 6, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "TaxAmt", "Tax Amount", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Sum, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "OrdDt", "Order Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "DelDt", "Delivery Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "OpnQty", "Open Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "StockQty", "Stock Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "DelQty", "Delivery Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "WhsCode", "Warehouse Code", 50, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("DELCL", "LnStatus", "Line Status", 5, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region SAP Object
                AddColumn("OWHS", "InterBranchGLCode", "Inter Branch GL Code", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "RecvWhse", "Receive Warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "IssueType", "Issue Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "GoodsReceiptNo", "Goods Receipt No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "GoodsRecEntry", "Goods Receive Entry", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGN", "GoodsIssueNo", "Goods Issue No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGN", "GoodsIssueEntry", "Goods Issue Entry", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGN", "IssueWhse", "Issue warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OITM", "InterComItemCode", "Inter Branch Item", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTQ", "GoodsIssueNo", "Goods Issue No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTQ", "IssueType", "Issue Type", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OWTQ", "GoodsIssueEntry", "Goods Issue Entry", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "ITransReqNo", "Inventory Transfer Request No", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                AddColumn("OIGE", "ITransReqEntry","Inventory Transfer Request Entry", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES);
                #endregion

                #region InterBranchTransfer Request Header
                CreateTable("IBTFRRH", "Inter Brnch trans Req Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("IBTFRRH", "Branch", "Branch", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "DocNum", "Doc Num", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "FrmWhse", "From Warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "ToWhse", "To Warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "FromBranch", "From Branch", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "ToBranch", "To Branch", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "CreatedBy", "Created By", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "IBTFRNum", "Inventory Transfer Number", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "IBTFREntry", "Inventory Transfer Entry", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "Remarks", "Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "CRemarks", "Closing Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "FBranchN", "From Branch Name", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "TBranchN", "To Branch Name", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "TransDoc", "Transfer DocEntry", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "Focus", "Focus", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);


                #endregion

                #region InterBranchTransfer Request Line
                CreateTable("IBTFRRC", "Inter Brnch trans Req Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("IBTFRRC", "ItemCode", "Item Code", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "OpnQty", "Open Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCL", "Disc", "Discount", 11, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "InterComCode", "Inter Company Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "Remarks", "Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "ToWhseStatus", "To Ware House Status", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRC", "LineId", "Sr. No", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region InterBranchTransfer Header
                CreateTable("IBTFRH", "Inter Branch trasn Header", SAPbobsCOM.BoUTBTableType.bott_Document);
                AddColumn("IBTFRH", "Branch", "Branch", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "DocDate", "Document Date", 10, SAPbobsCOM.BoFieldTypes.db_Date, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "DocNum", "Doc Num", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "FrmWhse", "From Warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "ToWhse", "To Warehouse", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "FromBranch", "From Branch", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "ToBranch", "To Branch", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("IBTFRH", "Status", "Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "CreatedBy", "Created By", 30, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "GoodReceiptEntry", "Goods Receipt Entry", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "GoodIssueEntry", "Goods Issue Entry", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "Status", "Status", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "FBranchN", "From Branch Name", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "TBranchN", "To Branch Name", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "ReqDocEntry", "Transfer Request ", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "GoodReceiptNo", "Goods Receipt No ", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRH", "GoodIssueNo", "Goods Issue No ", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRRH", "Focus", "Focus", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region InterBranchTransfer Line
                CreateTable("IBTFRC", "Inter Branch trasn Line", SAPbobsCOM.BoUTBTableType.bott_DocumentLines);
                AddColumn("IBTFRC", "ItemCode", "Item Code", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "Qty", "Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "OpnQty", "Open Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                //AddColumn("DELCL", "Disc", "Discount", 11, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Percentage, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "InterComCode", "Inter Company Code", 20, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "Remarks", "Remarks", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "ToWhseStatus", "To Ware House Status", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "LineId", "Sr. No", 11, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "BatchStatus", "Batch Status", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "BatchEntry", "Batch Entry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("IBTFRC", "LineNum", "Line Num", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region Batch Selection 
                CreateTable("BATCHS", "Batch Selection", SAPbobsCOM.BoUTBTableType.bott_MasterData);
                AddColumn("BATCHS", "ItemCode", "Item Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "ItemName", "Item Name", 100, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "WhsCode", "Warehouse Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "BatchNum", "Batch Number", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "BatchQty", "Batch Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "TotalStock", "Total Stock", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "BinCode", "Bin Code", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "ExpDate", "Exp Date", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "MnfDate", "Mnf Date", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "IntDocEntry", "Int DocEntry", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "LineNum", "Line Num", 10, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "ReqQty", "Required Quantity", 10, SAPbobsCOM.BoFieldTypes.db_Float, SAPbobsCOM.BoFldSubTypes.st_Quantity, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "Focus", "Focus", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                AddColumn("BATCHS", "BatchPost", "Batch Posted", 15, SAPbobsCOM.BoFieldTypes.db_Alpha, SAPbobsCOM.BoFldSubTypes.st_None, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO);
                #endregion

                #region CREATE UDO BATCH SELECTION
                //CreateUDO("IntBatchS", "Inter Branch - Batch Selection", "BATCHS", SAPbobsCOM.BoUDOObjType.boud_MasterData, oFindColumns, oChildTables);
                #endregion


                if ( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
 
				oProgressBar.Stop();
			}
			catch(Exception ex)
			{
				oProgressBar.Stop();

				if( Utilities.Application.Company.InTransaction )
					Utilities.Application.Company.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oProgressBar);
				oProgressBar = null;
				GC.Collect();
			}
		}

        #endregion

		#region DATABASE CREATION FUNCTIONS

		#region CREATE USER DEFINED TABLE
		
		private static void CreateTable(string oTable, string oDescription, SAPbobsCOM.BoUTBTableType oType)
		{
            GC.Collect();
            SAPbobsCOM.UserTablesMD oUserTable = null;

			try
			{
				oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
				
				if( !oUserTable.GetByKey(oTable) )
				{
					oUserTable.TableName        = oTable;
					oUserTable.TableDescription = oDescription;
					oUserTable.TableType        = oType;

					if( oUserTable.Add() != 0 )	
                    
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
																
				}
			}
			catch(Exception ex)
			{
                throw ex;
			}			
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
				oUserTable = null;
				GC.Collect();
			}
		}

        private static void UpdateTable(string oTable, SAPbobsCOM.BoUTBTableType oType)
        {
            SAPbobsCOM.UserTablesMD oUserTable = null;

            try
            {
                oUserTable = (SAPbobsCOM.UserTablesMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables);
                if (oUserTable.GetByKey(oTable) != null)
                {
                    if (oUserTable.TableType != SAPbobsCOM.BoUTBTableType.bott_MasterData)
                    {
                        //oUserTable.TableType = SAPbobsCOM.BoUTBTableType.bott_MasterData;
                        if (oUserTable.Remove() != 0)

                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable);
                oUserTable = null;
                GC.Collect();
            }
        }
		#endregion

		#region ADD COLUMN
		internal  static void AddColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserFieldsMD oUserField = null;
            int Fieldid;
			try
			{

                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tNO)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);

                    oUserField.TableName = oTable;
                    oUserField.Name = oName;
                    oUserField.Description = oDescription;
                    oUserField.Type = oType;
                    oUserField.SubType = oSubType;

                    if (oType != SAPbobsCOM.BoFieldTypes.db_Float)
                        oUserField.EditSize = oSize;
                    else
                        oUserField.EditSize = 0;

                    
                    oUserField.Mandatory = oMandatory;

                    if (oTable == "OINV" && oName == "PReqSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "PQuotSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }

                    if (oTable == "OINV" && oName == "POSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "POLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }

                   
                    if (oTable == "OINV" && oName == "GRPOSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "GRPOLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "APSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "APLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "APCRSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "APCRLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "GRReqSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "GRReqLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }

                    if (oTable == "OINV" && oName == "GRSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "GRLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "InvGRSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "InvGRLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "InvGISync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "InvGILineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "QuotSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "QuotLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "SOSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "SOLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "DelSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "DelLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "SalRetSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "SalRetReqSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "SalRetLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "SalInvSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "SalInvLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "ARCRSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "INV1" && oName == "ARCRLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "ORCT" && oName == "InPymntSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OVPM" && oName == "OutPymntSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OJDT" && oName == "JESync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINS" && oName == "CECSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINS" && oName == "CECUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OSCL" && oName == "SCSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OSCL" && oName == "SCUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "SCL4" && oName == "SCLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OOPR" && oName == "SOpprSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OIPF" && oName == "LandCostSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OPR1" && oName == "SOPLineSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OCRD" && oName == "BPSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OCRD" && oName == "BPUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OHEM" && oName == "EMPSync")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OHEM" && oName == "EMPUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";

                    }
                    if (oTable == "OINV" && oName == "PReqUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "PQuotUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "POUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "GRPOUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "APUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "APCRUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "GRReqUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "GRUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "QuotUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "SOUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "DelUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "SalRetReqUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "SalRetUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "SalInvUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }
                    if (oTable == "OINV" && oName == "ARCRUpdate")
                    {
                        oUserField.ValidValues.Value = "Yes";
                        oUserField.ValidValues.Description = "Yes";
                        oUserField.ValidValues.Add();
                        oUserField.ValidValues.Value = "No";
                        oUserField.ValidValues.Description = "No";
                        oUserField.ValidValues.Add();
                        oUserField.DefaultValue = "No";
                    }

                    if (oUserField.Add() != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                }
                
			}
			catch( Exception ex)
			{
                throw ex;
			}
			finally
			{
                if (oUserField!=null)
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
				oUserField = null;
				GC.Collect();
			}
		}

        internal static void UpdateColumn(string oTable, string oName, string oDescription, int oSize, SAPbobsCOM.BoFieldTypes oType, SAPbobsCOM.BoFldSubTypes oSubType, SAPbobsCOM.BoYesNoEnum oMandatory, SAPbobsCOM.BoYesNoEnum oIsSystemTable)
        {
            SAPbobsCOM.UserFieldsMD oUserField = null;
            string oTable1=oTable ;
            int Fieldid=0;
            try
            {
                
                if (ColumnExists(oTable, oName, oIsSystemTable,out  Fieldid) == SAPbobsCOM.BoYesNoEnum.tYES)
                {
                    oUserField = (SAPbobsCOM.UserFieldsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    if(oTable!="OCRD") 
                    oTable1 = @"@" + oTable;

                    oUserField.GetByKey(oTable1, Fieldid);
                    if (oTable == "B1F_GNDR" && (oName == "B1FINACT"))
                    {
                        if (oUserField.DefaultValue != "N")
                        {
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }
                    }
                    else if ((oTable == "OCRD" && oName == "Season") || (oTable == "B1F_ITEMSEAS" && oName == "B1FSEASD") || (oTable == "B1F_ITEMCAPS" && (oName == "B1FCAPSD" ||oName =="B1FDTST" || oName == "B1FDTCP")))
                    {
                        if (oUserField.Remove()  != 0)
                            throw new Exception(Utilities.Application.Company.GetLastErrorDescription());

                    }
                    else if (oTable == "OCRD" && (oName == "B1FBUYINP" || oName == "B1FBUYINPN"))
                    {
                        if (oUserField.ValidValues.Value == "")
                        {
                            oUserField.ValidValues.Value = "N";
                            oUserField.ValidValues.Description = "No";
                            oUserField.ValidValues.Add();
                            oUserField.ValidValues.Value = "Y";
                            oUserField.ValidValues.Description = "Yes";
                            oUserField.ValidValues.Add();
                            oUserField.DefaultValue = "N";
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                    else
                    {
                        if((oTable == "B1F_FSPM" && oName=="B1FSKUNUM") || (oTable == "B1F_FSPM" && oName=="B1FATLEN")|| (oTable == "B1F_FSAC" && oName=="B1FSEQ")||(oTable == "B1F_FSIM" && oName=="B1FSKULN")||(oTable == "B1F_VAT1" && oName=="B1FGRPSQ")||(oTable == "B1F_VAT2" && oName=="B1FGRPSQ"))
                        {
                            if (oUserField.EditSize  < oSize)
                            {
                                oUserField.EditSize = oSize;
                                if (oUserField.Update() != 0)
                                    throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                            }
                        }
                        if (oUserField.Description != oDescription)
                        {
                            oUserField.Description = oDescription;
                            if (oUserField.Update() != 0)
                                throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
                        }

                    }
                }
              

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                if (oUserField != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserField);
                oUserField = null;
                GC.Collect();
            }
        }

        #endregion

        //#region IS COLUMN EXISTS
        //private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable,out int FieldId)
        //{
        //    SAPbobsCOM.Recordset oRSColumn = null;
        //    string               oSQL      = string.Empty;
        //    FieldId = 0;
        //    try
        //    {
        //        if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
        //            oTable = @"@" + oTable;

        //        oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
        //        Utilities.ExecuteSQL(ref oRSColumn, oSQL);

        //        if ((int)oRSColumn.Fields.Item(0).Value == 0)
        //            return SAPbobsCOM.BoYesNoEnum.tNO;

        //        else
        //        {
        //            FieldId = (int)oRSColumn.Fields.Item(1).Value;
        //            return SAPbobsCOM.BoYesNoEnum.tYES;
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        throw ex;
        //    }
        //    finally
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
        //        oRSColumn = null;
        //        GC.Collect();
        //    }
        //}
        //#endregion

        #region IS COLUMN EXISTS
        private static SAPbobsCOM.BoYesNoEnum ColumnExists(string oTable, string oColumn, SAPbobsCOM.BoYesNoEnum oIsSystemTable, out int FieldId)
        {
            SAPbobsCOM.Recordset oRSColumn = null;
            string oSQL = string.Empty;
            FieldId = 0;
            try
            {
                if (oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO)
                    oTable = @"@" + oTable;

                //oSQL = "Select Count(*),FieldID From CUFD Where TableID = '" + oTable + "' And AliasID = '" + oColumn + "' group by FieldID";
                oSQL = @"Select " +
                        @"Count(*) "
                        + " , "
                        + @"""FieldID"""
                        + "From CUFD Where "
                        + @"""TableID"""
                         + "="
                        + "'" + oTable.Trim() + "'"
                        + " And "
                        + @"""AliasID"""
                        + "="
                        + "'" + oColumn.Trim() + "'"
                        + " group by "
                        + @"""FieldID""".Trim();


                Utilities.ExecuteSQL(ref oRSColumn, oSQL);

                if ((int)oRSColumn.Fields.Item(0).Value == 0)
                    return SAPbobsCOM.BoYesNoEnum.tNO;

                else
                {
                    FieldId = (int)oRSColumn.Fields.Item(1).Value;
                    return SAPbobsCOM.BoYesNoEnum.tYES;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
                oRSColumn = null;
                GC.Collect();
            }
        }
        #endregion


		#region CREATE UDO
		private static void CreateUDO(string oUniqueID, string oDescription, string oTable, SAPbobsCOM.BoUDOObjType oType, string[] oFindColumns, string[] oChildTables )
		{
			SAPbobsCOM.UserObjectsMD oUserObject = null;
			int                      oCount;

			try
			{
				oUserObject = (SAPbobsCOM.UserObjectsMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);

				if( !oUserObject.GetByKey(oUniqueID) )
				{
					oUserObject.Code        = oUniqueID;
					oUserObject.Name        = oDescription;
					oUserObject.ObjectType  = oType;
					oUserObject.TableName   = oTable;

					oUserObject.ManageSeries         = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanCancel            = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanClose             = SAPbobsCOM.BoYesNoEnum.tYES;
					oUserObject.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;
					oUserObject.CanDelete            = SAPbobsCOM.BoYesNoEnum.tNO;

					if( oFindColumns != null )
					{
						if( oFindColumns.Length > 0 )
						{
							oUserObject.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;

							for(oCount = 0; oCount <= oFindColumns.Length-1; oCount++)
							{
								if( oFindColumns[oCount] != null )
								{
									oUserObject.FindColumns.ColumnAlias = oFindColumns[oCount];
									oUserObject.FindColumns.Add();
								}

								oFindColumns[oCount] = null;
							}
						}
					}

					if( oChildTables != null )
					{
						if( oChildTables.Length > 0 )
						{
							for(oCount = 0; oCount <= oChildTables.Length-1; oCount++)
							{
								if( oChildTables[oCount] != null )
								{
									oUserObject.ChildTables.TableName = oChildTables[oCount];

									if( oCount != oChildTables.Length - 1 )
										oUserObject.ChildTables.Add();
								}

								oChildTables[oCount] = null;
							}
						}
					}

					if( oUserObject.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObject);
				oUserObject = null;
				GC.Collect();
			}
		}
		#endregion

        public static void CreateUserObject(string CodeID, string Name, string TableName, string Child) //used for registration of user defined table
        {
           
            int lRetCode = 0;
            string sErrMsg = null;
            SAPbobsCOM.UserObjectsMD oUserObjectMD = null;
            try
            {
                if (oUserObjectMD == null)
                    oUserObjectMD = ((SAPbobsCOM.UserObjectsMD)(Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)));

                if (oUserObjectMD.GetByKey(CodeID) == true)
                    return; //vijay 170708

                oUserObjectMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;


                oUserObjectMD.Code = CodeID;
                oUserObjectMD.Name = Name;
                oUserObjectMD.ObjectType = SAPbobsCOM.BoUDOObjType.boud_MasterData;
                oUserObjectMD.TableName = TableName;

                if (Child != "")
                {
                    oUserObjectMD.ChildTables.TableName = Child;
                    //oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES;
                }
                lRetCode = oUserObjectMD.Add();

                // check for errors in the process
                if (lRetCode != 0)
                    if (lRetCode == -1)
                    { }
                    else
                    { Utilities.Application.Company.GetLastError(out lRetCode, out sErrMsg); }
                else
                { }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
                oUserObjectMD = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

		#region IS KEY EXISTS
		private static SAPbobsCOM.BoYesNoEnum KeyExists(string oTable,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.Recordset oRSColumn = null;
			string               oSQL      = string.Empty;
			
			try
			{		
				if( oIsSystemTable == SAPbobsCOM.BoYesNoEnum.tNO )
					oTable = @"@" + oTable;
				oSQL = "Select Count(*) From OUKD Where TableName = '" + oTable + "' AND KeyName = '" + oKeyName + "' ";
				Utilities.ExecuteSQL(ref oRSColumn, oSQL);
            
				if( (int)oRSColumn.Fields.Item(0).Value == 0 )
					return SAPbobsCOM.BoYesNoEnum.tNO;

				else
					return SAPbobsCOM.BoYesNoEnum.tYES;
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSColumn);
				oRSColumn = null;
				GC.Collect();
			}
		}
		#endregion

		#region ADD KEYS
		private static void AddKey(string oTable, string[] oColumnName,string oKeyName,SAPbobsCOM.BoYesNoEnum oIsSystemTable)
		{
			SAPbobsCOM.UserKeysMD oUserKey = null;
			int                   oCount;
			try
			{				
				oUserKey = (SAPbobsCOM.UserKeysMD)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys);
				if(KeyExists(oTable,oKeyName,oIsSystemTable) == SAPbobsCOM.BoYesNoEnum.tNO)
				{
					oUserKey.TableName = oTable ;
					oUserKey.KeyName   = oKeyName ;
					if( oColumnName != null )
					{
						if( oColumnName.Length > 0 )
						{						
							for(oCount = 0; oCount <= oColumnName.Length-1; oCount++)
							{
								if( oColumnName[oCount] != null )
								{
									oUserKey.Elements.ColumnAlias = oColumnName[oCount];
									oUserKey.Elements.Add();
								}

								oColumnName[oCount] = null;
							}
						}
					}							
					oUserKey.Unique = SAPbobsCOM.BoYesNoEnum.tYES;					
					if( oUserKey.Add() != 0 )
						throw new Exception(Utilities.Application.Company.GetLastErrorDescription());
				}
			}
			catch(Exception ex)
			{
				throw ex ;
			}
			finally
			{
				System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKey);
				oUserKey = null;
				GC.Collect();
			}
		}
		#endregion

		#endregion

	}
}
