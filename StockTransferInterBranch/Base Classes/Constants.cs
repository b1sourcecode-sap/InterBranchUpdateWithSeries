using System;
using System.Data.SqlClient;
using System.Data;
using System.Text.RegularExpressions;
using CrystalDecisions.CrystalReports.Engine;  //USE WHILE WOKING WITH CRYSTAL REPORTS 
using CrystalDecisions.Shared; //USE WHILE WOKING WITH CRYSTAL REPORTS 
using System.Windows.Forms;

namespace StockTransferInterBranch
{
	/// <summary>
	/// Summary description for Constants.
	/// </summary>
	public class Constants
    {
        #region GLOBAL VARIABLE DECLARATIONS
        public static SqlConnection                     gobjSQLCon;
        public static string                            gConnStr; 
		public static string                            gSERVER;
		public static string                            gUSER_ID;
		public static string                            gUSER_PASSWORD;
        public static string                            gLicenseServer;

		public static Int32                             gTimeOut = 100000000;
        //public static ReportDocument                    goMainReportDoc = null; //USE WHILE WOKING WITH CRYSTAL REPORTS 
		public static SAPbouiCOM.UserDataSource         goUSRProperties;
		public static string                            gReportPath ;
		public static string                            gReportXMLPath ;
       // public static CRDisplay.frmReport               gReportForm;  //USE WHILE WOKING WITH CRYSTAL REPORTS 
		public static string                            gcfl_Text;
		public static string                            gcfl_Btn;
        public static string                             gHideMessage;
        internal static bool                            gShopActivated = false;
   

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        private Constants()
		{
        }
        ~Constants()
        {}
        #endregion

        #region MENUS
        public struct Menus
        {
            public const string MENUS_ADD = "Menus.xml";
            public const string REMOVE_MENUS = "RemoveMenus.xml";
        }
        #endregion	

		#region CONSTANTS FOR SYSTEM FORMS
		// Constants for System Forms
		public struct System_Forms
		{
            public const int ITEM_MASTER = 150;
            public const int ITEM_MASTER_udfPannel = -150;
            public const int SALES_ORDER = 139;
            public const int SALES_QUOTATION = 149;
            public const int AR_INVOICE = 133;
            public const int AR_CREDITNOTE = 179;
            public const int PURCHASE_ORDER = 142;
            public const int GRPO = 143;
            public const int BP_Master = 134;
            public const int BP_Master_UdfPannel = -134;
            public const int Delivery = 140;
            public const int Outgoing_Payment = 426;
            public const int AP_INVOICE = 141;
            public const int AP_CREDIT_MEMO = 181;
            public const int GOODS_ISSUE = 720;
            public const int INVENTORY_TRANSFER_REQUEST = 1250000940;
        }
		#endregion

		#region CONSTANTS FOR SYSTEM MENUS

		// Constants for System Menu's
		public struct System_Menus
		{			
			public const string mnu_FIND                        = "1281";
			public const string mnu_ADD                         = "1282";
			public const string mnu_NEXT                        = "1288";
			public const string mnu_PREVIOUS                    = "1289";
			public const string mnu_FIRST                       = "1290";
			public const string mnu_LAST                        = "1291";
			public const string mnu_ADD_ROW                     = "1292";
			public const string mnu_REMOVE_RECORD               = "1283";
            public const string mnu_DELETE_ROW                  = "1293";
			public const string mnu_SALES_ORDER                 = "2050";
			public const string mnu_OUTGOING_PAYMENTS           = "2818" ;
			public const string mnu_ROW_DETAILS                 = "5889" ;
			public const string mnu_GROSS_PROFIT                = "5891" ;
			public const string mnu_GL_ACCOUNT_DETERMINATION    = "8199" ;
			public const string mnu_GENERAL_SETTINGS            = "8194";
            public const string mnu_SORT                        = "4869";
            public const string mnu_FILTER                      = "4870";
            public const string mnu_PASTE                       = "773";
            public const string mnu_Duplicate                   = "1287";
            public const string mnu_Cancelled                   = "1284";
            public const string mnu_REFRESH                     = "1304";

        }
		#endregion				
 
		#region CONSTANTS FOR USER FORMS

		// Constants for XML Files
		public struct Forms
		{
         //   public const string FASHION_PARAMERTS                               = "FashionParameters.xml";  
           // public const string Barcode_Scanning = "Barcode Scanning.xml";

            public const string DeliveryCreation = "DeliveryCreation.xml";
            public const string InterBranchTransferRequest = "InterBranchTransReq.xml";
            public const string InterBranchTransfer = "InterBranchTrans.xml";
            public const string ApprovalTemplate = "ApprovalScreen.xml";
            public const string BatchSelection = "BatchSelection.xml";

        }
		#endregion 	
			
		#region CONSTANTS FOR USER FORMS MENUID

		// Constants for User defined Menu's
		public struct User_Menus
		{
         //   public const string MENU_FASHION_PARAMETER                   = "B1SOL_01_11";
            //public const string MENU_Barcode_Scanning = "BarCode";
            public const string MENU_DeliveryCreation = "DelCreation";
            public const string MENU_InterBranchTransferReq = "IntTFRR";
            public const string MENU_InterBranchTransfer = "IntTFR";
            public const string MENU_LGApproval = "IntApproval";
            public const string MENU_BatchSelection = "BatchSelection";

        }
		#endregion		

		#region CONSTANTS FOR KEYS
		public struct Keys
		{
			public const int TAB = 9;
			public const int BACKSPACE = 8;
			public const int DELELTE = 36;
			public const int ARROW_DOWN = 40;
			public const int ARROW_UP = 38;
		}
		#endregion

		#region CONSTANTS FOR VALIDATION RESULT
		public enum ValidationResult
		{
			CORRECT = 1,
			INCORRECT = 2
		};
		#endregion

        #region CONSTANTS FOR PropertyCondition
        internal struct PropertyCondition
        {
           internal const string OR = "OR";
           internal const string AND = "AND";
        };
        #endregion

        #region CFL Event
        public struct CFL_Event
        {
            public string ItemUID;
            public string ColUID;
            public int Row;
            public System.Data.DataTable oCFL_SelectedResult;
        }
        #endregion

        #region Enum for Item properties 
        public enum ItemPropertyBaseForm
        {
            ShopStockPattern = 1,
            clsSelcCrit_ShanfariLineSheet = 2,
            ShopIssues = 3,
            EditAndGenerateShopReplenishment
        };
        #endregion

        #region Shop Stock Patterns Maintain
        public enum ShopStockPatternsMaintain
        {
            Minimum = 1,
            Maximum = 2 ,
            Uplift = 3,
            Prioity = 4,
            Cover = 5,
        };
        #endregion

        #region Shop Issues
        public enum ShopIssues
        {
            Minimum = 1,
            Maximum = 2,
        };
        #endregion
	}
}
