using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using Microsoft.ApplicationBlocks.Data;
using System.Text;
using System.Xml;
using System.Collections.Generic;
using System.Collections;
using System.Threading;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Globalization;
using SAPbobsCOM;

namespace StockTransferInterBranch
{
	/// <summary>
	/// Summary description for Utilities.
	/// </summary>
	public class Utilities
    {
        #region CLASS LEVEL VARIABLE DECLARATION
        private static EventListener oApplication;
		private static int FormCounter;
		private static SAPbobsCOM.Recordset oRS1;
        private static bool _HideSystemMessege;
        private static int iHideCount = 1;
        public static string SelectedProperty = "";
         public static bool _AvailableDateCopy = false;

        //========== Added By Sundeep =====================
        [DllImport("user32.dll")]
         private static extern IntPtr GetForegroundWindow();
         public static OpenFileDialog OFD = new OpenFileDialog();
         private static FolderBrowserDialog FBD = new FolderBrowserDialog();
         private static SaveFileDialog SFD = new SaveFileDialog();
         public static string _strFileName = string.Empty, _AttachmentPath = string.Empty;
        //==================================================

        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public Utilities()
		{
        }
        ~Utilities()
        { }
        #endregion

        #region PROPERTY APPLICATION
        public static EventListener Application
		{
			get { return oApplication; }

			set { oApplication = value; }
		}		
		#endregion		 

        #region GET MAX COLUMN VALUE
        public static string getMaxColumnValue(string oTable, string oColumn)
        {
            SAPbobsCOM.Recordset oRS = null;
            string oSQL = string.Empty;
            string oCode = string.Empty;
            int oMaxCode;

            oSQL = "SELECT MAX(CAST(\"" + oColumn + "\" AS Numeric)) FROM \"" + oTable + "\"";
            ExecuteSQL(ref oRS, oSQL);

            if (Convert.ToString(oRS.Fields.Item(0).Value).Length > 0)
                oMaxCode = int.Parse(oRS.Fields.Item(0).Value.ToString()) + 1;
            else
                oMaxCode = 1;

            oCode = oMaxCode.ToString("00000000");

            return oCode;
        }
        #endregion

        #region GET MAX COLUMN VALUE
        public static string getMaxColumnValueNum(string oTable, string oColumn)
        {
            SAPbobsCOM.Recordset oRS = null;
            string oSQL = string.Empty;
            string oCode = string.Empty;
            int oMaxCode;

            oSQL = "SELECT MAX(CAST(\"" + oColumn + "\" AS Numeric)) FROM \"" + oTable + "\"";
            ExecuteSQL(ref oRS, oSQL);

            if (Convert.ToString(oRS.Fields.Item(0).Value).Length > 0)
                oMaxCode = int.Parse(oRS.Fields.Item(0).Value.ToString()) + 1;
            else
                oMaxCode = 1;

            oCode = oMaxCode.ToString("");

            return oCode;
        }
        #endregion		
		
		#region LOAD XML FILES
		public static void LoadForm(ref Base oObject, string oFile)
		{
			oObject.FormUID = LoadFromXML(oFile,true);
			oObject.Form = oApplication.SBO_Application.Forms.Item(oObject.FormUID);

			if ( !oApplication.Collection.ContainsKey(oObject.FormUID) )
			{
				oApplication.Collection.Add(oObject.FormUID,oObject);
			}
		}

		public static void LoadMenus(string oFileName)
		{
			LoadFromXML(oFileName, false);
			
		}

		private static string LoadFromXML( string oFileName, bool oIsForm ) 
		{        
			string                  oPath    = null; 
			string                  oFormUID = null;
			System.Xml.XmlDocument  oXmlDoc  = null;
			System.Xml.XmlNode      oXmlNode = null;
			System.Xml.XmlAttribute oAttri   = null;
        
			oXmlDoc = new System.Xml.XmlDocument();

            oPath = getApplicationPath() + @"\XML Files\" + oFileName;
			oXmlDoc.Load(oPath);

			if ( oIsForm )
			{
				oXmlNode = oXmlDoc.GetElementsByTagName("form").Item(0);
				oAttri = (System.Xml.XmlAttribute)oXmlNode.Attributes.GetNamedItem("uid");
				oAttri.Value = oAttri.Value.ToString().Trim() + FormCounter.ToString();
				oFormUID = oAttri.Value;
				FormCounter++;
			}
        
			string ostrXML = oXmlDoc.InnerXml.ToString();
			oApplication.SBO_Application.LoadBatchActions(ref ostrXML ); 

			return oFormUID;        
		} 
		#endregion
		
		#region GET APPLICATION PATH
		public static string getApplicationPath()
		{
			string oPath;

			  oPath = System.Windows.Forms.Application.StartupPath.Trim();
			//oPath = System.IO.Directory.GetParent(sPath).ToString(); 

			return oPath;
		}
		#endregion		

		#region EXECUTE QUERY
		public static void ExecuteSQL(ref SAPbobsCOM.Recordset oRecordSet, string oSql)
		{
			if( oRecordSet == null)
				oRecordSet = (SAPbobsCOM.Recordset)oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

			oRecordSet.DoQuery(oSql);
		}
        #endregion

        #region ExecuteScalar
        public static string ExecuteScalar(string sql)
        {
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(sql);
            if (!rs.EoF)
                return rs.Fields.Item(0).Value.ToString();
            return string.Empty;
        }
        #endregion

        #region SHOW MESSAGE
        public static void ShowErrorMessage( string oText )
		{
			Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
		}

		public static void ShowWarningMessage( string oText )
		{
			Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
		}

        public static void ShowSucessMessage(string oText)
        {
            Message(oText, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

		public static void Message(string oText, SAPbouiCOM.BoStatusBarMessageType oType)
		{
			if( oApplication!= null && oApplication.Company != null)
			{
				oApplication.SBO_Application.StatusBar.SetText(oText,SAPbouiCOM.BoMessageTime.bmt_Short, oType);
			}
			else
			{
                System.Windows.Forms.MessageBox.Show(oText,System.Reflection.Assembly.GetExecutingAssembly().GetName().Name.ToString(), System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error, System.Windows.Forms.MessageBoxDefaultButton.Button1, System.Windows.Forms.MessageBoxOptions.DefaultDesktopOnly);
			}
		}
		#endregion 

		#region TO DATE
		public static DateTime ToDate(string oDate)
		{
			oDate = oDate.Trim().Insert(4,"/").Insert(7,"/");
			return DateTime.Parse(oDate);			
		}
		#endregion
		
		#region GET CURRENT FISCAL START DATE
		public static string GetFiscalStart()
		{
			SAPbobsCOM.Recordset oRS  = null;
			DateTime             odtStartDate;
			string               oSQL = "Select FinancYear,dateadd(mm,12,FinancYear)-1 from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1";
			ExecuteSQL(ref oRS,oSQL);
			odtStartDate = Convert.ToDateTime(oRS.Fields.Item(0).Value);

			return odtStartDate.ToString("yyyyMMdd");
		}
		#endregion

		#region GET CURRENT FISCAL END DATE
		public static string GetFiscalEnd()
		{
			SAPbobsCOM.Recordset oRS  = null;
			DateTime             odtEndDate;
			string               oSQL = " Select FinancYear,dateadd(mm,12,FinancYear)-1 from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1";
			ExecuteSQL(ref oRS,oSQL);
			odtEndDate = Convert.ToDateTime(oRS.Fields.Item(1).Value);

			return odtEndDate.ToString("yyyyMMdd");
		}
		#endregion

		#region GET MID YEAR END DATE OF CURRENT FISCAL YEAR
		public static string GetMidYearEndDate()
		{
			SAPbobsCOM.Recordset oRS  = null;
			DateTime             odtMidYrDate;
			string               oSQL = " Select dateadd(mm,6,FinancYear-1) from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1";
			ExecuteSQL(ref oRS,oSQL);
			odtMidYrDate = Convert.ToDateTime(oRS.Fields.Item(0).Value);

			return odtMidYrDate.ToString("yyyyMMdd");
		}
		#endregion

		#region GET MID YEAR START DATE OF CURRENT FISCAL YEAR
		public static string GetMidYearStartDate()
		{
			SAPbobsCOM.Recordset oRS  = null;
			DateTime             odtMidYrDate;
			string               oSQL = " Select dateadd(mm,6,FinancYear) from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1";
			ExecuteSQL(ref oRS,oSQL);
			odtMidYrDate = Convert.ToDateTime(oRS.Fields.Item(0).Value);

			return odtMidYrDate.ToString("yyyyMMdd");
		}
		#endregion

		#region GET PREVIOUS FISCAL START DATE
		public static string GetPrevFiscalStart()
		{
			SAPbobsCOM.Recordset oRS  = null;
			DateTime             odtPrevStartDt;
			string               oSQL = " select max(FinancYear) from OACP where "
				                       +" FinancYear <(Select FinancYear from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1)";
			ExecuteSQL(ref oRS,oSQL);
			odtPrevStartDt = Convert.ToDateTime(oRS.Fields.Item(0).Value);

			return odtPrevStartDt.ToString("yyyyMMdd");
		}
		#endregion

		#region GET PREVIOUS FISCAL END DATE
		public static string GetPrevFiscalEnd()
		{
			SAPbobsCOM.Recordset oRS = null;
			DateTime             odtPrevEndDt;
			string               oSQL = " Select FinancYear-1 from OACP where getdate() between FinancYear and dateadd(mm,12,FinancYear)-1";
			ExecuteSQL(ref oRS,oSQL);
			odtPrevEndDt = Convert.ToDateTime(oRS.Fields.Item(0).Value);

			return odtPrevEndDt.ToString("yyyyMMdd");
		}
		#endregion

		#region Get CNF PRICELIST
		public static string GetPriceList()
		{
			SAPbobsCOM.Recordset oRS = null;
            string oSQL              = "select U_UserId from [@AG_CFUSER] where Code = '" + Utilities.Application.Company.UserName + "'";
			ExecuteSQL(ref oRS,oSQL);

			return oRS.Fields.Item(0).Value.ToString();
		}
		#endregion

		#region GET CALCULATION METHOD
		public static string GetCalcMethod()
		{
			SAPbobsCOM.Recordset oRS = null;
            string               oSQL = "select U_UserName from [@AG_CFUSER] where Code = '" + Utilities.Application.Company.UserName + "'";
			ExecuteSQL(ref oRS,oSQL);

            return oRS.Fields.Item(0).Value.ToString();
		}
		#endregion

		#region IS LOGGED IN USER A CNF AGENT
		public static bool isCFA()
		{
			SAPbobsCOM.Recordset oRS = null;
            string               oSQL = "select Code from [@AG_CFUSER] where Code =  '" + Utilities.Application.Company.UserName + "'";
			ExecuteSQL(ref oRS,oSQL);

			if (oRS.RecordCount == 0)
				return false;
			else
				return true;
			
		}
		#endregion
		
		#region FILL COMBO
		public static void FillCombo(ref SAPbouiCOM.ComboBox oCombo, string oSQL )
		{
                            SAPbobsCOM.Recordset oRS = null;

            try
            {
                ExecuteSQL(ref oRS, oSQL);
                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                oRS.MoveFirst();
                while (!oRS.EoF)
                {
                    oCombo.ValidValues.Add(oRS.Fields.Item(0).Value.ToString(), oRS.Fields.Item(1).Value.ToString());
                    oRS.MoveNext();
                }
            }
            finally
            {
                if (oRS != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRS);
                    oRS = null;
                }
            }
		}

        public static void FillCombo(ref SAPbouiCOM.ComboBox oCombo, DataTable dt)
        {
            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }
            foreach (DataRow  dr in dt.Rows)
            {
                oCombo.ValidValues.Add(dr[0].ToString(), dr[1].ToString());
            }
        }

		#endregion				

		#region GET ITEM SELLING PRICE
		public static double getItemsSellingPrice(string oCardCode, string oItemCode)
		{
			double                      oSellingPrice = 0;
			SAPbobsCOM.BusinessPartners oBP           = (SAPbobsCOM.BusinessPartners)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);
			SAPbobsCOM.Items            oItems        =(SAPbobsCOM.Items)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);

			if( oBP.GetByKey( oCardCode ) && oItems.GetByKey( oItemCode ) )
			{
				oItems.PriceList.SetCurrentLine( oBP.PriceListNum - 1 );
				oSellingPrice = oItems.PriceList.Price;				
			}
			
			return oSellingPrice;
		}
		#endregion

		#region ADD CHOOSE FROM LIST
		public static void AddChooseFromList(string oFormUID, string oCFL_Text, string oCFL_Button,  string oAliasName, string oCondVal, SAPbouiCOM.BoConditionOperation oOperation)
		{
			SAPbouiCOM.ChooseFromListCollection     oCFLs;
			SAPbouiCOM.Conditions                   oCons;
			SAPbouiCOM.Condition                    oCon;
			SAPbouiCOM.ChooseFromList               oCFL;
			SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;

			try
			{				
				oCFLs = oApplication.SBO_Application.Forms.Item(oFormUID).ChooseFromLists;
				oCFLCreationParams = ( (SAPbouiCOM.ChooseFromListCreationParams)(oApplication.SBO_Application.CreateObject( SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams ) ) ); 
				
				//' Adding 2 CFL, one for the button and one for the edit text.
				oCFLCreationParams.MultiSelection = false;
				oCFLCreationParams.ObjectType = (-1).ToString() ;
				oCFLCreationParams.UniqueID = oCFL_Text;
				oCFL = oCFLs.Add(oCFLCreationParams);

				//'Adding Conditions to CFL
				oCons = oCFL.GetConditions();
				if (oAliasName != "")
				{
					oCon = oCons.Add();
					oCon.Alias = oAliasName;
					oCon.Operation = oOperation;
					oCon.CondVal = oCondVal;
					oCFL.SetConditions(oCons);
				}

				oCFLCreationParams.UniqueID = oCFL_Button;
				oCFL = oCFLs.Add(oCFLCreationParams);
			}
			catch(Exception ex)
			{
				throw ex;
			}
			finally
			{
			}

		}

        public static void AddChooseFromList(string oFormUID, string oCFL_Text, string oCFL_Button, SAPbouiCOM.BoLinkedObject oObjectType, string oAliasName, string oCondVal, SAPbouiCOM.BoConditionOperation oOperation)
        {
            SAPbouiCOM.ChooseFromListCollection oCFLs;
            SAPbouiCOM.Conditions oCons;
            SAPbouiCOM.Condition oCon;
            SAPbouiCOM.ChooseFromList oCFL;
            SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams;

            try
            {
                oCFLs = oApplication.SBO_Application.Forms.Item(oFormUID).ChooseFromLists;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(oApplication.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                //' Adding 2 CFL, one for the button and one for the edit text.
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = ((int)oObjectType).ToString();
                oCFLCreationParams.UniqueID = oCFL_Text;
                oCFL = oCFLs.Add(oCFLCreationParams);

                //'Adding Conditions to CFL
                oCons = oCFL.GetConditions();
                if (oAliasName != "")
                {
                    oCon = oCons.Add();
                    oCon.Alias = oAliasName;
                    oCon.Operation = oOperation;
                    oCon.CondVal = oCondVal;
                    oCFL.SetConditions(oCons);
                }

                oCFLCreationParams.UniqueID = oCFL_Button;
                oCFL = oCFLs.Add(oCFLCreationParams);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
            }

        }		

		#endregion 		

		#region FILL TERRITORY
		/// <summary>
		/// Funtion to fill all locations in 
		/// </summary>
		public static void FillTerritry()
		{
			
			SqlTransaction oTran = null;
			try
			{
				Cursor.Current = Cursors.WaitCursor;
				oTran = Constants.gobjSQLCon.BeginTransaction();

				string oSql  ="delete from Territories";
				SqlCommand ocmd = new SqlCommand(oSql,Constants.gobjSQLCon,oTran);
				int oresult=ocmd.ExecuteNonQuery();

				oSql= "select distinct x1.territryid as 'counid',x1.descript as 'coun',x2.territryid as 'zoneid',x2.descript as 'zone'" +
					",x3.territryid as 'stateid',x3.descript as 'state',x4.u_ttype as 'child' from locations x1 inner join locations x2 on " +
					"x1.territryid=x2.parent inner join locations x3 on x2.territryid=x3.parent inner join locations x4 on " +
					"x3.territryid=x4.parent where x1.parent = -1";

				ocmd.CommandText = oSql;
				ocmd.Transaction = oTran;
				SqlDataAdapter oad = new SqlDataAdapter(ocmd);
				DataTable odt1 = new DataTable();
				oad.Fill(odt1);
			
				string oStateid = "";
				for(int i=0;i<=odt1.Rows.Count - 1;i++)
				{
					if (oStateid != odt1.Rows[i]["stateid"].ToString())
					{
						oStateid = odt1.Rows[i]["stateid"].ToString();
						if(odt1.Rows[i]["child"].ToString() == "4")
						{
							oSql = "select distinct x2.territryid ,x2.descript ,x2.u_ttype as 'child' from locations x1 " +
								   "inner join locations x2 on x1.territryid=x2.parent inner join locations x3 on " + 
								   "x2.territryid=x3.parent where x1.territryid=" + odt1.Rows[i]["stateid"].ToString();

							ocmd.CommandText = oSql;
							ocmd.Transaction = oTran;
							oad = new SqlDataAdapter(ocmd);
							DataTable odt2= new DataTable();
							oad.Fill(odt2);
							for (int j=0 ;j<=odt2.Rows.Count -1;j++)
							{
								if(odt2.Rows[j]["child"].ToString() == "4")
								{
									oSql = "select x2.territryid ,x2.descript ,x3.territryid as 'territryid2'" +
										   ",x3.descript as 'descript2',x3.u_ttype as 'child' from locations x1 inner join locations x2 on " +
									       "x1.territryid=x2.parent inner join locations x3 on x2.territryid=x3.parent where x2.territryid=" 
										+ odt2.Rows[j]["territryid"].ToString();
							
									ocmd.CommandText = oSql;
									ocmd.Transaction = oTran;
									oad = new SqlDataAdapter(ocmd);
									DataTable odt3= new DataTable();
									oad.Fill(odt3);
									for(int k=0;k<=odt3.Rows.Count - 1;k++)
									{
										if(odt3.Rows[k]["child"].ToString() == "5")
										{
											oSql = "select x2.territryid as 'territryid1' ,x2.descript as 'descript1' ,x2.u_ttype as 'child'," + 
											       "x3.territryid as 'territryid2',x3.descript as 'descript2'" +
												   "from locations x1 inner join locations x2 on x1.territryid=x2.parent inner join locations " + 
												   "x3 on x2.territryid=x3.parent where x2.territryid=" + odt3.Rows[k]["territryid2"].ToString();
											ocmd.CommandText = oSql;
											ocmd.Transaction = oTran;
											oad = new SqlDataAdapter(ocmd);
											DataTable odt4= new DataTable();
											oad.Fill(odt4);

											for(int l=0;l <= odt4.Rows.Count-1;l++)
											{
												oSql = "Insert into Territories values(" + odt1.Rows[i]["counid"].ToString() + ",'" + 
													odt1.Rows[i]["coun"].ToString() + "'," + odt1.Rows[i]["zoneid"].ToString() + ",'" +
													odt1.Rows[i]["zone"].ToString() + "'," + odt1.Rows[i]["stateid"].ToString() + ",'" +
													odt1.Rows[i]["state"].ToString() + "'," + odt2.Rows[j]["territryid"].ToString() + ",'" +
													odt2.Rows[j]["descript"].ToString() + "'," + odt4.Rows[l]["territryid1"].ToString() + ",'" +
													odt4.Rows[l]["descript1"].ToString() + "'," + odt4.Rows[l]["territryid2"].ToString() + ",'" +
													odt4.Rows[l]["descript2"].ToString() + "')";
												
												ocmd.CommandText = oSql;
												ocmd.Transaction = oTran;
												ocmd.ExecuteNonQuery();
											}		

										}
										else if(odt3.Rows[k]["child"].ToString() == "6")
										{
											oSql = "Insert into Territories values(" + odt1.Rows[i]["counid"].ToString() + ",'" + 
												odt1.Rows[i]["coun"].ToString() + "'," + odt1.Rows[i]["zoneid"].ToString() + ",'" +
												odt1.Rows[i]["zone"].ToString() + "'," + odt1.Rows[i]["stateid"].ToString() + ",'" +
												odt1.Rows[i]["state"].ToString() + "'," + odt2.Rows[j]["territryid"].ToString() + ",'" +
												odt2.Rows[j]["descript"].ToString() + "',null,null," + odt3.Rows[k]["territryid2"].ToString() + ",'" +
												odt3.Rows[k]["descript2"].ToString() + "')";

											ocmd.CommandText = oSql;
											ocmd.Transaction = oTran;
											ocmd.ExecuteNonQuery();		
										}
									}
								}
								if(odt2.Rows[j]["child"].ToString() == "5")
								{
									oSql = "select x2.territryid ,x2.descript ,x2.u_ttype as 'child' from locations x1 inner join " +
										   "locations x2 on x1.territryid=x2.parent where x1.territryid=" + odt2.Rows[j]["territryid"].ToString();
							
									ocmd.CommandText = oSql;
									ocmd.Transaction = oTran;
									oad = new SqlDataAdapter(ocmd);
									DataTable odt3= new DataTable();
									oad.Fill(odt3);

									for(int k=0;k<=odt3.Rows.Count-1;k++)
									{
										oSql = "Insert into Territories values(" + odt1.Rows[i]["counid"].ToString() + ",'" + 
											odt1.Rows[i]["coun"].ToString() + "'," + odt1.Rows[i]["zoneid"].ToString() + ",'" +
											odt1.Rows[i]["zone"].ToString() + "'," + odt1.Rows[i]["stateid"].ToString() + ",'" +
											odt1.Rows[i]["state"].ToString() + "',Null,Null," + odt2.Rows[j]["territryid"].ToString() + ",'" +
											odt2.Rows[j]["descript"].ToString() + "'," + odt3.Rows[k]["territryid"].ToString() + ",'" +
											odt3.Rows[k]["descript"].ToString() + "')";

										ocmd.CommandText = oSql;
										ocmd.Transaction = oTran;
										ocmd.ExecuteNonQuery();
									}		
								}
							}
						}			
						else
						{
							if(odt1.Rows[i]["child"].ToString() == "5")
							{
								oSql = "select x2.territryid ,x2.descript ,x2.u_ttype as 'child' from locations x1 inner join " +
								       "locations x2 on x1.territryid=x2.parent where x2.territryid=" + odt1.Rows[i]["stateid"].ToString();
								ocmd.CommandText = oSql;
								ocmd.Transaction = oTran;
								oad = new SqlDataAdapter(ocmd);
								DataTable odt3= new DataTable();
								oad.Fill(odt3);

								for(int k=0;k<=odt3.Rows.Count-1;k++)
								{
									oSql = "select x2.territryid ,x2.descript ,x2.u_ttype as 'child' from locations x1 inner join " +
										   "locations x2 on x1.territryid=x2.parent where x1.territryid=" + odt3.Rows[k]["territryid"].ToString();

									ocmd.CommandText = oSql;
									ocmd.Transaction = oTran;
									oad = new SqlDataAdapter(ocmd);
									DataTable odt4= new DataTable();
									oad.Fill(odt4);
									for(int l=0;l<=odt4.Rows.Count - 1;l++)
									{
										oSql = "select x2.territryid ,x2.descript ,x2.u_ttype as 'child' from locations x1 inner join " +
											   "locations x2 on x1.territryid=x2.parent where x1.territryid=" + odt4.Rows[l]["territryid"].ToString();
                                        
										ocmd.CommandText = oSql;
										ocmd.Transaction = oTran;
										oad = new SqlDataAdapter(ocmd);
										DataTable odt5= new DataTable();
										oad.Fill(odt5);

										for(int m=0;m<=odt5.Rows.Count - 1;m++)
										{
											oSql = "Insert into Territories values(" + odt1.Rows[i]["counid"].ToString() + ",'" + 
												odt1.Rows[i]["coun"].ToString() + "'," + odt1.Rows[i]["zoneid"].ToString() + ",'" +
												odt1.Rows[i]["zone"].ToString() + "'," + odt1.Rows[i]["stateid"].ToString() + ",'" +
												odt1.Rows[i]["state"].ToString() + "',Null,Null," + odt4.Rows[l]["territryid"].ToString() + ",'" +
												odt4.Rows[l]["descript"].ToString() + "'," + odt5.Rows[m]["territryid"].ToString() + ",'" +
												odt5.Rows[m]["descript"].ToString() + "')";

											ocmd.CommandText = oSql;
											ocmd.Transaction = oTran;
											ocmd.ExecuteNonQuery();
										}
									}
								}		
							}
						}
					}
				}
				oTran.Commit();
			}
			catch(Exception	 ex)
			{
				if (oTran != null) oTran.Rollback();
				throw ex;
			}
			finally
			{
				Cursor.Current = Cursors.Default;
				if (oTran != null) oTran.Dispose();
			}
		}
		#endregion 

		#region EXECUTE ADO SQL
		public static DataTable ExecuteADOSql(string oSQL)
		{
			if (Constants.gobjSQLCon.State == ConnectionState.Closed) Constants.gobjSQLCon.Open();
			Constants.gobjSQLCon.ChangeDatabase(Utilities.Application.Company.CompanyDB);
		
			SqlCommand     oCmd = new SqlCommand(oSQL,Constants.gobjSQLCon);
			SqlDataAdapter oAd = new SqlDataAdapter(oCmd);
			DataTable oDt = new DataTable();
			oAd.Fill(oDt);
			return oDt;
		}
		#endregion

        #region EXECUTE ADO SQL
        public static SAPbobsCOM.Recordset ExecuteSQL(string oSql)
        {
            SAPbobsCOM.Recordset oRecordSet = null;
            if (oRecordSet == null)
                oRecordSet = (SAPbobsCOM.Recordset)oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            oRecordSet.DoQuery(oSql);
            return oRecordSet;
        }
        #endregion

        #region TODAY'S DATE
        //public static DateTime TodayDate()
        public static string TodayDate()
        {
            SAPbobsCOM.Recordset oRS = null;
            //string oSQL = "select  getdate()";
            string oSQL = "select CURRENT_DATE from dummy";
            ExecuteSQL(ref oRS, oSQL);
            if (oRS.RecordCount > 0)
            {
                //return Convert.ToDateTime(oRS.Fields.Item(0).Value);
                return oRS.Fields.Item(0).Value.ToString();
            }
            //return DateTime.Today;

            return DateTime.Today.ToString();
        }
        #endregion

        #region TODAY'S DATE
        public static string TodayDateISOFormat()
        {
            SAPbobsCOM.Recordset oRS = null;
            string oSQL = "select  Convert(varchar(10),getdate(),112) ";
            ExecuteSQL(ref oRS, oSQL);
            if (oRS.RecordCount > 0)
            {
                return oRS.Fields.Item(0).Value.ToString();
            }
            return DateTime.Today.ToShortDateString();
        }
        #endregion

		#region	CHECK DUPLICATE VALUE IN MATRIX

		public static bool CheckDuplicateMatrixValue(SAPbouiCOM.Matrix oMatrix,string oColumnNo)
		{
			bool oFlag = false;
			for(int i=1;i<=oMatrix.RowCount;i++)
			{ 
				string oNewValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(oColumnNo).Cells.Item(i).Specific).String;
				for(int j=1;j<oMatrix.RowCount;j++)
				{
					if (i!=j)
					{
						string oOldValue = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(oColumnNo).Cells.Item(j).Specific).String;
						if (oNewValue == oOldValue)
						{
							oFlag = true;
							break;
						}
					}
				}
			}
			return oFlag ;
		}

		#endregion

		#region GENERATE LOG ENTRY

		public static void LogEntry(string oError, string oProcessName)
		{

            string date = DateTime.Now.Date.Day.ToString() + DateTime.Now.Date.Month.ToString() + DateTime.Now.Date.Year.ToString();
            if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Log\\Systemlog"))
            {
                Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Log\\Systemlog");
            }
            FileStream fs = File.Open(System.Windows.Forms.Application.StartupPath + "\\Log\\Systemlog\\Systemlog" + date + ".txt", FileMode.Append, FileAccess.Write);
            StreamWriter oSw = new StreamWriter(fs);
            oSw.WriteLine(oProcessName + ": " + oError);
            oSw.Flush();
            oSw.Close();
		}

		#endregion

		#region CHECK EXISTANCE OS EXCISABLE ITEM

		/// <summary>
		/// This function is written to check whether the excisable item 
		/// excists in particula
		/// </summary>
		/// <param name="Table"></param>
		/// <returns></returns>
		public static  bool CheckExcisable(string oTable,string oDocEntry)
		{
			string oSql = "Select IsNull(Count(*),0) From " + oTable + " Where excisable = 'Y' And DocEntry = " + oDocEntry;
			SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			Utilities.ExecuteSQL(ref oRecordSet,oSql);
			if (Convert.ToInt16(oRecordSet.Fields.Item(0).Value) > 0)
				return true;
			else
				return false;
		}

		#endregion

		#region GET WHAREHOUSES
		
		public static string GetQCWhs()
		{
			try
			{
				string oSql = "Select WhsCode From OWHS Where U_QCCheck = 'Y'";
				SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
				Utilities.ExecuteSQL(ref oRS,oSql);
				return oRS.Fields.Item("WhsCode").Value.ToString();
			}
			catch(Exception ex)
			{
				throw ex;
			}
		}

		public static string GetRejectedWhs()
		{
			try
			{
				string oSql = "Select WhsCode From OWHS Where U_RejWhs = 'Y'";
				SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
				Utilities.ExecuteSQL(ref oRS,oSql);
				return oRS.Fields.Item("WhsCode").Value.ToString();
			}
			catch(Exception ex)
			{
				throw ex;
			}
		}

		#endregion

		#region GET DEFAULT WHAREHOUSE

		public static string GetItemDefaultWhs(string oItemCode)
		{
			string oWhsCode = "";
			string oSql     = "select DfltWh from oitm Where ItemCode = '" + oItemCode.Trim() + "'";
			SAPbobsCOM.Recordset oRS =(SAPbobsCOM.Recordset)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
			Utilities.ExecuteSQL(ref oRS,oSql);
			if (oRS.Fields.Item(0).Value.ToString() != "")
				oWhsCode = oRS.Fields.Item(0).Value.ToString();

			return oWhsCode;
		}

		#endregion
		
		#region SERIALIZED MATRIX ROW 

		public static void SerializedMartix(ref SAPbouiCOM.Form _Form, ref SAPbouiCOM.Matrix oMatrix)
		{
			try
			{
				_Form.Freeze (true);
				for(int i=1;i<=oMatrix.RowCount;i++)
				{
                  
                  ((SAPbouiCOM.EditText)oMatrix.Columns.Item("0").Cells.Item(i).Specific).Value  = i.ToString();
				}
				
			}
			catch(Exception ex)
			{}
			finally
			{
				_Form.Freeze (false);
				_Form.Update();
			}
		}

		#endregion

		#region GET ITEM COST

		public static double GetItemCost(string oItemCode , string oWhs)
		{
			SAPbobsCOM.Items oItem ;
			double oPrice =0 ;
			int i;
			oItem = (SAPbobsCOM.Items)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
			oItem.GetByKey(oItemCode);
			for(i=0;i<=oItem.WhsInfo.Count - 1;i++)
			{
				oItem.WhsInfo.SetCurrentLine(i);
				
				if( oItem.WhsInfo.WarehouseCode == oWhs.Trim())
				{
					oPrice = oItem.WhsInfo.StandardAveragePrice;
					break;
				}
			}
			return oPrice;
		}

		public static double GetItemCost(string oItemCode)
		{
			string oSQL = "Select AvgPrice From OITM Where ItemCode = '" + oItemCode + "'";
			Utilities.ExecuteSQL(ref oRS1,oSQL);
			return Convert.ToDouble(oRS1.Fields.Item("AvgPrice").Value);
		}

        public static double GetBasePrice(string oItemCode)
        {
            SAPbobsCOM.Items oItem;
            double oPrice = 0;
            int i;
            oItem = (SAPbobsCOM.Items)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems);
            if (oItem.GetByKey(oItemCode))
            {
                oPrice = oItem.PriceList.Price;
            }
            return oPrice;
        }
		#endregion

        #region CREATE FOLDERS
        public static void CreateFolders()
        {
            try
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports"))
                {
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports");

                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Purchase"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Purchase");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\PurchaseReturn"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\PurchaseReturn");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Sales"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Sales");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\SalesReturn"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\SalesReturn");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\InputToProduction"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\InputToProduction");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\OutputFromProduction"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\OutputFromProduction");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Customers"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Customers");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Suppliers"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Suppliers");
                    }
                    if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + @"\Reports\Items"))
                    {
                        Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + @"\Reports\Items");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        #endregion

        #region INITIALIZE CR DISPLAY USED WHILE WOKING WITH CRYSTAL REPORTS
        //USE WHILE WOKING WITH CRYSTAL REPORTS 
        //public static void InitializeCRDisplay()
        //{
        //    Constants.gConnStr                   = Constants.gobjSQLCon.ConnectionString;
        //    CRDisplay.clsCRDisplay.SQLServerName = Utilities.Application.Company.Server;
        //    CRDisplay.clsCRDisplay.UserID        = Constants.gUSER_ID;
        //    CRDisplay.clsCRDisplay.Password      = Constants.gUSER_PASSWORD;
        //    CRDisplay.clsCRDisplay.Database      = Constants.gobjSQLCon.Database;
        //}
        #endregion

        #region ATTACHED FORMATTED SEARCH
        public static void AttachFormattedSearch(string oSQL, string oQueryCategoryName, string oFormID, string oItemID, string oColumnID, int oQueryID)
        {
            SAPbobsCOM.Recordset oRecordSet = null;

            oSQL = "Select ISNull(MAX(CategoryId),0) + 1 From OQCN Where CategoryId>0";
            Utilities.ExecuteSQL(ref oRecordSet, oSQL);

            SAPbobsCOM.QueryCategories oQueryCategory = (SAPbobsCOM.QueryCategories)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);

            //oQueryCategory.Name = "bizRental";
            oQueryCategory.Name = oQueryCategoryName;
            int oRetVal = oQueryCategory.Add();
            if (oRetVal != 0)
            {
                int oErrCode;
                string oErrMsg;
                Utilities.Application.Company.GetLastError(out oErrCode, out oErrMsg);
            }
            SAPbobsCOM.FormattedSearches oFormattedSearch = (SAPbobsCOM.FormattedSearches)Utilities.Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
            oFormattedSearch.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
            oFormattedSearch.ColumnID = oColumnID;
            oFormattedSearch.ByField = SAPbobsCOM.BoYesNoEnum.tYES;
            oFormattedSearch.FormID = oFormID;
            oFormattedSearch.ItemID = oItemID;
            oFormattedSearch.QueryID = oQueryID;
            oRetVal = oFormattedSearch.Add();
            if (oRetVal != 0)
            {
                int oErrCode;
                string oErrMsg;
                Utilities.Application.Company.GetLastError(out oErrCode, out oErrMsg);
            }
        }
        #endregion

        #region COMPANY SEPARATOR

        public static string GetCompanySeparator()
        {
            oRS1 = null;
            string oSql = "SELECT top 1 isnull(U_B1FCMSEP,'') U_B1FCMSEP FROM [@B1F_FSPM]"; //WHERE U_B1FCMNAM ='" + Utilities.Application.Company.CompanyName + "'";
            Utilities.ExecuteSQL(ref oRS1, oSql);
            return oRS1.Fields.Item("U_B1FCMSEP").Value.ToString();
        }
        #endregion

        #region GET ATTRIBUTE CODE LENGHT

        public static int GetAttriCodeLen(string oAttributeCode)
        {
            oRS1 = null;
            string oSql = "SELECT U_B1FATLEN FROM [@B1F_FSAM] WHERE U_B1FATCDE = '" + oAttributeCode + "' AND (U_B1FINACT IS NULL OR U_B1FINACT = 'N')";
            Utilities.ExecuteSQL(ref oRS1, oSql);
            return Convert.ToInt32(oRS1.Fields.Item("U_B1FATLEN").Value.ToString());
        }
        #endregion

       public static List<string> GetBatchNumList(string query)
       {
            List<string> result = new List<string>();
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oApplication.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            rs.DoQuery(query);
            while (!rs.EoF)
            {
                result.Add(rs.Fields.Item(0).Value.ToString().Trim());
                rs.MoveNext();
            }
            return result;
        }



        public static SqlConnection Connection()
        {
            SqlConnection conn = null;
            try
            {
                conn = new SqlConnection("SERVER='" + Constants.gSERVER + "';DATABASE='" + Utilities.Application.Company.CompanyDB + "';USER ID='" + Constants.gUSER_ID + "';Password='" + Constants.gUSER_PASSWORD + "'");
                if (conn.State == ConnectionState.Closed)
                {
                    conn.Open();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            return conn;
        }

        #region GET ITEM AVAILABLE STATUS
        public static bool VGetItemAvailableStatus(string oItemCode, string oVar1, string oVar2, string oColName)
        {
            string oSQL = string.Empty;
            DataSet oDataSet;

            //oCol = "U_B1FH" + oCol;
            if (oVar1 != "" && oVar2 == "")
            {
            //    Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
            // .Append(" SET @SQL = 'SELECT ' ")
            //.Append(" +(SELECT  'U_B1FH'")
            //.Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
            //.Append(oColName)
            //.Append("' AND U_B1FSEQ = (SELECT ")
            //.Append(" U_B1FGRPSQ FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
            //.Append(oItemCode)
            //.Append("' AND U_B1FATCDE = '")
            //.Append(oColName)
            //.Append("' ) AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
            // .Append(oItemCode)
            // .Append("' AND U_B1FATCDE = '")
            // .Append(oColName)
            // .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
            //.Append(oItemCode)
            //.Append("'' AND U_B1FVAR1= ''")
            //.Append(oVar1)
            //.Append("'' ' EXEC(@SQL)");
            //    oSQL = Sob.ToString();
                oSQL = "Select " + oColName + " from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "'  and U_B1FVAR1 = '" + oVar1 + "'";

            }
            //oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR1= ''" + oVar1 + "'' ' EXEC(@SQL)";
            else if (oVar1 == "" && oVar2 != "")
            {
              //  Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
              // .Append(" SET @SQL = 'SELECT ' ")
              //.Append(" +(SELECT  'U_B1FH'")
              //.Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
              //.Append(oColName)
              //      //.Append("' AND LINEID = (SELECT ")
              // .Append("' AND U_B1FSEQ = (SELECT ")
              //.Append(" U_B1FGRPSQ FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
              //.Append(oItemCode)
              //.Append("' AND U_B1FATCDE = '")
              //.Append(oColName)
              //.Append("' ) AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
              // .Append(oItemCode)
              // .Append("' AND U_B1FATCDE = '")
              // .Append(oColName)
              // .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
              //.Append(oItemCode)
              //.Append("'' AND U_B1FVAR2= ''")
              //.Append(oVar2)
              //.Append("'' ' EXEC(@SQL)");

              //  oSQL = Sob.ToString();

                oSQL = "Select " + oColName + " from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "'  and U_B1FVAR2 = '" + oVar2 + "'";

                //oSQL = "select (SELECT 'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR2='" + oVar2 + "'";
                //   oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR2= ''" + oVar2 + "'' ' EXEC(@SQL)";
            }
            else if (oVar1 != "" && oVar2 != "")
            {
             //   Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
             //.Append(" SET @SQL = 'SELECT ' ")
            //.Append(" +(SELECT ")
            //.Append(oColName)
            // .Append(" from [@B1F_FSC] where U_B1FITMNO = '")
            //.Append(oItemCode)
            //.Append("' AND U_B1FVAR1= '")
            //.Append(oVar1)
            // .Append("' AND U_B1FVAR2= '")
            //.Append(oVar2)
            //.Append(oVar2);
            ////.Append("'' ' EXEC(@SQL)");
            //    oSQL = Sob.ToString();
                oSQL = "Select " + oColName + " from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR1 ='" + oVar1 + "' and U_B1FVAR2 = '" + oVar2 + "'";
            }
            oDataSet = SqlHelper.ExecuteDataset(Constants.gobjSQLCon, CommandType.Text, oSQL);
            // Utilities.ExecuteSQL(ref oRS1, oSQL);
            if (oDataSet.Tables[0].Rows.Count > 0)
            {
                if (oDataSet.Tables[0].Rows[0][0].ToString() == "Y")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public static bool GetItemAvailableStatus(string oItemCode,string oVar1,string oVar2,string oCol, string oColName)
        {
            string oSQL=string.Empty;
            StringBuilder Sob = null;  
            DataSet oDataSet;

            oCol = "U_B1FH" + oCol;
            if (oVar1 != "" && oVar2 == "")
            {
                Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
          .Append(" SET @SQL = 'SELECT ' ")
            .Append(" +(SELECT  'U_B1FH'")
            .Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
            .Append(oColName)
            .Append("' AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
             .Append(oItemCode)
             .Append("' AND U_B1FATCDE = '")
             .Append(oColName)
             .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
            .Append(oItemCode)
            .Append("'' AND U_B1FVAR1= ''")
            .Append(oVar1)
            .Append("'' ' EXEC(@SQL)");
                oSQL = Sob.ToString();
            }
            //oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR1= ''" + oVar1 + "'' ' EXEC(@SQL)";
            else if (oVar1 == "" && oVar2 != "")
            {
                Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
             .Append(" SET @SQL = 'SELECT ' ")
            .Append(" +(SELECT  'U_B1FH'")
            .Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
            .Append(oColName)
            .Append("' AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
             .Append(oItemCode)
             .Append("' AND U_B1FATCDE = '")
             .Append(oColName)
             .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
              .Append(oItemCode)
              .Append("'' AND U_B1FVAR2= ''")
              .Append(oVar2)
              .Append("'' ' EXEC(@SQL)");

              // .Append(" SET @SQL = 'SELECT ' ")
              //.Append(" +(SELECT  'U_B1FH'")
              //.Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
              //.Append(oColName)
              // .Append("' AND U_B1FSEQ = (SELECT ")
              //.Append(" U_B1FGRPSQ FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
              //.Append(oItemCode)
              //.Append("' AND U_B1FATCDE = '")
              //.Append(oColName)
              //.Append("' ) AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
              // .Append(oItemCode)
              // .Append("' AND U_B1FATCDE = '")
              // .Append(oColName)
              // .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
              //.Append(oItemCode)
              //.Append("'' AND U_B1FVAR2= ''")
              //.Append(oVar2)
              //.Append("'' ' EXEC(@SQL)");

                oSQL = Sob.ToString();
                //oSQL = "select (SELECT 'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR2='" + oVar2 + "'";
             //   oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR2= ''" + oVar2 + "'' ' EXEC(@SQL)";
            }
            else if (oVar1 != "" && oVar2 != "")
            {
                Sob = new StringBuilder("DECLARE @SQL AS VARCHAR(1000)")
            .Append(" SET @SQL = 'SELECT ' ")
            .Append(" +(SELECT  'U_B1FH'")
            .Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
            .Append(oColName)
            .Append("' AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
             .Append(oItemCode)
             .Append("' AND U_B1FATCDE = '")
             .Append(oColName)
             .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
            .Append(oItemCode)
            .Append("'' AND U_B1FVAR1= ''")
            .Append(oVar1)
             .Append("'' AND U_B1FVAR2= ''")
            .Append(oVar2)
            .Append("'' ' EXEC(@SQL)");

            // .Append(" SET @SQL = 'SELECT ' ")
            //.Append(" +(SELECT  'U_B1FH'")
            //.Append("+(select CAST(U_B1FSEQ AS VARCHAR) From  [@B1F_FSAC] where U_B1FCODE ='")
            //.Append(oColName)
            ////.Append("' AND LINEID = (SELECT ")
            // .Append("' AND U_B1FSEQ = (SELECT ")

            //.Append(" U_B1FGRPSQ FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
            //.Append(oItemCode)
            //.Append("' AND U_B1FATCDE = '")
            //.Append(oColName)
            //.Append("' ) AND U_B1FGRP = (SELECT U_B1FATGRP FROM [@B1F_HOR1] WHERE U_B1FITMNO = '")
            // .Append(oItemCode)
            // .Append("' AND U_B1FATCDE = '")
            // .Append(oColName)
            // .Append("'))) + ' from [@B1F_FSC] where U_B1FITMNO = ''")
            //.Append(oItemCode)
            //.Append("'' AND U_B1FVAR1= ''")
            //.Append(oVar1)
            // .Append("'' AND U_B1FVAR2= ''")
            //.Append(oVar2)
            //.Append("'' ' EXEC(@SQL)");
                oSQL = Sob.ToString();
                //oSQL = "select (SELECT 'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR1='" + oVar1 + "' And U_B1FVAR2='" + oVar2 + "'";
               // oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR1 = ''" + oVar1 + "'' AND U_B1FVAR2= ''" + oVar2 + "'' ' EXEC(@SQL)";
            }
          oDataSet = SqlHelper.ExecuteDataset(Constants.gobjSQLCon, CommandType.Text, oSQL);
           // Utilities.ExecuteSQL(ref oRS1, oSQL);
            if (oDataSet.Tables[0].Rows.Count > 0)
            {
                if (oDataSet.Tables[0].Rows[0][0].ToString() == "Y")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }
        public static bool RGetItemAvailableStatus(string oItemCode, string oVar1, string oVar2, string oCol, string oColName)
        {
            string oSQL = "";
            DataSet oDataSet;

            oCol = "U_B1FH" + oCol;
            if (oVar1 != "" && oVar2 == "")
                oSQL = "Select " + oCol + "  from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' AND U_B1FVAR1= '" + oVar1 + "'";
                //oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR1= ''" + oVar1 + "'' ' EXEC(@SQL)";
            else if (oVar1 == "" && oVar2 != "")
                //oSQL = "select (SELECT 'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR2='" + oVar2 + "'";
                oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR2= ''" + oVar2 + "'' ' EXEC(@SQL)";
            else if (oVar1 != "" && oVar2 != "")
            {
                //oSQL = "select (SELECT 'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) from [@B1F_FSC] where U_B1FITMNO = '" + oItemCode + "' and U_B1FVAR1='" + oVar1 + "' And U_B1FVAR2='" + oVar2 + "'";
                oSQL = "DECLARE @SQL AS VARCHAR(1000) SET @SQL = 'SELECT ' + (SELECT  'U_B1FH' + CAST(U_B1FGRPSQ AS VARCHAR) FROM [@B1F_HOR1] WHERE U_B1FITMNO = '" + oItemCode + "' AND U_B1FATCDE = '" + oColName + "' ) + ' from [@B1F_FSC] where U_B1FITMNO = ''" + oItemCode + "'' AND U_B1FVAR1 = ''" + oVar1 + "'' AND U_B1FVAR2= ''" + oVar2 + "'' ' EXEC(@SQL)";
            }
            oDataSet = SqlHelper.ExecuteDataset(Constants.gobjSQLCon, CommandType.Text, oSQL);
            // Utilities.ExecuteSQL(ref oRS1, oSQL);
            if (oDataSet.Tables[0].Rows.Count > 0)
            {
                if (oDataSet.Tables[0].Rows[0][0].ToString() == "Y")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }
        }

        #endregion

        #region HIDE SYSTEM MESSEGE
        public static bool HideSystemMessege(ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                if (pVal.BeforeAction == true && pVal.FormType == 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {

                    SAPbouiCOM.Form oFrm;
                    oFrm =Utilities.Application.SBO_Application.Forms.GetForm("0", pVal.FormTypeCount);
                    oFrm.Select();
                    if (_HideSystemMessege == true && iHideCount == 1)
                    {
                        _HideSystemMessege = false;
                        iHideCount = 2;
                        oFrm.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }

                if (pVal.BeforeAction == true && pVal.FormType == 0 && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    iHideCount = 1;

                }
                // ------------ Add Form Type within this block  for Hiding System Message ------------------------------------------------
                if (pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    // Add Form Type
                    if (pVal.FormType == 2020090707)
                    {
                        _HideSystemMessege = true;
                    }
                }
                if (pVal.BeforeAction == false && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                {
                    // Add Form Type

                    if (pVal.FormType == 2020090707)
                    {
                        _HideSystemMessege = false;
                    }
                }
                //else if (pVal.BeforeAction == true && pVal.FormType == 2000200714)
                //{
                //    _HideSystemMessege = true;
                //}
                //--------------------------------------------------------------------------------------------------------------------
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
                return false;
            }
            return true;
        }

        #endregion

        #region INITIALIZE CR-DISPLAY
        public static void InitializeCRDisplay()
        {
            //Constants.gConnStr = Constants .gobjSQLCon.ConnectionString;
            //CRDisplay.clsCRDisplay.SQLServerName = Utilities.Application.Company.Server;
            //CRDisplay.clsCRDisplay.UserID = Constants.gUSER_ID;
            //CRDisplay.clsCRDisplay.Password = Constants.gUSER_PASSWORD;
            //CRDisplay.clsCRDisplay.Database = Constants.gobjSQLCon.Database;
        }
        #endregion

        public static void CreateConfig()
        {
            if (!File.Exists(getApplicationPath() + @"\AppConfig.xml"))
            {
                File.WriteAllLines(getApplicationPath() + @"\AppConfig.xml", File.ReadAllLines(getApplicationPath() + @"\ConfigTemplate.xml"));
            }
        }

        #region FUNCTION FOR WRITE IMPORTLOG.LOG FILE

        public static void WriteLogFile(string Message)
        {
            try
            {
                if (!Directory.Exists(System.Windows.Forms.Application.StartupPath + "\\Log"))
                    Directory.CreateDirectory(System.Windows.Forms.Application.StartupPath + "\\Log");

                string date = DateTime.Now.Date.Day.ToString() + DateTime.Now.Date.Month.ToString() + DateTime.Now.Date.Year.ToString();
                FileStream fs = File.Open(System.Windows.Forms.Application.StartupPath + "\\Log\\Delivery" + date + ".txt", FileMode.Append, FileAccess.Write);
                StreamWriter sr = new StreamWriter(fs);
                sr.WriteLine(Message);
                sr.Flush();
                sr.Close();
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        #endregion      

        #region Create Store Procedure
        /// <summary>
        /// 
        /// </summary>
        internal static void CreateFunction_RemoveSpecialChar()
          {
            string _SQL = "";

            _SQL = @"IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[RemoveSpecialChars]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
                    DROP FUNCTION [dbo].[RemoveSpecialChars]";
            Utilities.ExecuteSQL(_SQL);

            _SQL = @"EXEC dbo.sp_executesql @statement = N'CREATE  FUNCTION [dbo].[RemoveSpecialChars] ( @InputString VARCHAR(8000) )
                    RETURNS VARCHAR(8000)  
                    BEGIN
                        IF @InputString IS NULL
                            RETURN NULL
                        DECLARE @OutputString VARCHAR(8000)
                        SET @OutputString = ''''
                        DECLARE @l INT
                        SET @l = LEN(@InputString)
                        DECLARE @p INT
                        SET @p = 1
                        WHILE @p <= @l
                            BEGIN
                                DECLARE @c INT
                                SET @c = ASCII(SUBSTRING(@InputString, @p, 1))
                                IF @c BETWEEN 48 AND 57
                                    OR @c BETWEEN 65 AND 90
                                    OR @c BETWEEN 97 AND 122
                                      --OR @c = 32
                                    SET @OutputString = @OutputString + CHAR(@c)
                                SET @p = @p + 1
                            END
                        IF LEN(@OutputString) = 0
                            RETURN NULL
                        RETURN @OutputString
                    END ' ";
            Utilities.ExecuteSQL(_SQL);
        }
        #endregion

        #region SET FILTERS
        public static void setFilter()
        {
            SAPbouiCOM.EventFilters objFilters;
            SAPbouiCOM.EventFilter objFilter;
            objFilters = new SAPbouiCOM.EventFilters();

            #region FORM ACTIVATE
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE);
            objFilter.AddEx("b1FPackRatio");  //ClsKolliDetails
            #endregion

            #region FORM LOAD
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD);
            objFilter.AddEx("b1FPackRatio");  //ClsKolliDetails
            #endregion

            #region FORM RESIZE
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_RESIZE);
            #endregion

            #region FORM CLOSE
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_CLOSE);
            objFilter.AddEx("b1FPackRatio");                       //Minimum quantity form

            #endregion

            #region FORM UNLOAD
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD);
            objFilter.AddEx("b1FPackRatio");  //ClsKolliDetails
            #endregion
            
            #region CFL
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST);
           // objFilter.AddEx("DeliveryStatus");              //DeliveryStatus
            #endregion

            #region ITEM PRESSED
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED);
            objFilter.AddEx("b1FPackRatio");                //ClsKolliDetails
       
            #endregion
                     
            #region GOT FOCUS
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_GOT_FOCUS);
            #endregion

            #region LOST FOCUS
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_LOST_FOCUS);
            #endregion

            #region VALIDATE
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_VALIDATE);
            //objFilter.AddEx("140"); //Delivery
            #endregion

            #region RIGHT CLICK
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_RIGHT_CLICK);

            #endregion

            #region MENU CLICK
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MENU_CLICK);
           // objFilter.AddEx("DeliveryStatus");    //DeliveryStatus
            objFilter.AddEx("b1FPackRatio");
            #endregion

            #region KEY DOWN
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_KEY_DOWN);
            objFilter.AddEx("b1FPackRatio");      //DeliveryStatus
            #endregion

            #region DOUBLE CLICK
            //objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK);
            //objFilter.AddEx("134");
            #endregion

            #region LINK BUTTON
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED);
           // objFilter.AddEx("KolliDetails");  //ClsKolliDetails
            #endregion

            #region COMBO SELECT
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_COMBO_SELECT);
            //objFilter.AddEx("Status");  //Status

            #endregion

            #region CLICK
            objFilter = objFilters.Add(SAPbouiCOM.BoEventTypes.et_CLICK);
            //objFilter.AddEx("KolliDetails");  //ClsKolliDetails
            objFilter.AddEx("b1FPackRatio");
            #endregion

            oApplication.SBO_Application.SetFilter(objFilters);

        }
        #endregion

        #region Create Menu
        public static void CreateMenus()
        {
            
            string SQl = "";
            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            SAPbobsCOM.Recordset RsCreateMenu = null;
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.Menus oSubmenu = null;
            SAPbouiCOM.MenuItem oMenuItem = null;
            oMenus = Utilities.Application.SBO_Application.Menus;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Utilities.Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOL_01"); // Administration'
            oSubmenu = oMenuItem.SubMenus;
            try
            {
                SQl = @"select isnull(U_B1FIPACKA,''),  ISNULL(U_B1FRSACT,'') AS U_B1FRSACT from [@B1F_FSPM] ";
                RsCreateMenu = (SAPbobsCOM.Recordset)Utilities.ExecuteSQL(SQl);
                if (RsCreateMenu.RecordCount > 0)
                {
                    if (RsCreateMenu.Fields.Item(0).Value.Equals("Y"))
                    {
                        if (oMenuItem.SubMenus.Exists("B1SOL_03_11") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOL_03_11";
                            oCreationPackage.String = "HarshInternational Pack Ratios";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 0;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                    }
                    else
                    {
                        if (oMenuItem.SubMenus.Exists("B1SOL_03_11") == true)
                        {
                            oMenus.RemoveEx("B1SOL_03_11"); 
                        }
                    }
                    string _ShopType = RsCreateMenu.Fields.Item("U_B1FRSACT").Value.ToString();
                    if (_ShopType.Equals("Y"))
                    {
                        // Added By AMIT
                        //------------- Shop Under Sales AR -------------

                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("2048");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLS_23") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                            oCreationPackage.UniqueID = "B1SOLS_23";
                            oCreationPackage.String = "Shops";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 17;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOLS_23");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLS_23_1") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLS_23_1";
                            oCreationPackage.String = "Shop Sell Through Recording";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 0;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }

                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOLS_23");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLS_23_2") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLS_23_2";
                            oCreationPackage.String = "Update Shop Sales";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 1;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        if (oMenuItem.SubMenus.Exists("B1SOLC_30_01") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLC_30_01";
                            oCreationPackage.String = "Calculate Shop Replenishment";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 2;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        //Added By Shalini
                        if (oMenuItem.SubMenus.Exists("B1SOLE_30_01") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLE_30_01";
                            oCreationPackage.String = "Edit And Generate Shop Replenishment";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 3;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        //--Added By Shalini

                        //Sale AR --> Shops --> Shop Imports & Exports 
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOLS_23");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLS_24") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                            oCreationPackage.UniqueID = "B1SOLS_24";
                            oCreationPackage.String = "Shop Imports & Exports";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 4;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        //Sale AR --> Shops --> [Shop Imports & Exports] -->[Boots Daily Sales Import] 
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOLS_24");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLS_24_01") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLS_24_01";
                            oCreationPackage.String = "Boots Daily Sales Import";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 0;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        //Sale AR --> Shops --> [Generate Shop Replenishment]
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOLS_23");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOLG_30_01") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOLG_30_01";
                            oCreationPackage.String = "Generate Shop Replenishment";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 5;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        //-----------------------------------------------------
                        //-----  //------------- Shop Under Business Partner  -------------
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("43535"); // Administration'
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOL_25") != true) 
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                            oCreationPackage.UniqueID = "B1SOL_25";
                            oCreationPackage.String = "Shops";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 7;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("B1SOL_25");
                        oSubmenu = oMenuItem.SubMenus;
                        if (oMenuItem.SubMenus.Exists("B1SOL_25_01") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOL_25_01";
                            oCreationPackage.String = "Retail Calendar Maintenance";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 0;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        if (oMenuItem.SubMenus.Exists("B1SOL_25_02") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOL_25_02";
                            oCreationPackage.String = "Shop Stock Patterns";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 1;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                        if (oMenuItem.SubMenus.Exists("B1SOL_25_03") != true)
                        {
                            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                            oCreationPackage.UniqueID = "B1SOL_25_03";
                            oCreationPackage.String = "Shop Issues";
                            oCreationPackage.Enabled = true;
                            oCreationPackage.Position = 2;
                            oMenuItem.SubMenus.AddEx(oCreationPackage);
                        }
                    }
                    else if (_ShopType.Equals("N")) 
                    {
                        // Sales AR 
                        oMenuItem = Utilities.Application.SBO_Application.Menus.Item("2048");
                        oSubmenu = oMenuItem.SubMenus;
                         if (oMenuItem.SubMenus.Exists("B1SOLS_23"))
                        {
                            oMenuItem.SubMenus.RemoveEx("B1SOLS_23");
                        }

                         //-----  //------------- Shop Under Business Partner  -------------
                         oMenuItem = Utilities.Application.SBO_Application.Menus.Item("43535"); // Administration'
                         oSubmenu = oMenuItem.SubMenus;
                         if (oMenuItem.SubMenus.Exists("B1SOL_25"))
                         {
                             oMenuItem.SubMenus.RemoveEx("B1SOL_25");
                         }
                    }
                   
                   

                     
                    
                }
            }
            catch (Exception er)
            {
                Utilities.Application.SBO_Application.MessageBox("Menu Already Exists", 1, "Ok", "", "");
            }
            finally
            {
                if (RsCreateMenu != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(RsCreateMenu);
                     RsCreateMenu = null;
                }
                   if(oMenus!=null)
                   {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenus);
                    oMenus = null;
                   }
                if(oCreationPackage!=null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oCreationPackage);
                    oCreationPackage=null;
                }
                 
                if(oMenuItem!=null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oMenuItem);
                    oMenuItem=null;
                }
                if(oSubmenu!=null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSubmenu);
                    oSubmenu = null;
                }       
                }

                    }
        #endregion

        #region Add Formatted Searches
        /// <summary>
        /// Add formated search (Writing Query and call mail method for formatted search) 
        /// </summary>
        internal static void AddFormatedSearches()
        {
            string _SqlQuery = string.Empty;

            #region Add Formatted Search
            /* Form "Fashion Attribute */
            _SqlQuery = @"Select  Distinct U_B1FATCDE  , U_B1FATDSC  From  [@B1F_FSAM]
                                Where U_B1FAXIS = 1 AND ISNULL(U_B1FINACT,'') <> 'Y'";
            AddFormatedSearch("FMS_HorizontalAttribute", _SqlQuery, "2020090710", "MatAttrDtl", "B1FHORA", "U_B1FHORA", false);
            // Add Formated search on the Fashion Allowances
            _SqlQuery = @"Select P1.Code , P1.Name  From [@B1F_BRND] P1 Where IsNull(P1.U_B1FINACT,'N')  <> 'Y'";
            AddFormatedSearch("FMS_FashionAllowance", _SqlQuery, "2020090711", "6", "V_2", "U_B1FVALU", false);


            _SqlQuery = "SELECT WhsCode,WhsName  FROM OWHS ORDER BY WhsCode";
            AddFormatedSearch("b1sol_StockWarehouseCode", _SqlQuery, "134", "U_B1FSWHSE", "-1", "U_B1FSWHSE", false);

            _SqlQuery = "SELECT WhsCode,WhsName  FROM OWHS ORDER BY WhsCode";
            AddFormatedSearch("b1sol_TransitWarehouseCode", _SqlQuery, "134", "U_B1FSWHSET", "-1", "U_B1FSWHSET", false);

            _SqlQuery = "SELECT ListNum,ListName  FROM OPLN ORDER BY LISTNUM";
            AddFormatedSearch("b1sol_PriceList", _SqlQuery, "134", "U_B1FSTPL", "-1", "U_B1FSTPL", false);


            //-----------------------------------------
            AddQueryCatagoryInQueriesManager("Shops");

            _SqlQuery = @"DECLARE @DateFrom		AS VARCHAR(20) 
            DECLARE @DateTo			AS VARCHAR(20)  
            Declare @FatherCard		AS VARCHAR(30) 

            SELECT @DateFrom    = ISNULL(T1.[DocDueDate],'') FROM ODLN  T1  where T1.DocDueDate='[%0]' 
            SELECT @DateTo    = ISNULL(T1.[DocDueDate],'') FROM ODLN  T1  where T1.DocDueDate='[%1]' 
            SELECT @FatherCard   =  ISNULL(T2.[FatherCard],'')  FROM OCRD T2 INNER JOIN ODLN  T1 ON T2.CardCode = T1.CardCode Where T2.FatherCard='[%2]'   

            EXEC B1Sol_BootsASNFile  @DateFrom , @DateTo , @FatherCard";
            AddUserQueryinCatagoryOfQueryManager("Boots ASN File", _SqlQuery, "Shops");
            //--------------------------------------------------------------------------

            #endregion
        }
        #endregion

        #region Add Formatted Search Column
        /// <summary>
        /// Genric function for Add Fromatted Search in SAP b1
        /// </summary>
        /// <param name="queryName">Query Name </param>
        /// <param name="query"> SQL Query</param>
        /// <param name="Form">Form ID</param>
        /// <param name="item">Item ID (i.e MatrixID ,ItemID)</param>
        /// <param name="colid">Column ID</param>
        /// <param name="Field">Database Field</param>
        /// <param name="Autoref">Auto ref. </param>
        internal static void AddFormatedSearch(string queryName, string query, string Form, string item, string colid, string Field, bool Autoref)
        {
            SAPbobsCOM.Recordset oRs = null;
            SAPbobsCOM.UserQueries oQuery = null;


            //Add user query
            try
            {
                oRs = (SAPbobsCOM.Recordset)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ExecuteSQL(ref oRs, "Select count(*) From OUQR(nolock) where QName = '" + queryName + "'");
                oQuery = (SAPbobsCOM.UserQueries)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);
                if ((int)oRs.Fields.Item(0).Value == 0)
                {
                    oQuery.Query = query; // "SELECT U_ContactCode As ContactCode,U_ContactName as ContactName from [@ObjectContact] order by U_ContactCode";
                    oQuery.QueryCategory = -1;
                    oQuery.QueryDescription = queryName;// "ContactCFL";
                    oQuery.Add();
                }
                else if ((int)oRs.Fields.Item(0).Value == 1)
                {

                    // query = "Select  ItemCode , ItemName from OITM";
                    oRs.DoQuery("select IntrnalKey  ,QCategory from OUQR(nolock) Where QName = '" + queryName + "'");
                    if (oQuery.GetByKey((int)oRs.Fields.Item("IntrnalKey").Value, (int)oRs.Fields.Item("QCategory").Value) == true)
                    {
                        if (oQuery.Query != query)
                        {
                            oQuery.Query = query;
                            oQuery.Update();
                        }
                    }
                }

                string _sql = @" Select count(*) From CSHS(nolock) where FormId = '" + Form + "' and ItemId = '" + item + "' AND ColID = '" + colid + "'";
                ExecuteSQL(ref oRs,_sql  );
                if ((int)oRs.Fields.Item(0).Value == 0)
                {
                    ////Add formated search
                    SAPbobsCOM.FormattedSearches oFormatted = (SAPbobsCOM.FormattedSearches)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    oFormatted.FormID = Form; //"320";
                    oFormatted.ItemID = item;//"U_ContactCode";
                    oFormatted.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                    oFormatted.FieldID = Field;//"U_ContactCode";
                    oFormatted.ColumnID = colid;
                    oRs.DoQuery("select IntrnalKey from OUQR(nolock) Where QName = '" + queryName + "'");
                    oFormatted.QueryID = (int)oRs.Fields.Item("IntrnalKey").Value;
                    if (Autoref)
                    {
                        oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                    }
                    else
                    {
                        oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tNO;
                    }
                    oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO;
                    oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO;

                    int i = oFormatted.Add();
                }
                else
                {
                 _sql = @"Select isnull(QueryId,0),Indexid From CSHS(nolock) where FormId = '" + Form + "' and ItemId = '" + item + "' AND ColID = '" + colid + "'";
                    ExecuteSQL(ref oRs,_sql  );
                    if ((int)oRs.Fields.Item(0).Value <= 0)
                    {
                        
                    SAPbobsCOM.FormattedSearches oFormatted = (SAPbobsCOM.FormattedSearches)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oFormattedSearches);
                    if (oFormatted.GetByKey((int)oRs.Fields.Item(1).Value))
                    {
                        oFormatted.FormID = Form; //"320";
                        oFormatted.ItemID = item;//"U_ContactCode";
                        oFormatted.Action = SAPbobsCOM.BoFormattedSearchActionEnum.bofsaQuery;
                        oFormatted.FieldID = Field;//"U_ContactCode";
                        oFormatted.ColumnID = colid;
                        oRs.DoQuery("select IntrnalKey from OUQR(nolock) Where QName = '" + queryName + "'");
                        oFormatted.QueryID = (int)oRs.Fields.Item("IntrnalKey").Value;
                        if (Autoref)
                        {
                            oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tYES;
                        }
                        else
                        {
                            oFormatted.Refresh = SAPbobsCOM.BoYesNoEnum.tNO;
                        }
                        oFormatted.ForceRefresh = SAPbobsCOM.BoYesNoEnum.tNO;
                        oFormatted.ByField = SAPbobsCOM.BoYesNoEnum.tNO;

                        int i = oFormatted.Update();
                    }
                    }
                }
                  
            }
            catch (Exception ex)
            {
                Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oQuery);

            }

        }

        
        #endregion  

        public static XmlDocument AddAppSetingSection(string Section, XmlDocument XDoc)
        {
            System.Xml.XmlElement XElement;
            System.Xml.XmlDocument OXmlDocument = XDoc;
            try
            {
                if (OXmlDocument.SelectNodes("/Data/Schedule[@ID='" + Section + "']").Count > 0)
                {
                    throw new Exception("Section already exist");
                }
                System.Xml.XmlNode _Node = OXmlDocument.SelectSingleNode("/Data");

                if (_Node == null)
                {
                    _Node = OXmlDocument.CreateElement("Data");
                    OXmlDocument.AppendChild(_Node);
                }

                XElement = OXmlDocument.CreateElement("Schedule");
                System.Xml.XmlAttribute KeyNode = OXmlDocument.CreateAttribute("ID");
                KeyNode.Value = Section;
                XElement.Attributes.Append(KeyNode);
                _Node.AppendChild(XElement);

                return XDoc;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static bool AddAppSetingValue(string Section, string Name, object Value)
        {
            System.Xml.XmlElement XElement;
            System.Xml.XmlNode ParentNode;
           // System.Xml.XmlAttribute KeyNode, ValueNode;
            System.Xml.XmlDocument OXmlDocument = new System.Xml.XmlDocument();
            try
            {
                string FilePath = Utilities.getApplicationPath().ToString() + @"\" + Section +".xml";
                if (System.IO.File.Exists(FilePath))
                {
                    OXmlDocument.Load(FilePath);
                }

                if (OXmlDocument.SelectNodes("/Data/Schedule[@ID='" + Section + "']").Count == 0)
                {

                    OXmlDocument = AddAppSetingSection(Section, OXmlDocument);

                }
                System.Xml.XmlNode _Node = OXmlDocument.SelectSingleNode("/Data/Schedule[@ID='" + Section + "']/" + Name);
                if (_Node != null)
                {
                    _Node.InnerXml = Value.ToString();
                }
                else
                {
                    ParentNode = OXmlDocument.SelectSingleNode("/Data/Schedule[@ID='" + Section + "']");
                    XElement = OXmlDocument.CreateElement(Name);
                    XElement.InnerXml = Value.ToString();
                    ParentNode.AppendChild(XElement);
                }

                using (XmlTextWriter xTextWriter = new XmlTextWriter(FilePath, Encoding.UTF8))
                {
                    OXmlDocument.Save(xTextWriter);
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #region UTILITY: AMOUNT CONVERSION
        internal static  double AmountConvert(string Amount)
        {
            double Amt1=0;
            double Amt2=0;
            try
            {
            if (Amount != "")
            {
                string[] Amt = Amount.Split(' ');
                if (Amt.Length > 1)
                {
                    double.TryParse(Amt[0], out Amt1);
                    double.TryParse(Amt[1], out Amt2);
                }
                else 
                {
                    double.TryParse(Amt[0], out Amt1);
                }
            }
            }
            catch (Exception ex)
            {
                return 0;
            }
            return Amt1 + Amt2;

        }

        internal static string AmountConvert(string Amount,out string currency)
        {
            double Amt1 = 0;
            double Amt2 = 0;
            currency = "";
            try
            {
                if (Amount != "")
                {
                    string[] Amt = Amount.Split(' ');
                    if (Amt.Length > 1)
                    {
                        double.TryParse(Amt[0], out Amt1);
                        double.TryParse(Amt[1], out Amt2);
                        if (Amt1 == 0)
                        {
                            currency = Amt[0];
                        }
                        else
                        {
                            currency = Amt[1];
                        }
                        
                    }
                    else
                    {
                        
                        double.TryParse(Amt[0], out Amt1);
                        if (Amt1 == 0)
                        {
                            double.TryParse(Amt[0].Substring(3), out Amt1);
                            double.TryParse(Amt[0].Substring(0,Amt[0].Length-3) , out Amt2);
                            if (Amt1 > 0 )
                            {
                                currency = Amt[0].Substring(3);
                            }
                            else if (Amt2 > 0)
                            {
                                currency = Amt[0].Substring(Amt[0].Length - 3);

                            }
                        }
                    }
                }
              
            }
            catch (Exception ex)
            {
                currency = "";
                return "0";
            }
           
            return (Amt1 + Amt2).ToString() ;

        }
        #endregion

        #region Add New Query Catagory in the Queries Manager
        /// <summary>
        /// Add New Query Catagory in the Queries Manager 
        /// </summary>
        /// <param name="_CatagoryName">Set Catagory Name</param>
        internal static void AddQueryCatagoryInQueriesManager(string _CatagoryName)
        {
            SAPbobsCOM.QueryCategories _QueryCategories = null;
            SAPbobsCOM.Recordset oRs = null;
            try
            {
                _QueryCategories = (SAPbobsCOM.QueryCategories)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oQueryCategories);
                
                //-------------------------------------------
                string _Sql = String.Format(@"Select CategoryId  From OQCN  Where CatName = '{0}'", _CatagoryName);
                ExecuteSQL(ref oRs, _Sql);
                //-------------------------------------------
                if (!string.IsNullOrEmpty(_CatagoryName) && (string.IsNullOrEmpty(oRs.Fields.Item("CategoryId").Value.ToString()) || oRs.Fields.Item("CategoryId").Value.ToString() == "0"))
                {
                     // Assign New Catagory Name 
                    _QueryCategories.Name = _CatagoryName;
                    _QueryCategories.Add();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (_QueryCategories != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_QueryCategories);
                if (oRs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);   
            }
        }
        #endregion

        #region Add User Query in specific Catagory of Query manager
        /// <summary>
        /// Add User Query in specific Catagory of Query manager
        /// </summary>
        /// <param name="queryName">Set Query Name</param>
        /// <param name="_SqlQuery">Set Sql Query </param>
        /// <param name="_CatagoryName">Set Catagory Name</param>
        private static void AddUserQueryinCatagoryOfQueryManager(string queryName, string _SqlQuery ,string _CatagoryName)
        {
            SAPbobsCOM.Recordset oRs = null;
            SAPbobsCOM.UserQueries oQuery = null;
            try
            {
                oQuery = (SAPbobsCOM.UserQueries)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserQueries);

                //-------------------------------------------
                string _Sql = String.Format(@"Select CategoryId  From OQCN  Where CatName = '{0}'", _CatagoryName);
                ExecuteSQL(ref oRs, _Sql);
                int _CatagoryID = (int)oRs.Fields.Item(0).Value;
                //-------------------------------------------

                oRs = (SAPbobsCOM.Recordset)Application.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                ExecuteSQL(ref oRs, String.Format("Select count(*) From OUQR(nolock) where QName = '{0}'", queryName));
                //--------------------------------------------
                if ((int)oRs.Fields.Item(0).Value == 0)  // Add 
                {
                    oQuery.Query = _SqlQuery; 
                    oQuery.QueryCategory = _CatagoryID;
                    oQuery.QueryDescription = queryName;
                    oQuery.Add();
                }
                else if ((int)oRs.Fields.Item(0).Value == 1) // Update
                {
                    oRs.DoQuery(String.Format("select IntrnalKey  ,QCategory from OUQR(nolock) Where QName = '{0}'", queryName));
                    if (oQuery.GetByKey((int)oRs.Fields.Item("IntrnalKey").Value, (int)oRs.Fields.Item("QCategory").Value) == true)
                    {
                        if (oQuery.Query != _SqlQuery)
                        {
                            oQuery.Query = _SqlQuery;
                            oQuery.Update();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
            finally
            {
                if (oQuery != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oQuery);
                if (oRs != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }
        #endregion

        #region Brose Folder
        public static string BrowseFolder(string text)
        {
            try
            {
                FBD.Description = text;
                FBD.RootFolder = Environment.SpecialFolder.Desktop;
                System.Threading.Thread oThread = new System.Threading.Thread(new System.Threading.ThreadStart(GetFolderName));
                oThread.ApartmentState = System.Threading.ApartmentState.STA;
                oThread.Start();
                while (!oThread.IsAlive) ;
                System.Threading.Thread.Sleep(1);
                oThread.Join();
                _strFileName = FBD.SelectedPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error : " + ex.ToString());
            }
            return _strFileName;
        }
        #endregion

        #region Getting File Name From  Dialog Box
        public static void GetFileName()
        {
            try
            {
                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);
                if (OFD.ShowDialog(oWindow) != DialogResult.OK)
                {
                    OFD.FileName = string.Empty;
                }
                oWindow = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Get Folder Name 
        public static void GetFolderName()
        {
            try
            {
                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);
                if (FBD.ShowDialog(oWindow) != DialogResult.OK)
                {
                    FBD.SelectedPath = string.Empty;
                }
                oWindow = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region Get Folder File Name 
       
        public static void GetFolderFileName()
        {
            try
            {

                IntPtr ptr = GetForegroundWindow();
                WindowWrapper oWindow = new WindowWrapper(ptr);
                if (SFD.ShowDialog(oWindow) != DialogResult.OK)
                {
                    SFD.FileName = string.Empty;
                }
                oWindow = null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        #endregion

        #region WINDOW WRAPPER CLASS FOR WINDOWS OBJECT
        public class WindowWrapper : System.Windows.Forms.IWin32Window
        {

            private IntPtr hwnd;

            public virtual IntPtr Handle
            {
                get
                {
                    return hwnd;
                }
            }

            public WindowWrapper(IntPtr handle)
            {
                hwnd = handle;
            }
        }
        #endregion



    }
}
