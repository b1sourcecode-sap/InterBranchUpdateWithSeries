using System;
using SAPbobsCOM;
using SAPbouiCOM;
using System.IO;
using System.Data;
using System.Collections;
using System.Windows.Forms;
//using StockTransferIterBranch.Business_Logics.Custom_Form;

namespace StockTransferInterBranch
{
    /// <summary>
    /// public sealed class EventListener
    /// </summary>
    public sealed class EventListener
    {
        #region FORM VARIABLES DECLARATION
        private static EventListener    _Listener;
        private SAPbouiCOM.Application  _Application;
        private SAPbobsCOM.Company      _Company;
        private Base                    _Object;
        private Hashtable               _Collection;
        private Hashtable               _LookCollection;
        private string                  _FormUID;
        #endregion

        #region CONSTRUCTOR

        private EventListener()
        {
            this.Connect();

            this._Application.MenuEvent       += new _IApplicationEvents_MenuEventEventHandler(SboApplication_MenuEvent);
            this._Application.ItemEvent       += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SboApplication_ItemEvent);
            this._Application.AppEvent        += new _IApplicationEvents_AppEventEventHandler(SboApplication_AppEvent);
            this._Application.RightClickEvent += new _IApplicationEvents_RightClickEventEventHandler(SboApplication_RightClickEvent);
            this._Application.FormDataEvent   += new _IApplicationEvents_FormDataEventEventHandler(SboApplication_FormDataEvent);

            this._Collection                      = new Hashtable(10, (float)0.5);
            this._LookCollection                  = new Hashtable(10, (float)0.5);
        }

        public static EventListener getEventListener()
        {
            if (_Listener == null)
                _Listener = new EventListener();


            return _Listener;
        }

        #endregion

        #region PROPERTIES

        public SAPbobsCOM.Company Company
        {
            get { return this._Company; }
        }

        public SAPbouiCOM.Application SBO_Application
        {
            get { return this._Application; }
        }

        public Hashtable Collection
        {
            get { return this._Collection; }
        }

        public Hashtable LookUpCollection
        {
            get { return this._LookCollection; }
        }

        #endregion

        #region COMPANY CONNECT

        private void Connect()
        {
            this.SetApplication();
            this.ConnectToCompany();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi oSboGuiApi;
            string               oConnectionString;
            try
            {
                if (Environment.GetCommandLineArgs().Length > 0)
                {
                    //oSboGuiApi = new SAPbouiCOM.SboGuiApiClass();
                    oSboGuiApi = new SAPbouiCOM.SboGuiApi();
                    oConnectionString = Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                    oSboGuiApi.Connect(oConnectionString);
                    this._Application = oSboGuiApi.GetApplication(-1);
                    //this._Company = new SAPbobsCOM.CompanyClass();
                    this._Company = new SAPbobsCOM.Company();
                }
                else
                {
                    throw new Exception("Connection string missing.");
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void ConnectToCompany()
        {
            string oCookie, oConnectionContext;
            try
            {
                oCookie = this._Company.GetContextCookie();
                oConnectionContext = this._Application.Company.GetConnectionContext(oCookie);
                this._Company.SetSboLoginContext(oConnectionContext);

                if (this._Company.Connect() != 0)
                {
                    _Application.StatusBar.SetText("Add-On is not Connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    throw new Exception(_Company.GetLastErrorDescription());
                }
                else
                {
                    _Application.StatusBar.SetText("Add-On is Connecting...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        #region EVENT HANDLERS

        #region MENU EVENT

        private void SboApplication_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (!pVal.BeforeAction)
                {

                    switch (pVal.MenuUID)
                    {
                        //case Constants.User_Menus.MENU_DeliveryCreation:
                        //    this._Object = new ClsDeliveryCreation();
                        //    Utilities.LoadForm(ref this._Object, Constants.Forms.DeliveryCreation);
                        //    this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                        //    ((ClsDeliveryCreation)_Object).FormDefault();
                        //    break;

                        // System Menu's
                        case Constants.System_Menus.mnu_ADD:
                        case Constants.System_Menus.mnu_FIND:
                        case Constants.System_Menus.mnu_FIRST:
                        case Constants.System_Menus.mnu_LAST:
                        case Constants.System_Menus.mnu_NEXT:
                        case Constants.System_Menus.mnu_PREVIOUS:
                        case Constants.System_Menus.mnu_ADD_ROW:
                        case Constants.System_Menus.mnu_DELETE_ROW:
                        case Constants.System_Menus.mnu_SALES_ORDER:
                        case Constants.System_Menus.mnu_OUTGOING_PAYMENTS:
                        case Constants.System_Menus.mnu_GL_ACCOUNT_DETERMINATION:
                        case Constants.System_Menus.mnu_Duplicate:
                        case Constants.System_Menus.mnu_Cancelled:
                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;
                    }
                }
                else
                {
                    switch (pVal.MenuUID)
                    {
                        //case Constants.System_Menus.mnu_GROSS_PROFIT:
                        //    BubbleEvent = false;
                        //    break;
                        case Constants.System_Menus.mnu_ROW_DETAILS:
                            BubbleEvent = false;
                            break;

                        case Constants.System_Menus.mnu_DELETE_ROW:
                       
                        case Constants.System_Menus.mnu_Duplicate:
                            if (this._Collection.Contains(this._FormUID))
                            {
                                this._Object = (Base)this._Collection[this._FormUID];
                                ((Base)this._Object).Menu_Event(ref pVal, ref BubbleEvent);
                            }
                            break;

                        case Constants.User_Menus.MENU_InterBranchTransferReq:
                            this._Object = new ClsInterBranchTransferRequest();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.InterBranchTransferRequest);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsInterBranchTransferRequest)_Object).FormDefault();
                            break;

                        case Constants.User_Menus.MENU_InterBranchTransfer:
                            this._Object = new ClsInterBranchTransfer();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.InterBranchTransfer);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsInterBranchTransfer)_Object).FormDefault("");
                            break;

                        case Constants.User_Menus.MENU_LGApproval:
                            this._Object = new ClsApprovalTemplate();
                            Utilities.LoadForm(ref this._Object, Constants.Forms.ApprovalTemplate);
                            this._Object.Menu_Event(ref pVal, ref BubbleEvent);
                            ((ClsApprovalTemplate)_Object).FormDefault();
                            break;
                      

                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        internal void MessageBox(string v)
        {
            throw new NotImplementedException();
        }

        #endregion

        #region ITEM EVENT

        private void SboApplication_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                this._FormUID = FormUID;
               
                //Utilities.HideSystemMessege(ref pVal);
                if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD) //
                {
                    switch (pVal.FormType)
                    {
                        case Constants.System_Forms.GOODS_ISSUE:
                            this._Object = new ClsGoodsIssue();
                            _Collection.Add(FormUID, this._Object);
                            ((ClsGoodsIssue)this._Object).Form = _Application.Forms.Item(FormUID);
                            ((ClsGoodsIssue)this._Object).FormDefault();
                            break;

                    }
                }

                if (this._Collection.Contains(FormUID))
                {
                    this._Object = (StockTransferInterBranch.Base)this._Collection[FormUID];

                    if (pVal.BeforeAction && this._Object.IsLookUpOpen)
                    {
                        BubbleEvent = false;
                        _Application.Forms.Item(this._Object.LookUpUID).Select();
                    }
                    else
                    {
                        this._Object.Item_Event(FormUID, ref pVal, ref BubbleEvent);
                    }

                    if (!pVal.BeforeAction && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_UNLOAD)
                    {
                        if (_LookCollection.ContainsKey(FormUID))
                        {
                            this._Object = (StockTransferInterBranch.Base)_Collection[_LookCollection[FormUID]];
                            _LookCollection.Remove(FormUID);
                            this._Object.IsLookUpOpen = false;
                        }
                        this._Object.Dispose();
                        _Collection.Remove(FormUID);
                        System.GC.Collect();
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region RIGHT CLICK EVENT

        private void SboApplication_RightClickEvent(ref ContextMenuInfo eventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            if (this._Collection.Contains(eventInfo.FormUID))
            {
                this._Object = (StockTransferInterBranch.Base)this._Collection[eventInfo.FormUID];
                this._Object.RightClick_Event(ref eventInfo, ref BubbleEvent);
            }
        }

        #endregion

        #region APPLICATION EVENT
        private void SboApplication_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        Utilities.LoadMenus(Constants.Menus.REMOVE_MENUS);
                        Utilities.ShowWarningMessage("Add-On is Disconnected Successfully ... ");
                        CloseApplication();
                        break;
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }
        #endregion

        #region FORM DATA EVENT
        private void SboApplication_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (BusinessObjectInfo.FormUID != null)
                {
                    if (this._Collection.Contains(BusinessObjectInfo.FormUID))
                    {
                        this._Object = (StockTransferInterBranch.Base)this._Collection[BusinessObjectInfo.FormUID];
                        this._Object.FormData_Event(ref BusinessObjectInfo, ref BubbleEvent);
                    }
                }
            }
            catch (Exception ex)
            {
                Utilities.Message(ex.Message, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }
        #endregion
        #endregion

        #region CLOSE APPLICATION
        private void CloseApplication()
        {
            try
            {
                if (this._Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Application);
                    _Application = null;
                }
                if (this._Company != null)
                {
                    if (this._Company.Connected)
                        this._Company.Disconnect();

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(_Company);
                    _Company = null;
                }
                this._Collection = null;
                this._LookCollection = null;
                _Listener = null;
                System.Windows.Forms.Application.Exit();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                // clsUtilities.oApplication = null;
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }
        }
        #endregion
    }
}
