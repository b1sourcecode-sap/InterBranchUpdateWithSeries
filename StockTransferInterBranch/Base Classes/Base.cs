using System;

namespace StockTransferInterBranch
{
	/// <summary>
	/// Summary description for Base.
	/// </summary>
    
	public class Base :IDisposable
    {
        #region CLASS LEVEL VARIABLE DECLARATION
        protected Base              _Object;
		protected string            _FormUID;
		protected SAPbouiCOM.Form   _Form;
		protected bool              _LookUpOpen;
		protected string            _LookUpFrmUID;
        protected Constants.CFL_Event _CFLEvent;
        #endregion

        #region CONSTRUCTOR & DISTRUCTOR
        public Base()
		{
			this._Object = null;
			this._Form = null;
			this._LookUpOpen = false;			
		}

		~Base()
		{}
		#endregion

		#region VIRTUAL FUNCTIONS
		public virtual void Menu_Event(ref SAPbouiCOM.MenuEvent pVal, ref bool oBubbleEvent)
		{}

		public virtual void Item_Event(string oFormUID , ref SAPbouiCOM.ItemEvent pVal, ref bool oBubbleEvent)
		{}

		public virtual void RightClick_Event(ref SAPbouiCOM.ContextMenuInfo oEventInfo, ref bool oBubbleEvent)
		{}
        public virtual void FormData_Event(ref  SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool oBubbleEvent)
        { }

        #region Custom CHOOSE FROM LIST EVENT
        public virtual void ChooseFromList_Event()
        { }
        #endregion
		#endregion
		
		#region PROPERTIES  
		public SAPbouiCOM.Form Form
		{
			get { return this._Form; }			
			set { this._Form = value; } 
		}

		public string FormUID
		{
			get { return this._FormUID; }
			set { this._FormUID = value; }
		}

		public bool IsLookUpOpen
		{
			get { return this._LookUpOpen; }
			set { this._LookUpOpen = value; }
		}

        public string LookUpUID
        {
            get { return this._LookUpFrmUID; }
            set { this._LookUpFrmUID = value; }
        }

        public Constants.CFL_Event CFL_Event
        {
            get { return this._CFLEvent; }
            set { this._CFLEvent = value; }
        }
		#endregion

		#region OPEN CHILD FORM
		protected void OpenChildForm(string oXML,  bool oModal)
		{	
			if( oXML != "" )
			{
				Utilities.LoadForm(ref this._Object, oXML);
				this._LookUpFrmUID = this._Object.FormUID;
				this.IsLookUpOpen = oModal;

				if( oModal )
					Utilities.Application.LookUpCollection.Add(this._Object.FormUID,this._Form.UniqueID);
			}
		}
		#endregion

		#region IDISPOSABLE MEMBERS

		virtual public void Dispose()
		{
			// TODO:  Add BaseModule.Dispose implementation
		}

		#endregion
        
        #region OPEN SYSTEM FORM AS CHILD FORM
        protected void OpenSystemFormAsChildForm(string oMenuUID, bool oModal)
		{
			if( oMenuUID != string.Empty )
				Utilities.Application.SBO_Application.ActivateMenuItem( oMenuUID );

			this._Object = (Base)Utilities.Application.Collection[Utilities.Application.SBO_Application.Forms.ActiveForm.UniqueID];
			this._Object.FormUID = Utilities.Application.SBO_Application.Forms.ActiveForm.UniqueID;
			this._LookUpFrmUID = this._Object.FormUID;
			this.IsLookUpOpen = oModal;

			if( oModal )
				Utilities.Application.LookUpCollection.Add(this._Object.FormUID,this._Form.UniqueID);
        }
		#endregion

		#region OPEN CHILD FORM
		protected void OpenChildForm(string oXML, ref SAPbouiCOM.ItemEvent oPVal, bool oModal)
		{
			if (oXML != "")
			{
				Utilities.LoadForm(ref this._Object, oXML);
				this._LookUpFrmUID = this._Object.FormUID;
				this.IsLookUpOpen = oModal;

				if (oModal)
					Utilities.Application.LookUpCollection.Add(this._Object.FormUID, this._Form.UniqueID);
			}
		}
		#endregion
	}
}
