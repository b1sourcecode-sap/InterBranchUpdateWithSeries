using System;
using System.Data.SqlClient;

namespace StockTransferInterBranch
{
    /// <summary>
    /// Summary description for clsStart.
    /// </summary>
    public class Start
    {
        static void Main()
        {
            try
            {
                // Company connection and other initilizations.
                Utilities.Application = EventListener.getEventListener();

                bool oFlag = true;

                if (true)
                {
                    GC.Collect();
                    // initialize Database 
                    Database.InitializeDatabase();
                    Utilities.LoadMenus(Constants.Menus.MENUS_ADD);
                    Utilities.ShowSucessMessage("Add-On is Connected Successfully");
                    System.Windows.Forms.Application.Run();
                }
            }
            catch (Exception ex)
            {
                Utilities.ShowErrorMessage(ex.Message);
            }
        }
    }
}
