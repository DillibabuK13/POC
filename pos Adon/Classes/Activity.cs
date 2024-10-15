using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using General.Classes;
using pos_Adon;
using SAPbouiCOM;
using static General.Classes.@enum;

namespace pos_Adon.Classes
{
    class Activity:Connection
    {
        #region Variables  
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool SetForegroundWindow(IntPtr hWnd);
        SAPbouiCOM.Form oForm;
        SAPbouiCOM.Matrix oMatrix;
        SAPbouiCOM.EditText oEdit;
        SAPbouiCOM.ComboBox oCombo;
        SAPbouiCOM.Item oItem;
        SAPbouiCOM.Item oNewItem;
        SAPbouiCOM.Button obutton;
        public const string objType = "OCLG";
        SAPbouiCOM.DBDataSource oDbDataSource = null;
        const string formMenuUID = "OCLG";
        //const string formMenuUID1 = "SOF";
        public const string formTypeEx = "OCLG";
        clsCommon objclsComman = new clsCommon();
        StringBuilder sbQuery = new StringBuilder();

        public const string headerTable = "OCLG";
      
        #endregion

        public void MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region BeforeAction == true

                if (pVal.BeforeAction == true)
                {
                    try
                    {
                        //if (pVal.MenuUID == formMenuUID || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        //{
                        //    oForm = oApplication.Forms.ActiveForm;
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        //    {
                        //        //Record is directly added without validation
                        //        BubbleEvent = false;
                        //    }
                        //}
                  
                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = true : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region BeforeAction == false

                if (pVal.BeforeAction == false)
                {
                    try
                    {
                        if (pVal.MenuUID == formMenuUID || pVal.MenuUID == Convert.ToString((int)SAPMenuEnum.AddRecord))
                        {
                            LoadForm(pVal.MenuUID);
                        }
                       

                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }

        }


        public void ItemEvent(ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                #region Before_Action == true
                if (pVal.BeforeAction == true)
                {
                    try
                    {
                        oForm = oApplication.Forms.Item(pVal.FormUID);
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                        {
                           
                            if (pVal.ItemUID == "btnExcel")
                            {
                                ExcelApp(@"C:\Users\socius7\source\repos\pos Adon\Files\PosForm.xlsx");
                            }
                        }

                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion

                #region Before_Action == false
                if (pVal.BeforeAction == false)
                {
                    try
                    {
                        if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                        {
                            oForm = oApplication.Forms.Item(pVal.FormUID);
                        
                            oNewItem = oForm.Items.Add("btnExcel", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oItem = oForm.Items.Item("2"); 
                            obutton = oNewItem.Specific;
                            obutton.Caption = "Open Excel";
                            oNewItem.Top = oItem.Top;  
                            oNewItem.Left = oItem.Left + oItem.Width + 5;  
                            oNewItem.Height = oItem.Height;
                            oNewItem.Width = oItem.Width;
                            oNewItem.FromPane = 0;
                            oNewItem.ToPane = 0;

                        }


                    }
                    catch (Exception ex)
                    {
                        SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                        oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
                #endregion
            }
            catch (Exception ex)
            {
                SAPMain.logger.Error(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " Before action = false : " + ex.Message);
                oApplication.StatusBar.SetText(this.GetType().Name + " > " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }



        #region Methods
        public void LoadForm(string MenuID)
        {
            clsVariables.boolCFLSelected = false;

            if (MenuID == formMenuUID)
            {
                string formUID = "";
                objclsComman.LoadXML(MenuID, "", string.Empty, SAPbouiCOM.BoFormMode.fm_ADD_MODE);
                oForm = oApplication.Forms.ActiveForm;
                oForm.DataSources.UserDataSources.Add("Close", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 10);
                oForm.DataSources.UserDataSources.Item("Close").Value = "N";

            }


            oForm = oApplication.Forms.ActiveForm;
           
         
        }

        public void ExcelApp()
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            try
            {
                excelApp.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Add();

                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1]; // Access the first sheet

                worksheet.Name = "Sheet1";

                Thread.Sleep(1000);  // Adjust the delay as necessary

                // Find the Excel process by its name
                Process[] processes = Process.GetProcessesByName("EXCEL");
                if (processes.Length > 0)
                {
                    // Get the main window handle of Excel
                    IntPtr excelHandle = processes[0].MainWindowHandle;

                    if (excelHandle != IntPtr.Zero)
                    {
                        // Bring Excel to the front
                        bool success = SetForegroundWindow(excelHandle);
                        if (!success)
                        {
                            Console.WriteLine("Failed to bring Excel to the front.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Excel window handle is not available.");
                    }
                }
                else
                {
                    Console.WriteLine("Excel process not found.");
                }


                // Now, you can work with the sheet. For example, write something to a cell.
                // worksheet.Cells[1, 1].Value = "Hello, Excel!";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error opening Excel file: " + ex.Message);
            }
            finally
            {
                // Optionally, you can close the workbook after use
                // workbook.Close(false); // Set to true to save changes
                // excelApp.Quit();
            }
        }

        public void ExcelApp(string FilePath)
        {

            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();


            try
            {
                excelApp.Visible = true;
                Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(FilePath);

              
                Thread.Sleep(1000);  // Adjust the delay as necessary

                // Find the Excel process by its name
                Process[] processes = Process.GetProcessesByName("EXCEL");
                if (processes.Length > 0)
                {
                    // Get the main window handle of Excel
                    IntPtr excelHandle = processes[0].MainWindowHandle;

                    if (excelHandle != IntPtr.Zero)
                    {
                        // Bring Excel to the front
                        bool success = SetForegroundWindow(excelHandle);
                        if (!success)
                        {
                            Console.WriteLine("Failed to bring Excel to the front.");
                        }
                    }
                    else
                    {
                        Console.WriteLine("Excel window handle is not available.");
                    }
                }
                else
                {
                    Console.WriteLine("Excel process not found.");
                }


                // Now, you can work with the sheet. For example, write something to a cell.
                // worksheet.Cells[1, 1].Value = "Hello, Excel!";
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error opening Excel file: " + ex.Message);
            }
            finally
            {
                // Optionally, you can close the workbook after use
                // workbook.Close(false); // Set to true to save changes
                // excelApp.Quit();
            }
        }




        #endregion
    }
}
