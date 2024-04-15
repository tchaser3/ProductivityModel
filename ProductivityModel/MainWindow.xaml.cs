/* Title:       Employee Productivity Model
 * Date:        7-22-21
 * Author:      Terry Holmes
 * 
 * Description: This is used to calculate the Productivity */
    
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excell = Microsoft.Office.Interop.Excel;
using NewEventLogDLL;
using NewEmployeeDLL;
using DataValidationDLL;
using DateSearchDLL;
using EmployeeLaborRateDLL;
using EmployeeProjectAssignmentDLL;
using EmployeePunchedHoursDLL;
using ProjectTaskDLL;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Win32;

namespace ProductivityModel
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        EmployeeClass TheEmployeeClass = new EmployeeClass();
        DataValidationClass TheDataValidationClass = new DataValidationClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        EmployeeLaborRateClass TheEmployeeLaborRateClass = new EmployeeLaborRateClass();
        EmployeeProjectAssignmentClass TheEmployeeProjectAssignmentClass = new EmployeeProjectAssignmentClass();
        EmployeePunchedHoursClass TheEmployeePunchedHoursClass = new EmployeePunchedHoursClass();
        ProjectTaskClass TheProjectTaskClass = new ProjectTaskClass();

        //Select Data
        ComboEmployeeDataSet TheComboEmployeeDataSet = new ComboEmployeeDataSet();
        FindEmployeeByEmployeeIDDataSet TheFindEmployeeByEmployeeIDDataSet = new FindEmployeeByEmployeeIDDataSet();
        FindProductionTaskForProductivityDataSet TheFindProductionTaskForProductivityDataSet = new FindProductionTaskForProductivityDataSet();
        EmployeeProductivityDataSet TheEmployeeProductivityDataSet = new EmployeeProductivityDataSet();
        EmployeeDayRateDataSet TheEmployeeDayRateDataSet = new EmployeeDayRateDataSet();
        FindEmployeePunchedHoursDataSet TheFindEmployeePunchedHoursDataSet = new FindEmployeePunchedHoursDataSet();
        FindEmployeeProjectAssignmentForProductivityDataSet TheFindEmployeeProjectAssignmentForProductivityDataSet = new FindEmployeeProjectAssignmentForProductivityDataSet();

        decimal gdecTechRate;
        decimal gdecProductionRate;
        int gintEmployeeID;
        string gstrDepartment;
        int gintCounter;
        int gintNumberOfRecords;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void expCloseProgram_Expanded(object sender, RoutedEventArgs e)
        {
            expCloseProgram.IsExpanded = false;
            TheMessagesClass.CloseTheProgram();
        }

        private void Grid_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            DragMove();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            ResetControls();
        }
        private void ResetControls()
        {
            cboSelectEmployee.Items.Clear();
            cboSelectEmployee.Items.Add("Select Employee");
            cboSelectEmployee.SelectedIndex = 0;

            gdecProductionRate = 0;
            gdecTechRate = 0;

            cboTechLevel.Items.Clear();
            cboTechLevel.Items.Add("Select Tech Level");
            cboTechLevel.Items.Add("Tech 1");
            cboTechLevel.Items.Add("Tech 2");
            cboTechLevel.Items.Add("Tech 3");
            cboTechLevel.SelectedIndex = 0;
            txtEnterLastName.Text = "";

            TheEmployeeProductivityDataSet.employeeproductivity.Rows.Clear();

            dgrProductivitiy.ItemsSource = TheEmployeeProductivityDataSet.employeeproductivity;
        }

        private void expResetWindow_Expanded(object sender, RoutedEventArgs e)
        {
            ResetControls();
        }

        private void txtEnterLastName_TextChanged(object sender, TextChangedEventArgs e)
        {
            string strLastName;
            int intCounter;
            int intNumberOfRecords;

            try
            {
                strLastName = txtEnterLastName.Text;

                if(strLastName.Length > 2)
                {
                    TheComboEmployeeDataSet = TheEmployeeClass.FillEmployeeComboBox(strLastName);
                    cboTechLevel.SelectedIndex = 0;

                    intNumberOfRecords = TheComboEmployeeDataSet.employees.Rows.Count;

                    if(intNumberOfRecords < 1)
                    {
                        TheMessagesClass.ErrorMessage("The Employee Was Not Found");

                        return;
                    }

                    cboSelectEmployee.Items.Clear();
                    cboSelectEmployee.Items.Add("Select Employee");

                    for(intCounter = 0; intCounter< intNumberOfRecords; intCounter++)
                    {
                        cboSelectEmployee.Items.Add(TheComboEmployeeDataSet.employees[intCounter].FullName);
                    }

                    cboSelectEmployee.SelectedIndex = 0;

                    TheEmployeeProductivityDataSet.employeeproductivity.Rows.Clear();

                    dgrProductivitiy.ItemsSource = TheEmployeeProductivityDataSet.employeeproductivity;
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Productivity Model // Main Window // Enter Last Name Text Box " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void cboSelectEmployee_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int intSelectedIndex;
            int intCounter;
            int intNumberOfRecords;
            bool blnItemFound;
            int intSecondCounter;
            decimal decFootage;
            DateTime datTransactionDate;
            DateTime datSecondTransactionDate;
            int intProjectID;
            int intWorkTaskID;
            decimal decTaskHours;

            try
            {
                intSelectedIndex = cboSelectEmployee.SelectedIndex - 1;

                if(intSelectedIndex > -1)
                {
                    gintEmployeeID = TheComboEmployeeDataSet.employees[intSelectedIndex].EmployeeID;
                    TheEmployeeProductivityDataSet.employeeproductivity.Rows.Clear();
                    cboTechLevel.SelectedIndex = 0;

                    TheFindEmployeeByEmployeeIDDataSet = TheEmployeeClass.FindEmployeeByEmployeeID(gintEmployeeID);

                    gstrDepartment = TheFindEmployeeByEmployeeIDDataSet.FindEmployeeByEmployeeID[0].Department;

                    TheFindProductionTaskForProductivityDataSet = TheProjectTaskClass.FindProductionTaskForProductivity(gintEmployeeID);

                    intNumberOfRecords = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity.Rows.Count;
                    gintCounter = 0;
                    gintNumberOfRecords = 0;

                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        blnItemFound = false;
                        datTransactionDate = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].TransactionDate;
                        decFootage = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].FootagePieces;
                        intProjectID = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].ProjectID;
                        intWorkTaskID = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].WorkTaskID;

                        TheFindEmployeeProjectAssignmentForProductivityDataSet = TheEmployeeProjectAssignmentClass.FindEmployeeProjectAssignmentForProductivity(gintEmployeeID, intProjectID, intWorkTaskID, datTransactionDate);

                        decTaskHours = TheFindEmployeeProjectAssignmentForProductivityDataSet.FindEmployeeProjectAssignmentForProductivity[0].TotalHours;

                        if (gintCounter > 0)
                        {
                            for(intSecondCounter = 0; intSecondCounter < gintCounter; intSecondCounter++)
                            {
                                if(datTransactionDate == TheEmployeeProductivityDataSet.employeeproductivity[intSecondCounter].TransactionDate)
                                {
                                    if(decFootage == TheEmployeeProductivityDataSet.employeeproductivity[intSecondCounter].FootagePieces)
                                    {
                                        blnItemFound = true;
                                    }
                                }
                            }
                        }

                        if(blnItemFound == false)
                        {
                            EmployeeProductivityDataSet.employeeproductivityRow NewProductivityRow = TheEmployeeProductivityDataSet.employeeproductivity.NewemployeeproductivityRow();

                            NewProductivityRow.FootagePieces = decFootage;
                            NewProductivityRow.ProductivityPrice = 0;
                            NewProductivityRow.ProjectID = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].AssignedProjectID;
                            NewProductivityRow.ProjectName = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].ProjectName;
                            NewProductivityRow.TransactionDate = datTransactionDate;
                            NewProductivityRow.WorkTask = TheFindProductionTaskForProductivityDataSet.FindProductionTaskForProductivity[intCounter].WorkTask;
                            NewProductivityRow.ProductivityRate = 0;
                            NewProductivityRow.TaskHours = decTaskHours;

                            TheEmployeeProductivityDataSet.employeeproductivity.Rows.Add(NewProductivityRow);
                            gintNumberOfRecords = gintCounter;
                            gintCounter++;
                        }
                        
                    }

                    dgrProductivitiy.ItemsSource = TheEmployeeProductivityDataSet.employeeproductivity;
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Productivity Model // Main Window // Select Employee Combo Box " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void cboTechLevel_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            //setting up the variables
            int intSelectedIndex;
            int intCounter;
            int intNumberOfRecords;
            decimal decTotalPieces;
            decimal decTotalPay;
            decimal decTotalHours;
            decimal decNormalHours;
            decimal decOverTimeHours;
            decimal decTotalProductionPay = 0;
            DateTime datPayDate;
            DateTime datStartDate;
            DateTime datTransactionDate;
            int intRecordsReturned;
            decimal decHourlyRate;

            try
            {
                intSelectedIndex = cboTechLevel.SelectedIndex;                

                if (intSelectedIndex > 0)
                {
                    TheEmployeeDayRateDataSet.employeedayrate.Rows.Clear();

                    datPayDate = DateTime.Now;
                    datPayDate = TheDateSearchClass.RemoveTime(datPayDate);
                    datPayDate = TheDateSearchClass.SubtractingDays(datPayDate, 180);

                    while (datPayDate.DayOfWeek != DayOfWeek.Sunday)
                    {
                        datPayDate = TheDateSearchClass.AddingDays(datPayDate, 1);
                    }

                    if (intSelectedIndex == 1)
                    {
                        gdecTechRate = 10;

                        if(gstrDepartment == "UNDERGROUND")
                        {
                            gdecProductionRate = Convert.ToDecimal(.20);
                        }
                        else
                        {
                            gdecProductionRate = Convert.ToDecimal(.10);
                        }
                    }
                    else if(intSelectedIndex == 2)
                    {
                        gdecTechRate = 15;

                        if (gstrDepartment == "UNDERGROUND")
                        {
                            gdecProductionRate = Convert.ToDecimal(.30);
                        }
                        else
                        {
                            gdecProductionRate = Convert.ToDecimal(.20);
                        }
                    }
                    else if(intSelectedIndex == 3)
                    {
                        gdecTechRate = 20;

                        if (gstrDepartment == "UNDERGROUND")
                        {
                            gdecProductionRate = Convert.ToDecimal(.50);
                        }
                        else
                        {
                            gdecProductionRate = Convert.ToDecimal(.30);
                        }
                    }

                    intNumberOfRecords = TheEmployeeProductivityDataSet.employeeproductivity.Rows.Count;

                    datStartDate = TheDateSearchClass.AddingDays(datPayDate, 1);
                    datPayDate = TheDateSearchClass.AddingDays(datPayDate, 7);
                    decTotalHours = 0;
                    decTotalPay = 0;
                    decNormalHours = 0;
                    decHourlyRate = 0;

                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        decTotalPieces = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].FootagePieces;
                        datTransactionDate = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].TransactionDate;

                        TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProductivityPrice = gdecProductionRate;
                        TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProductivityRate = decTotalPieces * gdecProductionRate;

                        if(datTransactionDate >= datStartDate)
                        {
                            if(datTransactionDate <= datPayDate)
                            {
                                decTotalProductionPay += TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProductivityRate;
                            }
                            else if(datTransactionDate > datPayDate)
                            {
                                TheFindEmployeePunchedHoursDataSet = TheEmployeePunchedHoursClass.FindEmployeePunchedHours(gintEmployeeID, datStartDate, datPayDate);

                                intRecordsReturned = TheFindEmployeePunchedHoursDataSet.FindEmployeePunchedHours.Rows.Count;

                                if(intRecordsReturned > 0)
                                {
                                    decTotalHours = TheFindEmployeePunchedHoursDataSet.FindEmployeePunchedHours[0].PunchedHours;
                                    decOverTimeHours = 0;

                                    if (decTotalHours > 40)
                                    {
                                        decOverTimeHours = decTotalHours - 40;
                                        decNormalHours = 40;
                                    }
                                    else if(decTotalHours <= 40)
                                    {
                                        decNormalHours = decTotalHours;
                                    }

                                    decTotalPay = ((decNormalHours) * gdecTechRate) + (decOverTimeHours * gdecTechRate * Convert.ToDecimal(1.5));
                                    decHourlyRate = (decNormalHours * Convert.ToDecimal(25)) + (decOverTimeHours * Convert.ToDecimal(25) * Convert.ToDecimal(1.5));

                                    EmployeeDayRateDataSet.employeedayrateRow NewEmployeeRate = TheEmployeeDayRateDataSet.employeedayrate.NewemployeedayrateRow();

                                    NewEmployeeRate.PayPeriodDate = datPayDate;
                                    NewEmployeeRate.PayRate = decTotalPay;
                                    NewEmployeeRate.ProductionRate = decTotalProductionPay;
                                    NewEmployeeRate.TotalProductionPay = decTotalProductionPay + decTotalPay;
                                    NewEmployeeRate.CurrentHourlyPay = decHourlyRate;
                                    NewEmployeeRate.Hours = decTotalHours;

                                    TheEmployeeDayRateDataSet.employeedayrate.Rows.Add(NewEmployeeRate);

                                    decTotalPay = 0;

                                    datStartDate = TheDateSearchClass.AddingDays(datStartDate, 7);
                                    datPayDate = TheDateSearchClass.AddingDays(datPayDate, 7);

                                    decTotalProductionPay = TheEmployeeProductivityDataSet.employeeproductivity[intCounter].ProductivityRate;

                                }

                            }
                        }
                        if(datPayDate < datTransactionDate)
                        {
                            //intCounter = intCounter - 1;
                            datStartDate = datTransactionDate;

                            if(datStartDate.DayOfWeek == DayOfWeek.Monday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 6);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Tuesday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 5);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Wednesday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 4);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Thursday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 3);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Friday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 2);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Saturday)
                            {
                                datPayDate = TheDateSearchClass.AddingDays(datStartDate, 1);
                            }
                            else if (datStartDate.DayOfWeek == DayOfWeek.Sunday)
                            {
                                datPayDate = datStartDate;
                            }
                        }
                    }

                    dgrProductivitiy.ItemsSource = TheEmployeeDayRateDataSet.employeedayrate;
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Productivity Model // Main Window // Tech Level Combo Box " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

        private void expExportToExcel_Expanded(object sender, RoutedEventArgs e)
        {
            expExportToExcel.IsExpanded = false;

            ExportHours();

            ExportProduction();
        }
        private void ExportProduction()
        {
            int intCounter;
            int intNumberOfRecords;
            string strAssetDecription;
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                intNumberOfRecords = TheEmployeeProductivityDataSet.employeeproductivity.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    worksheet = workbook.ActiveSheet;

                    worksheet.Name = "OpenOrders";

                    int cellRowIndex = 1;
                    int cellColumnIndex = 1;
                    intRowNumberOfRecords = TheEmployeeProductivityDataSet.employeeproductivity.Rows.Count;
                    intColumnNumberOfRecords = TheEmployeeProductivityDataSet.employeeproductivity.Columns.Count;

                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheEmployeeProductivityDataSet.employeeproductivity.Columns[intColumnCounter].ColumnName;

                        cellColumnIndex++;
                    }

                    cellRowIndex++;
                    cellColumnIndex = 1;

                    //Loop through each row and read value from each column. 
                    for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                    {
                        for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = TheEmployeeProductivityDataSet.employeeproductivity.Rows[intRowCounter][intColumnCounter].ToString();

                            cellColumnIndex++;
                        }
                        cellColumnIndex = 1;
                        cellRowIndex++;
                    }

                    //Getting the location and file name of the excel to save from user. 
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveDialog.FilterIndex = 1;

                    saveDialog.ShowDialog();

                    workbook.SaveAs(saveDialog.FileName);
                    TheMessagesClass.InformationMessage("Export Successful");

                    excel.Quit();
                }
            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Productivity Model // Main Window // Export Productivity " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void ExportHours()
        {
            int intCounter;
            int intNumberOfRecords;
            string strAssetDecription;
            int intRowCounter;
            int intRowNumberOfRecords;
            int intColumnCounter;
            int intColumnNumberOfRecords;

            // Creating a Excel object. 
            Microsoft.Office.Interop.Excel._Application excel = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

            try
            {
                intNumberOfRecords = TheEmployeeDayRateDataSet.employeedayrate.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    worksheet = workbook.ActiveSheet;

                    worksheet.Name = "OpenOrders";

                    int cellRowIndex = 1;
                    int cellColumnIndex = 1;
                    intRowNumberOfRecords = TheEmployeeDayRateDataSet.employeedayrate.Rows.Count;
                    intColumnNumberOfRecords = TheEmployeeDayRateDataSet.employeedayrate.Columns.Count;

                    for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                    {
                        worksheet.Cells[cellRowIndex, cellColumnIndex] = TheEmployeeDayRateDataSet.employeedayrate.Columns[intColumnCounter].ColumnName;

                        cellColumnIndex++;
                    }

                    cellRowIndex++;
                    cellColumnIndex = 1;

                    //Loop through each row and read value from each column. 
                    for (intRowCounter = 0; intRowCounter < intRowNumberOfRecords; intRowCounter++)
                    {
                        for (intColumnCounter = 0; intColumnCounter < intColumnNumberOfRecords; intColumnCounter++)
                        {
                            worksheet.Cells[cellRowIndex, cellColumnIndex] = TheEmployeeDayRateDataSet.employeedayrate.Rows[intRowCounter][intColumnCounter].ToString();

                            cellColumnIndex++;
                        }
                        cellColumnIndex = 1;
                        cellRowIndex++;
                    }

                    //Getting the location and file name of the excel to save from user. 
                    SaveFileDialog saveDialog = new SaveFileDialog();
                    saveDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                    saveDialog.FilterIndex = 1;

                    saveDialog.ShowDialog();

                    workbook.SaveAs(saveDialog.FileName);
                    TheMessagesClass.InformationMessage("Export Successful");

                    excel.Quit();
                }

            }
            catch (Exception Ex)
            {
                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Productivity Model // Main Window // Export Hours " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }

    }
}
