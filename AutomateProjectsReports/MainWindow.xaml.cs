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
using DateSearchDLL;
using DesignProjectsDLL;
using ProductionProjectDLL;
using NewEventLogDLL;
using WorkOrderDLL;

namespace AutomateProjectsReports
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //setting up the classes
        WPFMessagesClass TheMessagesClass = new WPFMessagesClass();
        SendEmailClass TheSendEmailClass = new SendEmailClass();
        DateSearchClass TheDateSearchClass = new DateSearchClass();
        DesignProjectsClass TheDesignProjectsClass = new DesignProjectsClass();
        ProductionProjectClass TheProductionProjectClass = new ProductionProjectClass();
        EventLogClass TheEventLogClass = new EventLogClass();
        WorkOrderClass TheWorkOrderClass = new WorkOrderClass();

        //setting up the data
        FindProductionProjectsEnteredNewStatusDataSet TheFindProductionProjectsEnteredNewStatusDataSet = new FindProductionProjectsEnteredNewStatusDataSet();
        FindWorkOrderStatusSortedDataSet TheFindWorkOrderStatusSortedDataSet = new FindWorkOrderStatusSortedDataSet();
        FindProductionProjectsNotUpdatedDataSet TheFindProductionProjectsNotUpdatedDataSet = new FindProductionProjectsNotUpdatedDataSet();

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            int intCounter;
            int intNumberOfRecords;
            int intStatusID;
            int intProjectCounter;
            int intProjectNumberOfRecords;
            DateTime datTransactionDate = DateTime.Now;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strMessage = "";
            string strStatus;

            try
            {
                

                //setting up the time
                datTransactionDate = TheDateSearchClass.RemoveTime(datTransactionDate);
                datTransactionDate = TheDateSearchClass.SubtractingDays(datTransactionDate, 1);

                if(datTransactionDate.DayOfWeek == DayOfWeek.Sunday)
                {
                    datTransactionDate = TheDateSearchClass.SubtractingDays(datTransactionDate, 2);
                }
                if (datTransactionDate.DayOfWeek == DayOfWeek.Saturday)
                {
                    datTransactionDate = TheDateSearchClass.SubtractingDays(datTransactionDate, 1);
                }

                if(datTransactionDate.DayOfWeek == DayOfWeek.Monday)
                {
                    //running not updated Method
                    ProjectsNotUpdated();
                }

                TheFindWorkOrderStatusSortedDataSet = TheWorkOrderClass.FindWorkOrderStatusSorted();

                intNumberOfRecords = TheFindWorkOrderStatusSortedDataSet.FindWorkOrderStatusSorted.Rows.Count;

                if(intNumberOfRecords > 0)
                {
                    for(intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        intStatusID = TheFindWorkOrderStatusSortedDataSet.FindWorkOrderStatusSorted[intCounter].StatusID;
                        strStatus = TheFindWorkOrderStatusSortedDataSet.FindWorkOrderStatusSorted[intCounter].WorkOrderStatus;

                        strHeader = "Projects Changed Status to " + strStatus;
                        strMessage = "<h1>Projects Changed Status to " + strStatus + "</h1>";

                        TheFindProductionProjectsEnteredNewStatusDataSet = TheProductionProjectClass.FindProductionProjectsEnteredNewStatus(intStatusID, datTransactionDate);

                        intProjectNumberOfRecords = TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus.Rows.Count;

                        if(intProjectNumberOfRecords > 0)
                        {
                            strMessage += "<table>";
                            strMessage += "<tr>";
                            strMessage += "<td><b>Customer Project ID</b></td>";
                            strMessage += "<td><b>Assigned Project ID</b></td>";
                            strMessage += "<td><b>Project Name</b></td>";
                            strMessage += "<td><b>Work Order Status</b></td>";
                            strMessage += "<td><b>Department</b></td>";
                            strMessage += "<td><b>Assigned Office</b></td>";
                            strMessage += "<td><b>ECD Date</b></td>";
                            strMessage += "</tr>";

                            for (intProjectCounter = 0; intProjectCounter < intProjectNumberOfRecords; intProjectCounter++)
                            {
                                strMessage += "<tr>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].CustomerAssignedID + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].AssignedProjectID + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].ProjectName + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].WorkOrderStatus + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].Department + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].AssignedOffice + "</td>";
                                strMessage += "<td>" + TheFindProductionProjectsEnteredNewStatusDataSet.FindProductionProjectsEnteredNewStatus[intProjectCounter].ECDDate + "</td>";
                                strMessage += "</tr>";
                            }

                            strMessage += "</table>";

                            TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);
                        }
                    }
                }

                Application.Current.Shutdown();
            }
            catch (Exception Ex)
            {
                TheSendEmailClass.SendEventLog("Automate Project Reports // Main Window // Window Loaded " + Ex.ToString());

                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Automate Project Reports // Main Window // Window Loaded " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }
        }
        private void ProjectsNotUpdated()
        {
            int intCounter;
            int intNumberOfRecords;
            string strEmailAddress = "ERPProjectReports@bluejaycommunications.com";
            string strHeader;
            string strMessage = "";

            try
            {
                TheFindProductionProjectsNotUpdatedDataSet = TheProductionProjectClass.FindProductionProjectsNotUpdated();

                intNumberOfRecords = TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated.Rows.Count;

                if (intNumberOfRecords > 0)
                {
                    strHeader = "Projects That Have Not Been Updated in 15 Dates";
                    strMessage = "<h1>Projects That Have Not Been Updated in 15 Dates</h1>";

                    strMessage += "<table>";
                    strMessage += "<tr>";
                    strMessage += "<td><b>Customer Project ID</b></td>";
                    strMessage += "<td><b>Assigned Project ID</b></td>";
                    strMessage += "<td><b>Project Name</b></td>";
                    strMessage += "<td><b>Business Address</b></td>";
                    strMessage += "<td><b>ECD Date</b></td>";
                    strMessage += "<td><b>Work Order Status</b></td>";
                    strMessage += "<td><b>Assigned Office</b></td>";
                    strMessage += "<td><b>Department</b></td>";
                    strMessage += "<td><b>Status Change Date</b></td>";

                    strMessage += "</tr>";

                    for (intCounter = 0; intCounter < intNumberOfRecords; intCounter++)
                    {
                        strMessage += "<tr>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].CustomerAssignedID + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].AssignedProjectID + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].ProjectName + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].BusinessAddress + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].ECDDate + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].WorkOrderStatus + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].AssignedOffice + "</td>";
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].Department + "</td>";                        
                        strMessage += "<td>" + TheFindProductionProjectsNotUpdatedDataSet.FindProductionProjectsNotUpdated[intCounter].StatusChangeDate + "</td>";
                        strMessage += "</tr>";
                    }

                    strMessage += "</table>";

                    TheSendEmailClass.SendEmail(strEmailAddress, strHeader, strMessage);
                }
            }
            catch (Exception Ex)
            {
                TheSendEmailClass.SendEventLog("Automate Project Reports // Main Window // Projects Not Updated " + Ex.ToString());

                TheEventLogClass.InsertEventLogEntry(DateTime.Now, "Automate Project Reports // Main Window // Projects Not Updated " + Ex.Message);

                TheMessagesClass.ErrorMessage(Ex.ToString());
            }

        }
    }
}
