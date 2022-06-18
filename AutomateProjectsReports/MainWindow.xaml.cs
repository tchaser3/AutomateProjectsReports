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
using DepartmentDLL;

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
        DepartmentClass TheDepartmentClass = new DepartmentClass();

        public MainWindow()
        {
            InitializeComponent();
        }
    }
}
