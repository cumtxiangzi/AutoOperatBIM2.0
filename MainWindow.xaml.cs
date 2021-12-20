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
using Selenium;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Diagnostics;
using System.Windows.Interop;
using System.Runtime.InteropServices;


namespace AutoOperateBIM2._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        AutoOperatClass operat = new AutoOperatClass();
        public MainWindow()
        {
            InitializeComponent();
        }
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var desktopWorkingArea = SystemParameters.WorkArea;
            Left = desktopWorkingArea.Right - Width;
            Top = desktopWorkingArea.Bottom - Height;
            Topmost = true;
        }
        private void MainForm_Closed(object sender, EventArgs e)
        {
            operat.stopRun();
        }
        private void BIM2Button_Click(object sender, RoutedEventArgs e)
        {
            //AutoOperatBIMPlateForm2("https://bim.sinoma-tianjin.com/", "1739", "110122227");
        }
        private void BIM1Button_Click(object sender, RoutedEventArgs e)
        {

            operat.startRun();
            //BIMPlateForm1("http://bim.sinoma-tianjin.com:8080/Login", "1739", "110122227", "1007", "919");
        }


    }
}
