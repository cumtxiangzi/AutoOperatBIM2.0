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

        //BIM1.0平台操作
        private void BIM0Button_Click(object sender, RoutedEventArgs e)
        {
            operat.startRun(0);
        }
        private void BIM1Button_Click(object sender, RoutedEventArgs e)
        {
            operat.startRun(1);
        }

        private void BIM2Button_Click(object sender, RoutedEventArgs e)
        {
            operat.stopRun();
        }

        //BIM2.0平台操作
        private void BIM20Button_Click(object sender, RoutedEventArgs e)
        {
            operat.startRun(20);
        }

        private void BIM21Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BIM22Button_Click(object sender, RoutedEventArgs e)
        {
            operat.stopRun();
        }

        //BIM3.0平台操作
        private void BIM30Button_Click(object sender, RoutedEventArgs e)
        {
            operat.startRun(30);
        }

        private void BIM31Button_Click(object sender, RoutedEventArgs e)
        {

        }

        private void BIM32Button_Click(object sender, RoutedEventArgs e)
        {
            operat.stopRun();
        }
    }
}
