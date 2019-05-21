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
using System.Windows.Shapes;
using Microsoft.Win32;
using System.IO;

namespace TaskManager
{
    /// <summary>
    /// Setting.xaml 的交互逻辑
    /// </summary>
    public partial class Setting : Window
    {
        public bool isNeedToUpdate=false;
        private string gSourcePath="";
        public string SourcePath
        {
            get
            {
                return gSourcePath;
            }
            set
            {
                if (gSourcePath == "" || gSourcePath != value)
                {
                    gSourcePath = value;
                    txt_SourcePath.Text = value;
                    isNeedToUpdate = true;
                }
            }
        }

        private string gChoosenSheet="";
        public string ChoosenSheet
        {
            get
            {
                return gChoosenSheet;
            }
            set
            {
                if (gChoosenSheet == "" || gChoosenSheet != value)
                {
                    gChoosenSheet = value;
                    txt_Sheet.Text = value;
                    isNeedToUpdate = true;
                }
            }
        }

        public bool g_isLoadHistory = true;
        public bool isLoadHistory
        {
            get
            {
                return g_isLoadHistory;
            }
            set
            {
                g_isLoadHistory = value;
                isNeedToUpdate = true;
                if (value)
                {
                    btn_isLoadHistory.Content = "✔";
                }
                else
                {
                    btn_isLoadHistory.Content = " ";
                }
            }
        }

        public string ProgramPath = System.AppDomain.CurrentDomain.BaseDirectory;

        public Setting()
        {
            InitializeComponent();
            Expend.LoadImage(btn_Open, @"Texture\pic_File.png");
            //INI
            if (!File.Exists("Setting.ini") || !File.Exists(INIHelper.ContentValue("ALL", "ExcelPath", ProgramPath + "Setting.ini")))
            {
                MessageBox.Show("这或许是你第一次使用本产品，请从计算机中打开一个Excel表格作为你的计划表");
                StreamWriter writer = new StreamWriter("Setting.ini");
                writer.WriteLine("[ALL]");
                writer.WriteLine("ExcelPath=");
                writer.WriteLine("Sheet=");
                writer.Close();
                OpenFileDialog open = new OpenFileDialog();
                open.ShowDialog();
                while (open.FileName == "")
                {
                    open.ShowDialog();
                }
                SourcePath = open.FileName;
                INIHelper.WritePrivateProfileString("ALL", "ExcelPath", SourcePath,ProgramPath+ "Setting.ini");
                INIHelper.WritePrivateProfileString("ALL", "Sheet", ChoosenSheet, ProgramPath + "Setting.ini");
            }
            else
            {
                LoadFromINI();
            }
        }

        private void LoadFromINI()
        {
            SourcePath = INIHelper.ContentValue("ALL", "ExcelPath", ProgramPath + "Setting.ini");
            ChoosenSheet = INIHelper.ContentValue("ALL", "Sheet", ProgramPath + "Setting.ini");
            isLoadHistory = INIHelper.ContentValue("ALL", "IsLoadHistory", ProgramPath + "Setting.ini") == "True";
        }

        private void btn_Confirm_Click(object sender, RoutedEventArgs e)
        {
            INIHelper.WritePrivateProfileString("ALL", "ExcelPath", SourcePath, ProgramPath + "Setting.ini");
            INIHelper.WritePrivateProfileString("ALL", "Sheet", ChoosenSheet, ProgramPath + "Setting.ini");
            INIHelper.WritePrivateProfileString("ALL", "IsLoadHistory", isLoadHistory.ToString(), ProgramPath + "Setting.ini");
            this.Hide();
        }

        private void btn_Cannel_Click(object sender, RoutedEventArgs e)
        {
            //回档
            isNeedToUpdate = false;
            LoadFromINI();
            this.Hide();
        }

        private void btn_MiniSize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btn_Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.ShowDialog();
            if (open.FileName != "")
            {
                SourcePath = open.FileName;
            }
        }

        private void btn_Close_Click(object sender, RoutedEventArgs e)
        {
            btn_Cannel_Click(sender, e);
        }

        private void txt_SourcePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            gSourcePath = txt_SourcePath.Text;
        }

        private void Window_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            this.DragMove();
        }

        private void txt_Sheet_TextChanged(object sender, TextChangedEventArgs e)
        {
            gChoosenSheet = txt_Sheet.Text;
        }

        private void Btn_isLoadHistory_Click(object sender, RoutedEventArgs e)
        {
            isLoadHistory = !isLoadHistory;
        }

        private void Btn_ClearHistory_Click(object sender, RoutedEventArgs e)
        {
            MessageBoxResult res =MessageBox.Show("确定删除？这会清除你所有的历史记录！","清除确认",button:MessageBoxButton.YesNo);
            if (res == MessageBoxResult.Yes)
            {
                using (StreamWriter writer = new StreamWriter("Data.xml"))
                {
                    writer.WriteLine("<Datas>");
                    writer.WriteLine("</Datas>");
                }
            }
        }
    }
}
