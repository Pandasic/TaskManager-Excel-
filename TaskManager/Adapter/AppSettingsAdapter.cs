using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace TaskManager.Adapter
{
    public class AppSettingsAdapter
    {
        private string pathAppSetting = "";

        public string ExcelPath
        {
            get
            {
                return INIHelper.ContentValue("ALL", "ExcelPath", this.pathAppSetting);
            }
            set
            {
                if (File.Exists(value) && value.EndsWith(".xlsx"))
                {
                    INIHelper.WritePrivateProfileString("ALL", "ExcelPath", value, this.pathAppSetting);
                }
            }
        }

        public string Sheet
        {
            get
            {
                return INIHelper.ContentValue("ALL", "Sheet", this.pathAppSetting);
            }
            set
            {
                INIHelper.WritePrivateProfileString("ALL", "Sheet", value, this.pathAppSetting);
            }
        }

        public bool isLoadHistory 
        {
            get
            {
                return INIHelper.ContentValue("ALL", "IsLoadHistory",  this.pathAppSetting) == "True";
            }
            set
            {
                if (value)
                {
                    INIHelper.WritePrivateProfileString("ALL", "IsLoadHistory", "True", this.pathAppSetting);
                }
                else
                {
                    INIHelper.WritePrivateProfileString("ALL", "IsLoadHistory", "False", this.pathAppSetting);
                }
            }
        }
        public AppSettingsAdapter()
        {
            //获取程序而不是快捷方式所在目录
            pathAppSetting = System.AppDomain.CurrentDomain.BaseDirectory + "/Setting.ini";

            string ProgramPath = System.AppDomain.CurrentDomain.BaseDirectory;
            if (!File.Exists(this.pathAppSetting))
            {
                StreamWriter writer = new StreamWriter(this.pathAppSetting);
                writer.WriteLine("[ALL]");
                writer.WriteLine("ExcelPath=");
                writer.WriteLine("Sheet=");
                writer.Close();
            }

            if (!File.Exists(INIHelper.ContentValue("ALL", "ExcelPath", this.pathAppSetting)))
            {
                MessageBox.Show("这或许是你第一次使用本产品，请从计算机中打开一个Excel表格作为你的计划表");
                setExcelPathInExplorer();
                this.setExcelPathInExplorer();
                this.Sheet = "Sheet1";
            }
        }

        public void setExcelPathInExplorer()
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = ".xlsx";
            open.ShowDialog();
            while (open.FileName == "")
            {
                open.ShowDialog();
            }
            string SourcePath = open.FileName;
            ExcelPath = SourcePath;
        }
    }
}
