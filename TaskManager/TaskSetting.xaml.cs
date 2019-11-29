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
using System.Xml;
using System.IO;
using Microsoft.Win32;
using System.Text.RegularExpressions;
using System.Collections;
using System.Diagnostics;

namespace TaskManager
{
    /// <summary>
    /// TaskSetting.xaml 的交互逻辑
    /// </summary>
    public partial class TaskSetting : Window
    {
        XmlDocument xmlDoc = new XmlDocument ();

        static TaskSetting()
        {

        }

        public TaskSetting(TaskInformation Task) 
        {
            InitializeComponent();
            Expend.LoadImage(btn_Open, @"Texture\pic_File.png");

            this.Title = string.Format("{0} - TaskSettings",Task.Name);

            xmlDoc.Load("TaskSetting.xml");
            XmlNode root = xmlDoc.FirstChild;

            txt_Name.Text = Task.Name;
            XmlElement ele = null;
            foreach (XmlElement xe in root.ChildNodes)
            {
                if (Task.Name.Contains(xe.GetAttribute("Name")))
                {
                    if (ele == null || xe.GetAttribute("Name").Length > ele.GetAttribute("Name").Length)
                        ele = xe;
                }
            }
            if (ele == null)//如果不在XML表
            {
                
                ele = xmlDoc.CreateElement("SETTING");
                ele.SetAttribute("Name",Task.Name);
                ele.SetAttribute("AimSource", "");
                root.AppendChild(ele);
            }
            else
            {
                txt_Name.Text = ele.GetAttribute("Name");
                txt_AimSource.Text = ele.GetAttribute("AimSource");
            }
        }

        public TaskSetting(string TaskName) : this(new TaskInformation() { Name = TaskName})
        {

        }

        private void btn_Confirm_Click(object sender, RoutedEventArgs e)
        {
            XmlNode root = xmlDoc.FirstChild;
            TaskInformation task = new TaskInformation() { Name = txt_Name.Text };
            XmlElement ele = (XmlElement)root.SelectSingleNode(string.Format("//SETTING[@Name='{0}']", txt_Name.Text));
            if (ele == null)//如果不在XML表
            {
                ele = xmlDoc.CreateElement("SETTING");
            }
            ele.SetAttribute("Name", txt_Name.Text);
            ele.SetAttribute("AimSource",txt_AimSource.Text);
            root.AppendChild(ele);
            xmlDoc.Save("TaskSetting.xml");
            this.Close();
        }

        private void btn_MiniSize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void btn_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_Open_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.ShowDialog();
            if (open.FileName != "")
            {
                txt_AimSource.Text = open.FileName;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
                this.DragMove();
        }

        private void btn_Cannel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void lab_Title_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            try
            {
                Process p = Process.Start("TaskSetting.xml");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    
    }
}
