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
using System.Data;
using System.Data.OleDb;
using System.Windows.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.IO;
using Microsoft.Win32;
using System.Xml;
using System.Diagnostics;
using System.Threading;

namespace TaskManager
{
    public static class Expend
    {
        public static void LoadImage(this Control sender, string path)
        {
            try
            {
                (sender.Background as ImageBrush).ImageSource = new BitmapImage(new Uri(path, UriKind.RelativeOrAbsolute));
            }
            catch
            {
                MessageBox.Show("图片资源加载失败");
            }
        }

        public static bool IsNum(this string str)
        {
            try
            {
                int.Parse(str);
                return true;
            }
            catch
            {
                
                return false;
            }
        }

        public static DateTime DateTimeParse(string year,string month,string day)
        {
            return new DateTime(int.Parse(year), int.Parse(month), int.Parse(day));
        }

        public static DateTime DateTimeParse(TextBox year, TextBox month, TextBox day)
        {
            return new DateTime(int.Parse(year.Text), int.Parse(month.Text), int.Parse(day.Text));
        }
    }
}
