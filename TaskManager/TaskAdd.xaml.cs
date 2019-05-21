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
using System.Data.OleDb;
using System.Data;
using System.IO;

namespace TaskManager
{
    /// <summary>
    /// TaskAdd.xaml 的交互逻辑
    /// </summary>
    public partial class TaskAdd : Window
    {
        public TaskAdd()
        {
            InitializeComponent();
        }

        void btn_Close_Click(Object o, EventArgs e)
        {
            this.Close();
        }

        private void txt_Number_TextChanged(object sender, TextChangedEventArgs e)
        {
            TextBox text = sender as TextBox;
            if (text.Text.IsNum())
            {
                text.Foreground = new SolidColorBrush(Colors.White);
            }
            else
            {
                text.Foreground =new SolidColorBrush( Colors.Red);
            }
        }

        private void txt_Confirm_Click(object sender, RoutedEventArgs e)
        {
            string strExcelFileName = @"D:\test.xlsx";
            string strConn;
            if (strExcelFileName.Split('.').Last() == "xls")
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strExcelFileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=0\"";
            else
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + strExcelFileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=0\"";
            OleDbConnection myConn = new OleDbConnection(strConn);
            string strCom = "select * from [Sheet1$]";
            myConn.Open();

            OleDbDataAdapter myCommand = new OleDbDataAdapter(strCom, myConn);
            OleDbCommandBuilder builder = new OleDbCommandBuilder(myCommand);
            //QuotePrefix和QuoteSuffix主要是对builder生成InsertComment命令时使用。 
            builder.QuotePrefix = "[";     //获取insert语句中保留字符（起始位置） 
            builder.QuoteSuffix = "]"; //获取insert语句中保留字符（结束位置） 
            DataSet ds = new DataSet();
            myCommand.Fill(ds, "[Sheet1$]");
            DataTable dt = ds.Tables["[Sheet1$]"];

            dt.PrimaryKey = new DataColumn[1] {dt.Columns[0]};
            if (!dt.Columns.Contains(txt_Name.Text))
            {
                dt.Columns.Add(txt_Name.Text);
            }
            for (DateTime Date = Expend.DateTimeParse(txt_From_Year, txt_From_Month, txt_From_Day);
                Date < Expend.DateTimeParse(txt_From_Year, txt_From_Month, txt_From_Day) + TimeSpan.FromDays(int.Parse(txt_Num.Text) * int.Parse(txt_Span.Text)); 
                Date += TimeSpan.FromDays(int.Parse(txt_Span.Text)))
            {
                DataRow aimRow = null;
                foreach (DataRow r in dt.Rows)
                {
                    if (r[0]!= null || (DateTime)r[0] == Date)
                    {
                        aimRow = r;
                        break;
                    }
                }
                if (aimRow ==null)
                {
                    aimRow = dt.NewRow();
                    aimRow[0] = Date;
                    dt.Rows.Add(aimRow);
                }
                aimRow[txt_Name.Text] = txt_Name.Text;
            }

            myCommand.Update(ds, "[Sheet1$]");
            myConn.Close();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            DateTime date = DateTime.Today;
            txt_From_Year.Text = date.Year.ToString();
            txt_From_Month.Text = date.Month.ToString();
            txt_From_Day.Text = date.Day.ToString();
        }

        private void btn_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_Cannel_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
