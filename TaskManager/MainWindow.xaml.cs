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
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        const int ItemHeight = 75;
        const int TitleHeight = 30;
        const int RouteOffest = TitleHeight;
        string xaml;
        Random rand = new Random((int)DateTime.Today.ToBinary() % int.MaxValue);
        List<TaskInformation> list_Tasks = new List<TaskInformation>();

        public static Setting setting = new Setting();

        string ProgramPath = System.AppDomain.CurrentDomain.BaseDirectory;

        public static Color color_Background = new Color();

        //通知栏图标
        WindowState wsl;
        System.Windows.Forms.NotifyIcon notifyIcon = null;

        static MainWindow()
        {
            if (!File.Exists("Data.xml"))
            {
                using (StreamWriter writer = new StreamWriter("Data.xml"))
                {
                    writer.WriteLine("<Datas>");
                    writer.WriteLine("</Datas>");
                }
            }

            if (!File.Exists("TaskSetting.xml"))
            {
                using (StreamWriter writer = new StreamWriter("TaskSetting.xml"))
                {
                    writer.WriteLine("<Settings>");
                    writer.WriteLine("</Settings>");
                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            //UI
            color_Background = Color.FromArgb(200, (byte)rand.Next(255), (byte)rand.Next(255), (byte)rand.Next(255)); //背景色

            this.notifyIcon = new System.Windows.Forms.NotifyIcon();
            ToolBarIcon_init();
            Process[] arrayProcess = Process.GetProcessesByName("TaskManager");
            if (arrayProcess.Count() > 2)
            {
                MessageBox.Show("程序已经再运行！");
                //this.Close();
            }

        }

        private void ToolBarIcon_init()
        {
            this.ShowInTaskbar = false;
            this.notifyIcon.BalloonTipText = "TaskManager 开始你今日的任务"; //设置程序启动时显示的文本
            this.notifyIcon.Text = "TaskManager";//最小化到托盘时，鼠标点击时显示的文本
            this.notifyIcon.Icon = new System.Drawing.Icon("ICON.ico");//程序图标
            this.notifyIcon.Visible = true;

            //右键菜单--打开
            System.Windows.Forms.MenuItem open = new System.Windows.Forms.MenuItem("Show");
            open.Click += new EventHandler((s,e)=>this.WindowState = WindowState.Normal);
            
            //右键菜单--最小化
            System.Windows.Forms.MenuItem miniSize = new System.Windows.Forms.MenuItem("MiniSize");
            miniSize.Click += new EventHandler((s, e) => this.Hide());

            //右键菜单--退出
            System.Windows.Forms.MenuItem exit = new System.Windows.Forms.MenuItem("Exit");
            exit.Click += new EventHandler((s, e) => this.Close());

            //关联托盘控件
            System.Windows.Forms.MenuItem[] childen = new System.Windows.Forms.MenuItem[] { open, miniSize ,exit};
            notifyIcon.ContextMenu = new System.Windows.Forms.ContextMenu(childen);

            notifyIcon.MouseDoubleClick += (s,e)=>
            {
                this.Show();
                this.WindowState = WindowState.Minimized;
                this.WindowState = WindowState.Normal;
            };
            this.notifyIcon.ShowBalloonTip(1000);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            //UI

            this.Icon = new BitmapImage(new Uri(@"ICON.ico", UriKind.Relative));
            Expend.LoadImage(btn_Mode_Setting,@"Texture\pic_Setting.png");

            this.Left = SystemParameters.WorkArea.Width - this.Width;
            this.Top = 0;
            xaml = System.Windows.Markup.XamlWriter.Save(grid_Task_Exmaple);

            this.ResizeMode = ResizeMode.NoResize;

            lab_Title.Content = "今天是:" + DateTime.Today.Date.ToString().Split(' ')[0];
            grid_Tasks.Children.Remove(grid_Task_Exmaple);

            try
            {
                LoadTaskFromDatabase();
            }
            catch
            {

            }
        }

        private void LoadTaskFromDatabase()
        {
            list_Tasks.Clear();
            grid_Tasks.Children.Clear();

            DataTable data = null;

            try
            {
                data = ExcelToDataTable(setting.SourcePath, setting.ChoosenSheet);
            }
            catch
            {
                return;
            }
            //XML
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("Data.xml");
            string strToday = DateTime.Today.ToLongDateString().ToString();
            XmlNode root = xmlDoc.SelectSingleNode("Datas");

            //添加未完成任务
            foreach (XmlElement ele in root.SelectNodes("//TASK[@IsFinish=\"False\"]"))
            {
                TaskInformation task = new TaskInformation()
                {
                    Name = ele.GetAttribute("Name"),
                    Time = ele.GetAttribute("Time"),
                    IsFinish = false
                };
                if (setting.isLoadHistory)
                {
                    list_Tasks.Add(task);
                }
                else
                {
                    if (task.Time == strToday)
                    {
                        list_Tasks.Add(task);
                    }
                }
            }

            //添加今日任务
            foreach (DataRow row in data.Rows)
            {
                if (Regex.IsMatch(row[0].ToString(), string.Format(@"{0}/{1}/{2} 0:00:00", DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day)) ||
                    Regex.IsMatch(row[0].ToString(), string.Format("({0})?年?{1}月{2}日", DateTime.Today.Year, DateTime.Today.Month, DateTime.Today.Day)))
                {
                    for (int index = 1; index < row.ItemArray.Count(); index++)
                    {
                        if (row[index].ToString() != "")
                        {
                            //获得今日任务
                            TaskInformation task = new TaskInformation()
                            {
                                Name = row[index].ToString(),
                                Time = DateTime.Today.ToLongDateString(),
                                IsFinish = false
                            };
                            
                            XmlElement ele = (XmlElement)root.SelectSingleNode(string.Format("//TASK[@Name='{0}'][@Time ='{1}']", task.Name, task.Time));
                            if (ele == null)//如果不在XML表
                            {
                                ele = xmlDoc.CreateElement("TASK");
                                ele.SetAttribute("Name", task.Name);
                                ele.SetAttribute("Time", task.Time);
                                ele.SetAttribute("IsFinish", task.IsFinish.ToString());
                                root.AppendChild(ele);
                                task.IsFinish = false;
                                //计入结构
                                list_Tasks.Add(task);
                            }
                            else//如果在表中获取是否完成
                            {
                                task.IsFinish = ele.GetAttribute("IsFinish") == "True" ? true : false;
                                if (task.IsFinish)
                                {
                                    //计入结构
                                    list_Tasks.Add(task);
                                }
                            }
                        }
                    }
                }
            }
            xmlDoc.Save("Data.xml");

            sortTaskList(list_Tasks);
            
            //UI
            foreach (TaskInformation t in list_Tasks)
            {
                AddTask_UI(t);
            }
            AddLastTaskRow();
        }

        private void sortTaskList(List<TaskInformation> list_Tasks)
        {
            list_Tasks.Sort((left,right) => 
            {
                if(left.Name.StartsWith("星期") && left.Name.Length == 3)
                {
                    return 1;
                }
                else if(right.Name.StartsWith("星期") && left.Name.Length == 3)
                {
                    return -1;
                }

                int left_val = -1, right_val = -1;
                if(left.Name.Contains("-") && left.Name.Split('-').Last().IsNum())
                {
                    left_val = int.Parse(left.Name.Split('-').Last());
                }
                if (right.Name.Contains("-") && right.Name.Split('-').Last().IsNum())
                {
                    right_val = int.Parse(right.Name.Split('-').Last());
                }
                return left_val - right_val;
            }) ;
        }

        public static DataTable ExcelToDataTable(string strExcelFileName, string strSheetName = "Sheet1")
        {
            string strConn = "";
            //源的定义  
            if (strExcelFileName.Split('.').Last() == "xls")
                strConn = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + strExcelFileName + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                strConn = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + strExcelFileName + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            //Sql语句  
            //string strExcel = string.Format("select * from [{0}$]", strSheetName); 这是一种方法  
            string strExcel = "select * from   ["+strSheetName+"$]";

            //定义存放的数据表  
            DataSet ds = new DataSet();

            //连接数据源  
            OleDbConnection conn = new OleDbConnection(strConn);

            try
            {
                conn.Open();

                //适配到数据源  
                OleDbDataAdapter adapter = new OleDbDataAdapter(strExcel, strConn);
                adapter.Fill(ds, strSheetName);

                conn.Close();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.Message);
                MessageBox.Show("数据源配置错误请设置后重启重试");
                setting.Show();
                throw;
            }
            return ds.Tables[strSheetName];
        }

        private void AddTask_UI(string taskName)
        {

            //UI界面
            Grid grid = System.Windows.Markup.XamlReader.Parse(xaml) as Grid;
            grid.Margin = new Thickness(0, grid.Height * (grid_Tasks.Children.Count), 0, 0);

            //grid.Background = new SolidColorBrush(Color.FromArgb(100, (byte)rand.Next(255), (byte)rand.Next(255), (byte)rand.Next(255)));
            grid.Background = new SolidColorBrush(color_Background * (1-grid_Tasks.Children.Count*0.75f/list_Tasks.Count));
            foreach (UIElement UIEle in grid.Children)
            {
                if (UIEle is Label && (UIEle as Label).Name == "lab_TaskName")
                {
                    Label tx = (UIEle as Label);
                    tx.Content =taskName;
                    tx.Foreground = new SolidColorBrush(Colors.White);
                }
                if (UIEle is Label && (UIEle as Label).Name == "lab_Date")
                {
                    Label tx = (UIEle as Label);
                    tx.Visibility = Visibility.Hidden;
                }
                if (UIEle is Button && (UIEle as Button).Name == "btn_Setting")
                {
                    Button btn = UIEle as Button;
                    btn.Click += btn_Setting_Click;
                    btn.Visibility = Visibility.Hidden;
                }
            }

            grid_Tasks.Height = ItemHeight * (grid_Tasks.Children.Count+1);
            grid_Main.Height = grid_Title.ActualHeight + grid_Tasks.Height;
            this.Height = grid_Main.Height;
            grid_Tasks.Children.Add(grid);
        }

        private void AddLastTaskRow()
        {

            //UI界面
            Grid grid = System.Windows.Markup.XamlReader.Parse(xaml) as Grid;
            grid.Margin = new Thickness(0, grid.Height * (grid_Tasks.Children.Count), 0, 0);
            grid.Background = new SolidColorBrush(color_Background * (1 - grid_Tasks.Children.Count * 0.75f / list_Tasks.Count));
            foreach (UIElement UIEle in grid.Children)
            {
                if (UIEle is Label && (UIEle as Label).Name == "lab_TaskName")
                {
                    Label tx = (UIEle as Label);
                    if (list_Tasks.Count == 0)
                        tx.Content = "今日暂无计划";
                    else
                        tx.Content = string.Format("共计{0}个项目",list_Tasks.Count);
                    tx.Foreground = new SolidColorBrush(Colors.White);
                }
                if (UIEle is Label && (UIEle as Label).Name == "lab_Date")
                {
                    Label tx = (UIEle as Label);
                    tx.Visibility = Visibility.Hidden;
                }
                if (UIEle is Button && (UIEle as Button).Name == "btn_Setting")
                {
                    Button btn = UIEle as Button;
                    btn.Content = "+";
                    btn.Click += btn_AddTask;
                    btn.Visibility = Visibility.Visible;
                }
            }

            grid_Tasks.Height = ItemHeight * (grid_Tasks.Children.Count + 1);
            grid_Main.Height = grid_Title.ActualHeight + grid_Tasks.Height;
            this.Height = grid_Main.Height;
            grid_Tasks.Children.Add(grid);
        }

        private void btn_AddTask(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("此方法复杂而且低效 建议直接在EXCEL中操作('◡')");
            TaskAdd dig = new TaskAdd();
            dig.ShowDialog();
        }

        private void AddTask_UI(TaskInformation task)
        {
            //UI界面
            Grid grid = System.Windows.Markup.XamlReader.Parse(xaml) as Grid;
            grid.Margin = new Thickness(0, grid.Height * (grid_Tasks.Children.Count), 0, 0);
            grid.Background = new SolidColorBrush(color_Background * (1-grid_Tasks.Children.Count*0.75f/ list_Tasks.Count));
            foreach (UIElement UIele in grid.Children)
            {
                if (UIele is Label && (UIele as Label).Name == "lab_TaskName")
                {
                    Label tx = (UIele as Label);
                    tx.Content = task.Name;
                    
                    if (task.Name.Length > 15)
                    {
                        tx.FontSize -= 3;
                    }
                    tx.Foreground = new SolidColorBrush(Colors.White);
                    tx.Tag = task;

                    tx.MouseDoubleClick += this.lab_TaskName_MouseDoubleClick;
                }
                if (UIele is Label && (UIele as Label).Name == "lab_Date")
                {
                    Label tx = (UIele as Label);
                    if (task.Time ==DateTime.Today.ToLongDateString())
                    {
                        tx.Visibility = Visibility.Hidden;
                    }
                    tx.Content = task.Time;
                    tx.Foreground = new SolidColorBrush(Colors.White);
                    tx.Tag = task;
                }
                if (UIele is Button && (UIele as Button).Name == "btn_Setting")
                {
                    Button btn = UIele as Button;
                    btn.Tag = task;
                    if (task.IsFinish)
                    {
                        btn.Content = "√";
                    }
                    btn.Click += btn_Setting_Click;
                }
            }

            grid_Tasks.Height = ItemHeight * (grid_Tasks.Children.Count+1);
            grid_Main.Height = grid_Title.ActualHeight + grid_Tasks.Height;
            this.Height = grid_Main.Height;
            grid_Tasks.Children.Add(grid);
        }

        private void btn_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void btn_MiniSize_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
        }

        private void btn_Setting_Click(object sender, RoutedEventArgs e)
        {
            Button btn = sender as Button;
            if (btn.Content.ToString().Equals("?"))
            {
                btn.Content = "√";
                (btn.Tag as TaskInformation).IsFinish = true;
            }
        }

        delegate void GridAnimation(double value);
        private void RunAnimation(GridAnimation animation, double From, double To, int span = 1, int Step = 5000, object actor = null, GridAnimation EndStep = null)
        {
            int itor = 0;
            double nownum = From;
            DispatcherTimer timer = new DispatcherTimer();
            timer.Tick +=
            (o, e) =>
            {
                animation(nownum);
                itor++;
                nownum = itor * (To - From) / Step + From;
                if (itor == Step)
                {
                    timer.IsEnabled = false;
                    EndStep?.Invoke(0);
                };
            };
            timer.IsEnabled = true;
        }

        private void btn_Mode_Setting_Click(object sender, RoutedEventArgs e)
        {
            if (setting.Visibility== Visibility.Visible)
            {
                setting.Visibility = Visibility.Hidden;
            }
            else
            {
                setting.Visibility = Visibility.Visible;
            }

            if (setting.isNeedToUpdate)
            {
                LoadTaskFromDatabase();
                setting.isNeedToUpdate = false;
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void Window_Closed(object sender, EventArgs e)
        {
            setting.Close();
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.Load("Data.xml");
            XmlNode root = xmlDoc.SelectSingleNode("Datas");
            foreach (TaskInformation t in list_Tasks)
            {
                XmlElement ele = (XmlElement)root.SelectSingleNode(string.Format("//TASK[@Name='{0}'][@Time ='{1}']", t.Name, t.Time));
                ele.SetAttribute("IsFinish",t.IsFinish.ToString());
            }
            xmlDoc.Save("Data.xml");
        }

        private void lab_Title_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                try
                {
                    Process p = Process.Start(setting.SourcePath);
                    p.Exited += (o, ev) => { LoadTaskFromDatabase(); };
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void lab_TaskName_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            TaskInformation task = (TaskInformation)(sender as Label).Tag;
            if (e.RightButton == MouseButtonState.Pressed)
            {
                TaskSetting dig = new TaskSetting(task);
                dig.ShowDialog();
            }
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load("TaskSetting.xml");
                XmlNode root = xmlDoc.FirstChild;
                XmlElement xe = null;
                foreach(XmlElement ele in root.ChildNodes)
                {
                    if (task.Name.Contains(ele.GetAttribute("Name"))
                        && (xe == null || 
                         xe.GetAttribute("Name").Length < ele.GetAttribute("Name").Length))
                    {
                        xe = ele;
                    }
                }
                if (xe != null)
                {
                    string strAim = xe.GetAttribute("AimSource");
                    if (Regex.IsMatch(strAim, @"\w+://.+")|| Regex.IsMatch(strAim, @"\w:\\.+"))
                    {
                        try
                        {
                            Process p = Process.Start(strAim);
                        }
                        catch(Exception ex)
                        {
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show(strAim);
                    }
                }
            }
        }
 
        private void TASK_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            Thickness t = grid_Tasks.Margin;
            if (e.Delta>0)
            {
                if (t.Top <= 0)
                    t.Top = t.Top + RouteOffest;
            }
            else 
            {
                if (t.Top >=-(ItemHeight * list_Tasks.Count)+this.ActualHeight-150)
                    t.Top = t.Top - RouteOffest;
            }
            grid_Tasks.Margin = t;
        }
    }
}

    