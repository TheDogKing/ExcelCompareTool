using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using System.Reflection;
using Binding = System.Windows.Data.Binding;
using System.Globalization;
using System.ComponentModel;
using System.Threading;
using Application = System.Windows.Application;
using System.Text;
using System.Xml;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;
using Label=System.Windows.Controls.Label;
using MessageBox = System.Windows.MessageBox;
using MessageBoxButton= System.Windows.MessageBoxButton;
using System.Windows.Controls;
using Microsoft.Win32;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using System.Windows.Input;

namespace ExcelCompare
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public  partial class MainWindow : Window
    {
        TextBoxVisibility hatbv;
        //TextBoxVisibility datbv;
        TextBoxVisibility kctbv;
        TextBoxVisibility ictbv;
        ComboBoxiVisibility hacbv;
        //ComboBoxiVisibility dacbv;
        ComboBoxiVisibility kccbv;
        ComboBoxiVisibility iccbv;
        public string exeaddress = AppDomain.CurrentDomain.BaseDirectory;
        public string savepath = string.Empty;
        public string openfileaddr = AppDomain.CurrentDomain.BaseDirectory;
        public ObservableCollection<string> ignorecolumncollection = new ObservableCollection<string>();
        public ObservableCollection<string> dataareacollection = new ObservableCollection<string>();
        //public ObservableCollection<string> targetsheetcolumncollection = new ObservableCollection<string>();
        public ObservableCollection<string> old_excelsheetcollection = new ObservableCollection<string>();
        public ObservableCollection<string> new_excelsheetcollection = new ObservableCollection<string>();
        public ObservableCollection<string> keycolumncollection = new ObservableCollection<string>();
        public ObservableCollection<string> headareacollection = new ObservableCollection<string>();
        public string targetsheet;
        public string keycolumn;
        public string ignorecolumn;
        public string headarea;
        public string dataarea;
        WatchPath watchpath = new WatchPath();
        Excel.Application excelapp;
        Excel.Application excelapphead;
        Excel.Workbook workbook;
        public DateTime starttime = DateTime.Now;
        public string address = "";

        private BackgroundWorker bw = new BackgroundWorker();
        //public Progress progressbar;
        public int headareaheight = 0;
        public Regex regex = new Regex("[a-zA-Z]+");
        public Regex intregex = new Regex("[0-9]+");
        public MainWindow()
        {
            InitializeComponent();
            SaveFolder.Text = "请选择结果保存地址";
            watchpath.OnHavePath += new WatchPath.WhenHavePath(GetWorkSheets);
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            //this.Topmost = true;
            GetConfig();
            Bind();
            ButtonEnable();
            BindBox();
            BindComboBox();

            bw.WorkerReportsProgress = true;
            bw.WorkerSupportsCancellation = true;//支持取消
            bw.DoWork += new DoWorkEventHandler(BGWorker_DoWork);//开始工作
            bw.ProgressChanged += new ProgressChangedEventHandler(BgWorker_ProgessChanged);//进度改变事件
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BgWorker_WorkerCompleted);//进度完成事件
        }
        private void BGWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            int progresscount = 0;
            //在这里执行耗时的运算。
            //(string)e.Argument == "参数";
            Parameter ok =(Parameter)e.Argument;
            savepath = ok.SavePath;
            List<string> sheetlist = ok.Sheetlist;
            int sheetnums = sheetlist.Count;
            Boolean IsFinish = true;
            double tempdouble = (100 / sheetnums);
            int peervalue=Convert.ToInt32((Math.Round(tempdouble)) /3);
            for (int selectindex = 0; selectindex < sheetnums && IsFinish; selectindex++)
            {
                bw.ReportProgress(progresscount);
                IsFinish = false;
                //OldExcelSheetCombo.SelectedIndex = selectindex;
                //NewExcelSheetCombo.SelectedIndex = selectindex;
                if (sheetlist[selectindex].Equals("首页"))
                {
                    progresscount += 3*peervalue;
                    bw.ReportProgress(progresscount);
                    IsFinish = true;
                    continue;
                }
                try
                {
                    targetsheet = sheetlist[selectindex];
                    //获取excel数据
                    Func<string, string, DataTable> func = GetDataTable;
                    IAsyncResult newref = func.BeginInvoke(ok.NewAdd, targetsheet, null, null);
                    IAsyncResult oldref = func.BeginInvoke(ok.OldAdd, targetsheet, null, null);
                    //oldadd.DataContext = "";
                    //newadd.DataContext = "";
                    DataTable newdata = func.EndInvoke(newref);
                    DataTable olddata = func.EndInvoke(oldref);
                    progresscount += peervalue;
                    bw.ReportProgress(progresscount,"获取Excel数据");
                    //建立新的excel
                    Func<List<Excel.Worksheet>> createexcel = CreateExcel;
                    IAsyncResult excelref = createexcel.BeginInvoke(null, null);

                    //比对excel信息
                    Func<DataTable, string, Dictionary<string, List<string>>> getno_info = Datatohashtable;
                    IAsyncResult newnoref = getno_info.BeginInvoke(newdata, "新", null, null);
                    IAsyncResult oldnoref = getno_info.BeginInvoke(olddata, "旧", null, null);
                    Dictionary<string, List<string>> newno_info = getno_info.EndInvoke(newnoref);
                    Dictionary<string, List<string>> oldno_info = getno_info.EndInvoke(oldnoref);
                    List<Dictionary<string, List<string>>> result = Compare(oldno_info, newno_info);

                    int count = 0;
                    AsyncCallback IsEnd = p =>
                    {
                        count++;
                        if (count >= 3)
                        {
                            string savefile = string.Empty;
                            savefile = savepath + "\\" + targetsheet + "_result.xlsx";
                            if (System.IO.File.Exists(savefile))
                            {
                                System.IO.File.Delete(savefile);
                            }
                            workbook.SaveAs(savefile, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Excel.XlSaveAsAccessMode.xlNoChange, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                            workbook.Close(true);
                            //Ok.KillExcelApp(excelapp);
                            //Ok.KillExcelApp(excelapphead);
                            progresscount += peervalue;
                            bw.ReportProgress(progresscount);
                            Console.WriteLine("AllEnd" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
                            IsFinish = true;
                        }
                    };

                    List<Excel.Worksheet> worksheets = new List<Excel.Worksheet>();
                    worksheets=createexcel.EndInvoke(excelref);
                    progresscount += peervalue;
                    bw.ReportProgress(progresscount);
                    if (result[0].Count == 0 && result[1].Count == 0 && result[2].Count == 0)
                    {
                        workbook.Close(false);
                        IsFinish = true;
                        progresscount += peervalue;
                        bw.ReportProgress(progresscount);
                        continue;
                    }
                    Action<Excel.Worksheet, Dictionary<string, List<string>>, bool> PrintThread = Print;
                    IAsyncResult sheet1=PrintThread.BeginInvoke(worksheets[0], result[0], false, IsEnd, null);
                    IAsyncResult sheet2 = PrintThread.BeginInvoke(worksheets[1], result[1], false, IsEnd, null);
                    IAsyncResult sheet3 = PrintThread.BeginInvoke(worksheets[2], result[2], true, IsEnd, null);
                    PrintThread.EndInvoke(sheet1);
                    PrintThread.EndInvoke(sheet2);
                    PrintThread.EndInvoke(sheet3);
                    Thread.Sleep(3000);
                    Clipboard.Clear();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "warning", MessageBoxButton.OK);
                    if (!(excelapp is null))
                    {
                        ExcelMgr.KillExcelApp(excelapp);
                    }
                    if (!(excelapphead is null))
                    {
                        ExcelMgr.KillExcelApp(excelapphead);
                    }
                    return;
                }
                
            }
            Thread.Sleep(2000);
            ExcelMgr.KillExcelApp(excelapphead);
            ExcelMgr.KillExcelApp(excelapp);
            bw.ReportProgress(100);
            MessageBox.Show("结果文件生成在" + savepath, "成功", MessageBoxButton.OK);
        }
        public void BgWorker_ProgessChanged(object sender, ProgressChangedEventArgs e)
        {
            progressLabel.Content = "1";
            for (int i= (int)progressbar.Value; i<=e.ProgressPercentage;i++)
            {
                Thread.Sleep(10);
                progressbar.Value =i;
            }
            //(string)e.UserState=="Working"
            //progressbar.detail.Content = e.ProgressPercentage.ToString()+"%";
            //progressbar.Value = e.ProgressPercentage;//取得进度更新控件，不用Invoke了
        }
        public void BgWorker_WorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            CompareButton.IsEnabled = true;
            //progressbar.Close();
            //var collections = Application.Current.Windows;
            ////e.Error == null 是否发生错误
            ////e.Result
            ////x.Content = null;
            //collections = Application.Current.Windows;
            //foreach (Window item in collections)
            //{
            //    item.Close();
            //    break;
            //}
            //collections = Application.Current.Windows;
            //e.Cancelled 完成是由于取消还是正常完成
        }

        public void BindComboBox()
        {
            //hacbv = new ComboBoxiVisibility();
            BindingOperations.SetBinding(HeadAreaCombo, ComboBox.VisibilityProperty, new Binding("VisibilityP") { Source=hacbv= new ComboBoxiVisibility()});
            //BindingOperations.SetBinding(DataAreaCombo, ComboBox.VisibilityProperty, new Binding("VisibilityP") { Source=dacbv= new ComboBoxiVisibility()});
            BindingOperations.SetBinding(IgnoreColumnCombo, ComboBox.VisibilityProperty, new Binding("VisibilityP") { Source = iccbv = new ComboBoxiVisibility()});
            BindingOperations.SetBinding(KeyColumnCombo, ComboBox.VisibilityProperty, new Binding("VisibilityP") { Source = kccbv = new ComboBoxiVisibility()});
        }
        public void BindBox()
        {
            hatbv = new TextBoxVisibility();
            Binding habinding = new Binding();
            habinding.Source = hatbv;
            habinding.Path = new PropertyPath("VisibilityP");
            BindingOperations.SetBinding(HeadAreaBox, TextBox.VisibilityProperty, habinding);

            //datbv = new TextBoxVisibility();
            //Binding dabinding = new Binding();
            //dabinding.Source = datbv;
            //dabinding.Path = new PropertyPath("VisibilityP");
            //BindingOperations.SetBinding(DataAreaBox, TextBox.VisibilityProperty, dabinding);

            kctbv = new TextBoxVisibility();
            Binding kcbinding = new Binding();
            kcbinding.Source = kctbv;
            kcbinding.Path = new PropertyPath("VisibilityP");
            BindingOperations.SetBinding(KeyColumnBox, TextBox.VisibilityProperty, kcbinding);

            ictbv = new TextBoxVisibility();
            Binding icbinding = new Binding();
            icbinding.Source = ictbv;
            icbinding.Path = new PropertyPath("VisibilityP");
            BindingOperations.SetBinding(IgnoreColumnBox, TextBox.VisibilityProperty, icbinding);
        }
        public void Bind()
        {
            try
            {
                //targetsheetcolumncollection.Add("选取Excel后，获取选项；");
                IgnoreColumnCombo.ItemsSource = ignorecolumncollection;
                KeyColumnCombo.ItemsSource = keycolumncollection;
                //TargetSheetCombo.ItemsSource = targetsheetcolumncollection;
                HeadAreaCombo.ItemsSource = headareacollection;
                //DataAreaCombo.ItemsSource = dataareacollection;
                OldExcelSheetCombo.ItemsSource = old_excelsheetcollection;
                NewExcelSheetCombo.ItemsSource = new_excelsheetcollection;

                KeyColumnCombo.SelectedIndex = 0;
                HeadAreaCombo.SelectedIndex = 0;
                IgnoreColumnCombo.SelectedIndex = 0;
                //TargetSheetCombo.SelectedIndex = 0;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "warning", MessageBoxButton.OK);
                throw ex;
            }

        }
        public void ButtonEnable()
        {
            Binding oldbind = new Binding
            {
                Path = new PropertyPath("Text"),
                Source = OldExcelSheetCombo
            };
            Binding newbind = new Binding
            {
                Path = new PropertyPath("Text"),
                Source = NewExcelSheetCombo
            };
            MultiBinding bindings = new MultiBinding
            {
                Mode = BindingMode.OneWay
            };
            bindings.Bindings.Add(oldbind);
            bindings.Bindings.Add(newbind);
            bindings.Converter = new Multibindingconverter();
            CompareButton.SetBinding(IsEnabledProperty, bindings);
        }
        public class Multibindingconverter : IMultiValueConverter
        {

            public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
            {
                if (!values.Cast<string>().Any(text => string.IsNullOrEmpty(text)))
                {
                    return true;
                }
                return false;
                //throw new NotImplementedException();
            }

            public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
            {
                throw new NotImplementedException();
            }

        }
        private void Browse_Excel(object sender, RoutedEventArgs e)
        {
            try
            {
                string buttontype = ((System.Windows.Controls.Button)sender).Name;
                string fileaddress = "";
                string filefulladdress = "";
                OpenFileDialog BrowseFile = new OpenFileDialog
                {
                    InitialDirectory = openfileaddr,
                    Filter = "Excel files (*.xlsx,*.xlsm,*.xls)|*.xlsx;*xlsm;*.xls",
                    FilterIndex = 1,
                    //RestoreDirectory = true
                };

                if (BrowseFile.ShowDialog() == true)
                {
                    fileaddress = BrowseFile.SafeFileName;
                    filefulladdress = BrowseFile.FileName;
                    if (buttontype == "BrowseNew")
                    {
                        newadd.Text = fileaddress;
                        newadd.DataContext = filefulladdress;
                        watchpath.oc = new_excelsheetcollection;
                        watchpath.Path = filefulladdress;

                        NewExcelSheetCombo.SelectedIndex = 0;
                    }
                    else if (buttontype == "BrowseOld")
                    {
                        oldadd.Text = fileaddress;
                        oldadd.DataContext = filefulladdress;
                        watchpath.oc = old_excelsheetcollection;
                        watchpath.Path = filefulladdress;

                        OldExcelSheetCombo.SelectedIndex =0;
                    }
                    openfileaddr = BrowseFile.FileName.Substring(0, filefulladdress.LastIndexOf('\\'));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show( ex.Message, "warning", MessageBoxButton.OK);
                throw ex;
            }
        }
        private void Compare_Click(object sender, RoutedEventArgs e)
        {
            CompareButton.IsEnabled = false;
            if (!Directory.Exists(SaveFolder.Text))
            {
                MessageBox.Show("所选地址不存在，请点击Save Folder并从新选取", "Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            progressbar.Value = 3;
            Parameter ok = new Parameter();
            keycolumn = KeyColumnCombo.Text;
            ignorecolumn = IgnoreColumnCombo.Text;
            headarea = HeadAreaCombo.Text;
            //dataarea = DataAreaCombo.Text;
            address = oldadd.DataContext.ToString();
            ok.NewAdd = newadd.DataContext.ToString();
            ok.OldAdd = oldadd.DataContext.ToString();
            ok.SavePath = SaveFolder.Text;
            Match hastart=intregex.Match( headarea.Split(':')[0]);
            Match haend=intregex.Match( headarea.Split(':')[1]);
            headareaheight = Convert.ToInt32(haend.Value)- Convert.ToInt32 (hastart.Value)+1;
            List<string> lcolComparetemp = new List<string>();
            if (IsAllCompare.IsChecked.Value)
            {
                foreach(string item in old_excelsheetcollection)
                {
                    if (new_excelsheetcollection.Contains(item))
                    {
                        lcolComparetemp.Add(item);
                    }
                }
            }
            else
            {
                if (new_excelsheetcollection.Contains(OldExcelSheetCombo.Text))
                {
                    lcolComparetemp.Add(OldExcelSheetCombo.Text);
                }
            }
            if(lcolComparetemp.Count==0)
            {
                MessageBox.Show("新版Excel中不含 旧版选中sheet名", "注意", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            ok.Sheetlist = lcolComparetemp;
            //progressbar = new Progress();

            excelapp = new Excel.Application
            {
                Visible = false,
                StandardFontSize = 8,
                DisplayAlerts = false
            };

            excelapphead = new Excel.Application
            {
                DisplayAlerts = false
            };

            bw.RunWorkerAsync(ok);
            //progressbar.ShowDialog();
            string savename = SaveFolder.Text + "\\" + oldadd.Text.Split('.')[0]+".txt";
            if (System.IO.File.Exists(savename))
            {
                System.IO.File.Delete(savename);
            }
            FileStream fs = new FileStream(savename, FileMode.CreateNew,FileAccess.ReadWrite);
            StreamWriter sw = new StreamWriter(fs);
            sw.WriteLine("删除sheet");
            foreach (string item in old_excelsheetcollection)
            {
                if (!new_excelsheetcollection.Contains(item))
                {
                    sw.WriteLine(item);
                }
            }
            sw.WriteLine("新增sheet");
            foreach (string item in new_excelsheetcollection)
            {
                if (!old_excelsheetcollection.Contains(item))
                {
                    sw.WriteLine(item);
                }
            }
            sw.Close();
            fs.Close();
            SaveConfig();
            //System.Diagnostics.Process.Start(SaveFolder.Text);
        }
        public Dictionary<string,List<string>> Datatohashtable(DataTable dt,string ok)
        {
            string dataStart = headarea.Split(':')[0];
            string dataend = headarea.Split(':')[1];
            int columnStart =ColumnIndex[regex.Match(dataStart).Value.ToUpper()];
            int columnEnd = ColumnIndex[regex.Match(dataend).Value.ToUpper()];
            int rowmark = Convert.ToInt32(intregex.Match(dataend).Value);
            Console.WriteLine("Datatohashtable-start" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            Dictionary<string, List<string>> NO_Info = new Dictionary<string, List<string>>();
            List<int> keyCol = new List<int>(GetNum(keycolumn));
            int i = rowmark;
            while (i < dt.Rows.Count)
            {
                string key = "";
                if (dt.Rows[i][keyCol[0] - columnStart].ToString() == "")
                {
                    i++;
                    continue;
                }
                //Pipeinfo tempinfo = new Pipeinfo();
                //PropertyInfo[] properties = tempinfo.GetType().GetProperties();
                List<string> lcoltempinfo = new List<string>();
                //foreach (PropertyInfo property in properties)
                //{
                //    //.ToString()
                //    string tempproperty = "";
                //    tempproperty = dt.Rows[i][j].ToString();
                //    property.SetValue(tempinfo, tempproperty, null);
                //    j++;
                //}
                for (int j = columnStart - 1; j<= columnEnd; j++)
                {
                    lcoltempinfo.Add(dt.Rows[i][j].ToString());
                }
                foreach (int item in keyCol)
                {
                    key += lcoltempinfo[item-columnStart];
                }
                if(NO_Info.ContainsKey(key))
                {
                    throw new Exception(string.Format("请检查版本较{0}Excel表格，表内此关键字{1}存在多处,请重新设置Key Column",ok,key));
                }
                else
                {
                    NO_Info.Add(key, lcoltempinfo);
                }
                i++;
            }
            Console.WriteLine("Datatohashtable-end" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            return NO_Info;
        }
        public DataTable GetDataTable(string path,string excelsheet)
        {
            //List<int> keyCol = new List<int>(GetNum(keycolumn));
            //string filtersql = string.Empty;
            //foreach(int item in keyCol)
            //{
            //    filtersql += (filtersql == "") ? string.Format("F{0}", item) : string.Format("and F{0}", item);
            //}
            Console.WriteLine("GetDataTable-start" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            DataTable dt=new DataTable();
            string filetype = System.IO.Path.GetExtension(path);
            string strConn = "";
            switch (filetype)
            {
                case ".xls":
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";" + "Extended Properties=\"Excel 8.0;HDR=No;IMEX=1;\"";
                    break;
                case ".xlsx":
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties=\"Excel 12.0;HDR=No;IMEX=1;\"";
                    //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)  
                    //备注： "HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                    break;
                default:
                    strConn = null;
                    break;
            }
            OleDbConnection conn = new OleDbConnection(string.Format("{0}", strConn));
            try
            {
                conn.Open();
                string strExcel = "";
                OleDbDataAdapter myCommand = null;
                DataSet ds = null;
                strExcel = string.Format("select * from [{0}$]", excelsheet);
                // where {1} is not null filtersql
                myCommand = new OleDbDataAdapter(strExcel, strConn);
                ds = new DataSet();
                myCommand.Fill(ds, "table1");
                dt = ds.Tables["table1"];
            }
            catch (Exception ex)
            {
                if (conn.State != ConnectionState.Open)
                    throw new Exception("未能读取Excel，原因如下：" + "\r\n" + ex.Message);
                else if (conn.State == ConnectionState.Open)
                    throw new Exception(ex.Message);
            }
            finally
            {
                conn.Close();
            }
            Console.WriteLine("GetDataTable-end" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            return dt;
        }
        public List<Dictionary<string, List<string>>> Compare(Dictionary<string, List<string>> oldtable, Dictionary<string, List<string>> newtable)
        {
            Dictionary<string, List<string>> delitems = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> newitems = new Dictionary<string, List<string>>();
            Dictionary<string, List<string>> updateitems = new Dictionary<string, List<string>>();
            List<int> ignorelist = GetNum(ignorecolumn);
            //for(int x=0;x<ignorelist.Count;x++)
            //{
            //    ignorelist[x] -= 65;
            //}
            foreach (KeyValuePair<string,List<string>> item in oldtable)
            {
                if (!newtable.ContainsKey(item.Key))
                {
                    delitems.Add(item.Key, item.Value);
                }
                else
                {
                    //Pipeinfo oldpipeinfo = (Pipeinfo)temp.Value;
                    //Pipeinfo newpipeinfo = (Pipeinfo)newtable[temp.Key];
                    //Pipeinfo updatepipeinfo = new Pipeinfo(newpipeinfo);
                    //PropertyInfo[] oldproperties = oldpipeinfo.GetType().GetProperties();
                    //PropertyInfo[] newproperties = newpipeinfo.GetType().GetProperties();
                    List<string> olditeminfo = item.Value;
                    List<string> newiteminfo = newtable[item.Key];
                    List<string> updateinfo = new List<string>(newiteminfo);

                    Boolean IsUpdated = false;
                    for (int i =0;i<newiteminfo.Count;i++)
                    {
                        string oldvalue = olditeminfo[i];
                        string newvalue = newiteminfo[i];
                        try
                        {
                            double doldvalue = Convert.ToDouble(oldvalue);
                            double dnewvalue = Convert.ToDouble(newvalue);
                            if (doldvalue != dnewvalue && !ignorelist.Contains(i+1))
                            {
                                if (string.IsNullOrEmpty(oldvalue) | string.IsNullOrWhiteSpace(oldvalue))
                                {
                                    oldvalue = "空";
                                }
                                if (string.IsNullOrEmpty(newvalue) | string.IsNullOrWhiteSpace(newvalue))
                                {
                                    newvalue = "空";
                                }
                                IsUpdated = true;
                                updateinfo[i] = string.Format("{0}->{1}", oldvalue, newvalue);
                            }
                        }
                        catch (Exception)
                        {
                            if (oldvalue != newvalue && !ignorelist.Contains(i+1))
                            {
                                if (string.IsNullOrEmpty(oldvalue) | string.IsNullOrWhiteSpace(oldvalue))
                                {
                                    oldvalue = "空";
                                }
                                if (string.IsNullOrEmpty(newvalue) | string.IsNullOrWhiteSpace(newvalue))
                                {
                                    newvalue = "空";
                                }
                                IsUpdated = true;
                                updateinfo[i] = string.Format("{0}->{1}", oldvalue, newvalue);
                            }
                        }
                    }
                    if (IsUpdated)
                    {
                        updateitems.Add(item.Key, updateinfo);
                    }
                }
            }
            foreach (KeyValuePair<string, List<string>> temp in newtable)
            {
                if (!oldtable.ContainsKey(temp.Key))
                {
                    newitems.Add(temp.Key, temp.Value);
                }
            }
            List<Dictionary<string, List<string>>> result = new List<Dictionary< string, List< string >>>
            {
                newitems,
                delitems,
                updateitems
            };
            return result;
        }
        public void Print(Excel.Worksheet worksheet, Dictionary<string, List<string>> items, Boolean Isupdate)
        {
            //worksheet.Activate();
            worksheet.Columns["B"].NumberFormat = "@";
            Console.WriteLine("print-start" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            int rowno = headareaheight+1;
            if(items.Count==0)
            {
                worksheet.Delete();
                return;
            }
            foreach (KeyValuePair<string,List<string>> item in items)
            {
                //Pipeinfo tempinfo = (Pipeinfo)temp.Value;
                //PropertyInfo[] properties = tempinfo.GetType().GetProperties();
                List<string> lcoliteminfo = item.Value;
                int colno = 1;
                foreach (string iteminfo in lcoliteminfo)
                {
                    //string value = Propertyinfo.GetValue(tempinfo).ToString();
                    worksheet.Cells[colno][rowno].value  = iteminfo;
                    if (Isupdate)
                    {
                        if (iteminfo.Contains("->"))
                        {
                            worksheet.Cells[colno][rowno].Interior.ColorIndex = 35;
                        }
                    }
                    colno++;
                }
                rowno++;
            }
            worksheet.Columns.EntireColumn.AutoFit();
            Console.WriteLine("print-end" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
        }
        public List<Excel.Worksheet> CreateExcel()
        {
            Console.WriteLine("CreateExcel-start" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            List<Excel.Worksheet> temp = CreateSheets();
            CreateHead(temp);
            Console.WriteLine("CreateExcel-end" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"));
            return temp;
        }
        public List<Excel.Worksheet> CreateSheets()
        {
            Excel.Worksheet newsheet;
            Excel.Worksheet delsheet;
            Excel.Worksheet updatesheet;
            workbook = excelapp.Workbooks.Add();
            workbook.Sheets.Add();
            workbook.Sheets.Add();
            newsheet = workbook.Sheets[1];
            newsheet.Name = "NewItems";
            delsheet = workbook.Sheets[2];
            delsheet.Name = "DelItems";
            updatesheet = workbook.Sheets[3];
            updatesheet.Name = "UpdateItems";
            List<Excel.Worksheet> ok = new List<Excel.Worksheet>
            {
                newsheet,
                delsheet,
                updatesheet
            };
            return ok;
        }
        public void CreateHead(List<Excel.Worksheet> worksheets)
        {
            //先拷贝一份表头
            Excel.Workbook workbookhead;
            workbookhead = excelapphead.Workbooks.Open(address);
            try
            {
                Excel.Worksheet worksheethead = new Excel.Worksheet();
                worksheethead = workbookhead.Sheets[targetsheet];
                //string selectitem = HeadArea.Items.CurrentItem.ToString();
                headarea = headarea.Split(';')[0];
                worksheethead.Activate();
                Excel.Range ran = worksheethead.Range[headarea];
                ran.Copy();
                Thread.Sleep(10000);
                foreach (Excel.Worksheet worksheet in worksheets)
                {
                    worksheet.Activate();
                    Excel.Range newran = worksheet.Range[GetHeadArea(headarea)];
                    newran.Select();
                    //worksheet.PasteSpecial("文本",false,false);
                    newran.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationAdd, Type.Missing, Type.Missing);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                workbookhead.Close(false);
            }
        }
        public List<int> GetNum(string lstrcolumn)
        {
            List<int> lcolindex = new List<int>();
            string[] lcoltemp;
            lstrcolumn.Trim();
            lstrcolumn.ToUpper();
            lcoltemp = lstrcolumn.Split(';');
            int i = 0;
            foreach (string item in lcoltemp)
            {
                if(item=="")
                {
                    break;
                }
                //byte[] array = new byte[1];
                //array = Encoding.ASCII.GetBytes(s[i]);
                int index=ColumnIndex[item];
                lcolindex.Add(index);
                //array[0]
                i++;
            }
            return lcolindex;
        }
        private Dictionary<string, int> ColumnIndex = new Dictionary<string, int>();
        public string GetHeadArea(string area)
        {
            area = area.ToUpper();
            List<int> num = new List<int>();
            foreach (Match something in intregex.Matches(area))
            {
                num.Add(Convert.ToInt32(something.Value));
            }
            List<string> column = new List<string>();
            foreach (Match something in regex.Matches(area))
            {
                column.Add(something.Value);
            }
            int x = GetNum(column[1])[0] - GetNum(column[0])[0]+1;
            int c = num[1] - num[0] + 1;
            string end = ColumnIndex.FirstOrDefault(q => q.Value == x).Key;
            string result = "A1:" + end + c;
            return result;
        }
        public void GetConfig()
        {
            string xmladdress = string.Format(@"{0}Excel Compare.xml", exeaddress);
            XmlDocument doc = new XmlDocument();
            XmlReaderSettings settings = new XmlReaderSettings
            {
                IgnoreComments = true//忽略文档里面的注释
            };
            try
            {
                XmlReader reader = XmlReader.Create(xmladdress, settings);
                doc.Load(reader);
                //XmlNode TargetSheets = doc.SelectSingleNode("Config/TargetSheets");
                XmlNode HeadAreaNode = doc.SelectSingleNode("Config/HeadAreas");
                XmlNode KeyColumnNode = doc.SelectSingleNode("Config/KeyColumns");
                XmlNode IgnoreColumnNode = doc.SelectSingleNode("Config/IgnoreColumns");
                XmlNode SavePathNode = doc.SelectSingleNode("Config/SavePath");
                XmlNode ColumnIndexsNode = doc.SelectSingleNode("Config/ColumnIndexs");
                GetItems(IgnoreColumnNode, ignorecolumncollection);
                GetItems(HeadAreaNode, headareacollection);
                GetItems(KeyColumnNode, keycolumncollection);
                GetColumnIndex(ColumnIndexsNode, ColumnIndex);
                SaveFolder.Text = SavePathNode.InnerText;
                reader.Close();
            }
            catch
            {

                MessageBox.Show(exeaddress + "下Excel Compare.xml.xml配置文件丢失！,单击确认生成默认Excel Compare.xml.xml", "Warning", MessageBoxButton.OK);
                string xml = string.Format(@"<Config>
                                                  <HeadAreas>
                                                    <HeadArea>A7:Y10</HeadArea>
                                                    <HeadArea>A7:Z10</HeadArea>
                                                  </HeadAreas>
                                                  <KeyColumns>
                                                    <KeyColumn>B</KeyColumn>
                                                    <KeyColumn>A;B</KeyColumn>
                                                  </KeyColumns>
                                                  <IgnoreColumns>
                                                    <IgnoreColumn>A</IgnoreColumn>
                                                    <IgnoreColumn></IgnoreColumn>
                                                  </IgnoreColumns>
                                                  <SavePath>C:\</SavePath>
                                                    <ColumnIndexs>
                                                    <ColumnIndex name='A'>1</ColumnIndex>
                                                    <ColumnIndex name='B'>2</ColumnIndex>
                                                    <ColumnIndex name='C'>3</ColumnIndex>
                                                    <ColumnIndex name='D'>4</ColumnIndex>
                                                    <ColumnIndex name='E'>5</ColumnIndex>
                                                    <ColumnIndex name='F'>6</ColumnIndex>
                                                    <ColumnIndex name='G'>7</ColumnIndex>
                                                    <ColumnIndex name='H'>8</ColumnIndex>
                                                    <ColumnIndex name='I'>9</ColumnIndex>
                                                    <ColumnIndex name='J'>10</ColumnIndex>
                                                    <ColumnIndex name='K'>11</ColumnIndex>
                                                    <ColumnIndex name='L'>12</ColumnIndex>
                                                    <ColumnIndex name='M'>13</ColumnIndex>
                                                    <ColumnIndex name='N'>14</ColumnIndex>
                                                    <ColumnIndex name='O'>15</ColumnIndex>
                                                    <ColumnIndex name='P'>16</ColumnIndex>
                                                    <ColumnIndex name='Q'>17</ColumnIndex>
                                                    <ColumnIndex name='R'>18</ColumnIndex>
                                                    <ColumnIndex name='S'>19</ColumnIndex>
                                                    <ColumnIndex name='T'>20</ColumnIndex>
                                                    <ColumnIndex name='U'>21</ColumnIndex>
                                                    <ColumnIndex name='V'>22</ColumnIndex>
                                                    <ColumnIndex name='W'>23</ColumnIndex>
                                                    <ColumnIndex name='X'>24</ColumnIndex>
                                                    <ColumnIndex name='Y'>25</ColumnIndex>
                                                    <ColumnIndex name='Z'>26</ColumnIndex>
                                                    <ColumnIndex name='AA'>27</ColumnIndex>
                                                    <ColumnIndex name='AB'>28</ColumnIndex>
                                                    <ColumnIndex name='AC'>29</ColumnIndex>
                                                    <ColumnIndex name='AD'>30</ColumnIndex>
                                                    <ColumnIndex name='AE'>31</ColumnIndex>
                                                    <ColumnIndex name='AF'>32</ColumnIndex>
                                                    <ColumnIndex name='AG'>33</ColumnIndex>
                                                    <ColumnIndex name='AH'>34</ColumnIndex>
                                                    <ColumnIndex name='AI'>35</ColumnIndex>
                                                    <ColumnIndex name='AJ'>36</ColumnIndex>
                                                    <ColumnIndex name='AK'>37</ColumnIndex>
                                                    <ColumnIndex name='AL'>38</ColumnIndex>
                                                    <ColumnIndex name='AM'>39</ColumnIndex>
                                                    <ColumnIndex name='AN'>40</ColumnIndex>
                                                    <ColumnIndex name='AO'>41</ColumnIndex>
                                                    <ColumnIndex name='AP'>42</ColumnIndex>
                                                    <ColumnIndex name='AQ'>43</ColumnIndex>
                                                    <ColumnIndex name='AR'>44</ColumnIndex>
                                                    <ColumnIndex name='AS'>45</ColumnIndex>
                                                    <ColumnIndex name='AT'>46</ColumnIndex>
                                                    <ColumnIndex name='AU'>47</ColumnIndex>
                                                    <ColumnIndex name='AV'>48</ColumnIndex>
                                                    <ColumnIndex name='AW'>49</ColumnIndex>
                                                    <ColumnIndex name='AX'>50</ColumnIndex>
                                                    <ColumnIndex name='AY'>51</ColumnIndex>
                                                    <ColumnIndex name='AZ'>52</ColumnIndex>
                                                    <ColumnIndex name='BA'>53</ColumnIndex>
                                                    <ColumnIndex name='BB'>54</ColumnIndex>
                                                    <ColumnIndex name='BC'>55</ColumnIndex>
                                                    <ColumnIndex name='BD'>56</ColumnIndex>
                                                    <ColumnIndex name='BE'>57</ColumnIndex>
                                                    <ColumnIndex name='BF'>58</ColumnIndex>
                                                    <ColumnIndex name='BG'>59</ColumnIndex>
                                                    <ColumnIndex name='BH'>60</ColumnIndex>
                                                    <ColumnIndex name='BI'>61</ColumnIndex>
                                                    <ColumnIndex name='BJ'>62</ColumnIndex>
                                                    <ColumnIndex name='BK'>63</ColumnIndex>
                                                    <ColumnIndex name='BL'>64</ColumnIndex>
                                                    <ColumnIndex name='BM'>65</ColumnIndex>
                                                    <ColumnIndex name='BN'>66</ColumnIndex>
                                                    <ColumnIndex name='BO'>67</ColumnIndex>
                                                    <ColumnIndex name='BP'>68</ColumnIndex>
                                                    <ColumnIndex name='BQ'>69</ColumnIndex>
                                                    <ColumnIndex name='BR'>70</ColumnIndex>
                                                    <ColumnIndex name='BS'>71</ColumnIndex>
                                                    <ColumnIndex name='BT'>72</ColumnIndex>
                                                    <ColumnIndex name='BU'>73</ColumnIndex>
                                                    <ColumnIndex name='BV'>74</ColumnIndex>
                                                    <ColumnIndex name='BW'>75</ColumnIndex>
                                                    <ColumnIndex name='BX'>76</ColumnIndex>
                                                    <ColumnIndex name='BY'>77</ColumnIndex>
                                                    <ColumnIndex name='BZ'>78</ColumnIndex>
                                                    </ColumnIndexs>
                                             </Config>");
                doc.LoadXml(xml);
                doc.Save(xmladdress);
                GetConfig();
            }
            finally
            {
            }

        }
        public void GetItems(XmlNode xn, ObservableCollection<string> oc)
        {
            XmlNodeList xnl = xn.ChildNodes;
            foreach (XmlElement xe in xnl)
            {
                oc.Add(xe.InnerText.ToString());
            }
        }
        public void GetColumnIndex(XmlNode xn, Dictionary<string, int> lcol)
        {
            XmlNodeList xnl = xn.ChildNodes;
            foreach (XmlElement xe in xnl)
            {
                lcol.Add(xe.GetAttribute("name"), Convert.ToInt32(xe.InnerText));
            }
        }
        public void SaveConfig()
        {
            string xmladdress = string.Format(@"{0}Excel Compare.xml", exeaddress);

            if (File.Exists(xmladdress))
            {
                File.Delete(xmladdress);
            }
            XmlDocument doc = new XmlDocument();

            string strheadarea = string.Empty;
            List<string> temp = new List<string>(headareacollection);
            temp.Sort();
            foreach (string item in temp)
            {
                strheadarea += string.Format("<HeadArea>{0}</HeadArea>", item);
            }
            string strkeycolumn = string.Empty;
            temp = new List<string>(keycolumncollection);
            temp.Sort();
            foreach (string item in temp)
            {
                strkeycolumn += string.Format("<KeyColumn>{0}</KeyColumn>", item);
            }
            string strignorecolumn = string.Empty;
            temp = new List<string>(ignorecolumncollection);
            temp.Sort();
            foreach (string item in ignorecolumncollection)
            {
                strignorecolumn += string.Format("<IgnoreColumn>{0}</IgnoreColumn>", item);
            }
            string xml = string.Format(@"<Config>
                                                  <HeadAreas>{0}
                                                  </HeadAreas>
                                                  <KeyColumns>{1}
                                                  </KeyColumns>
                                                  <IgnoreColumns>{2}
                                                  </IgnoreColumns>
                                                    <SavePath>{3}</SavePath>
<ColumnIndexs>
                                                    <ColumnIndex name='A'>1</ColumnIndex>
                                                    <ColumnIndex name='B'>2</ColumnIndex>
                                                    <ColumnIndex name='C'>3</ColumnIndex>
                                                    <ColumnIndex name='D'>4</ColumnIndex>
                                                    <ColumnIndex name='E'>5</ColumnIndex>
                                                    <ColumnIndex name='F'>6</ColumnIndex>
                                                    <ColumnIndex name='G'>7</ColumnIndex>
                                                    <ColumnIndex name='H'>8</ColumnIndex>
                                                    <ColumnIndex name='I'>9</ColumnIndex>
                                                    <ColumnIndex name='J'>10</ColumnIndex>
                                                    <ColumnIndex name='K'>11</ColumnIndex>
                                                    <ColumnIndex name='L'>12</ColumnIndex>
                                                    <ColumnIndex name='M'>13</ColumnIndex>
                                                    <ColumnIndex name='N'>14</ColumnIndex>
                                                    <ColumnIndex name='O'>15</ColumnIndex>
                                                    <ColumnIndex name='P'>16</ColumnIndex>
                                                    <ColumnIndex name='Q'>17</ColumnIndex>
                                                    <ColumnIndex name='R'>18</ColumnIndex>
                                                    <ColumnIndex name='S'>19</ColumnIndex>
                                                    <ColumnIndex name='T'>20</ColumnIndex>
                                                    <ColumnIndex name='U'>21</ColumnIndex>
                                                    <ColumnIndex name='V'>22</ColumnIndex>
                                                    <ColumnIndex name='W'>23</ColumnIndex>
                                                    <ColumnIndex name='X'>24</ColumnIndex>
                                                    <ColumnIndex name='Y'>25</ColumnIndex>
                                                    <ColumnIndex name='Z'>26</ColumnIndex>
                                                    <ColumnIndex name='AA'>27</ColumnIndex>
                                                    <ColumnIndex name='AB'>28</ColumnIndex>
                                                    <ColumnIndex name='AC'>29</ColumnIndex>
                                                    <ColumnIndex name='AD'>30</ColumnIndex>
                                                    <ColumnIndex name='AE'>31</ColumnIndex>
                                                    <ColumnIndex name='AF'>32</ColumnIndex>
                                                    <ColumnIndex name='AG'>33</ColumnIndex>
                                                    <ColumnIndex name='AH'>34</ColumnIndex>
                                                    <ColumnIndex name='AI'>35</ColumnIndex>
                                                    <ColumnIndex name='AJ'>36</ColumnIndex>
                                                    <ColumnIndex name='AK'>37</ColumnIndex>
                                                    <ColumnIndex name='AL'>38</ColumnIndex>
                                                    <ColumnIndex name='AM'>39</ColumnIndex>
                                                    <ColumnIndex name='AN'>40</ColumnIndex>
                                                    <ColumnIndex name='AO'>41</ColumnIndex>
                                                    <ColumnIndex name='AP'>42</ColumnIndex>
                                                    <ColumnIndex name='AQ'>43</ColumnIndex>
                                                    <ColumnIndex name='AR'>44</ColumnIndex>
                                                    <ColumnIndex name='AS'>45</ColumnIndex>
                                                    <ColumnIndex name='AT'>46</ColumnIndex>
                                                    <ColumnIndex name='AU'>47</ColumnIndex>
                                                    <ColumnIndex name='AV'>48</ColumnIndex>
                                                    <ColumnIndex name='AW'>49</ColumnIndex>
                                                    <ColumnIndex name='AX'>50</ColumnIndex>
                                                    <ColumnIndex name='AY'>51</ColumnIndex>
                                                    <ColumnIndex name='AZ'>52</ColumnIndex>
                                                    <ColumnIndex name='BA'>53</ColumnIndex>
                                                    <ColumnIndex name='BB'>54</ColumnIndex>
                                                    <ColumnIndex name='BC'>55</ColumnIndex>
                                                    <ColumnIndex name='BD'>56</ColumnIndex>
                                                    <ColumnIndex name='BE'>57</ColumnIndex>
                                                    <ColumnIndex name='BF'>58</ColumnIndex>
                                                    <ColumnIndex name='BG'>59</ColumnIndex>
                                                    <ColumnIndex name='BH'>60</ColumnIndex>
                                                    <ColumnIndex name='BI'>61</ColumnIndex>
                                                    <ColumnIndex name='BJ'>62</ColumnIndex>
                                                    <ColumnIndex name='BK'>63</ColumnIndex>
                                                    <ColumnIndex name='BL'>64</ColumnIndex>
                                                    <ColumnIndex name='BM'>65</ColumnIndex>
                                                    <ColumnIndex name='BN'>66</ColumnIndex>
                                                    <ColumnIndex name='BO'>67</ColumnIndex>
                                                    <ColumnIndex name='BP'>68</ColumnIndex>
                                                    <ColumnIndex name='BQ'>69</ColumnIndex>
                                                    <ColumnIndex name='BR'>70</ColumnIndex>
                                                    <ColumnIndex name='BS'>71</ColumnIndex>
                                                    <ColumnIndex name='BT'>72</ColumnIndex>
                                                    <ColumnIndex name='BU'>73</ColumnIndex>
                                                    <ColumnIndex name='BV'>74</ColumnIndex>
                                                    <ColumnIndex name='BW'>75</ColumnIndex>
                                                    <ColumnIndex name='BX'>76</ColumnIndex>
                                                    <ColumnIndex name='BY'>77</ColumnIndex>
                                                    <ColumnIndex name='BZ'>78</ColumnIndex>
                                                    </ColumnIndexs>
                                             </Config>", strheadarea, strkeycolumn, strignorecolumn,savepath);
            doc.LoadXml(xml);
            doc.Save(xmladdress);
        }

        public void GetWorkSheets()
        {
            string path = watchpath.Path;
            string filetype = System.IO.Path.GetExtension(path);
            string strConn = "";
            switch (filetype)
            {
                case ".xls":
                    strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + ";" + "Extended Properties=\"Excel 8.0;HDR=yes;IMEX=1;\"";
                    break;
                case ".xlsx":
                    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";" + "Extended Properties=\"Excel 12.0;HDR=yes;IMEX=1;\"";
                    //此连接可以操作.xls与.xlsx文件 (支持Excel2003 和 Excel2007 的连接字符串)  
                    //备注： "HDR=yes;"是说Excel文件的第一行是列名而不是数，"HDR=No;"正好与前面的相反。
                    //"IMEX=1 "如果列中的数据类型不一致，使用"IMEX=1"可必免数据类型冲突。 
                    break;
                default:
                    strConn = null;
                    break;
            }
            using (OleDbConnection conn = new OleDbConnection(strConn))
            {
                watchpath.oc.Clear();
                try
                {
                    conn.Open();
                    //获取所有的worksheet
                    List<string> tableName = new List<string>();
                    DataTable worksheetnamedt = conn.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables_Info, null);
                    foreach (System.Data.DataRow row in worksheetnamedt.Rows)
                    {
                        string strSheetTableName = row["TABLE_NAME"].ToString();
                        //过滤无效SheetName   
                        if (strSheetTableName.Contains("$") && strSheetTableName.Replace("'", "").EndsWith("$"))
                        {
                            strSheetTableName = strSheetTableName.Replace("'", "");   //可能会有 '1X$' 出现
                            strSheetTableName = strSheetTableName.Substring(0, strSheetTableName.Length - 1);
                            watchpath.oc.Add(strSheetTableName);
                        }
                    }
                    if(watchpath.oc.Count==0)
                    {
                        watchpath.oc.Add("此文件中无sheet表!");
                    }
                    //TargetSheetCombo.SelectedIndex = 0;
                }
                catch (Exception ex)
                {
                    watchpath.Path = "";
                    if (conn.State != ConnectionState.Open)
                        throw new Exception("未能读取Excel，原因如下：" + "\r\n" + ex.Message);
                    else if (conn.State == ConnectionState.Open)
                        throw new Exception(ex.Message);
                }
                finally
                {
                    conn.Close();
                }
            }
        }
        public class WatchPath
        {
            public delegate void WhenHavePath();
            public event WhenHavePath OnHavePath;
            string path;
            public ObservableCollection<string> oc { get; set; }
            public string ExcelType { get; set; }
            public string Path
            {
                get { return path; }
                set
                {
                    if (path!=value&&value!="")
                    {
                        path = value;
                        OnHavePath();
                    }
                    if(value == "")
                    {
                        path = value;
                    }
                }
            }
        }
        private void Label_PreviewMouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            try
            {
                var LabelItem = sender as Label;
                string content = LabelItem.Content.ToString();
                switch (content)
                {
                    case "Head Area":
                        hatbv.VisibilityP = Visibility.Visible;
                        hacbv.VisibilityP = Visibility.Collapsed;
                        HeadAreaBox.Focus();
                        break;
                    case "Key Column":
                        kctbv.VisibilityP = Visibility.Visible;
                        kccbv.VisibilityP = Visibility.Collapsed;
                        KeyColumnBox.Focus();
                        break;
                    case "Ignore Column":
                        ictbv.VisibilityP = Visibility.Visible;
                        iccbv.VisibilityP = Visibility.Collapsed;
                        IgnoreColumnBox.Focus();
                        break;
                    case "Save Folder":
                        System.Windows.Forms.FolderBrowserDialog m_Dialog = new System.Windows.Forms.FolderBrowserDialog();
                        System.Windows.Forms.DialogResult result = m_Dialog.ShowDialog();
                        if (result == System.Windows.Forms.DialogResult.Cancel)
                        {
                            return;
                        }
                        else
                        {
                            SaveFolder.Text = m_Dialog.SelectedPath.Trim();
                        }
                        break;
                    //case "Data Area":
                    //    datbv.VisibilityP = Visibility.Visible;
                    //    dacbv.VisibilityP = Visibility.Collapsed;
                    //    DataAreaBox.Focus();
                    //    break;
                }

            }
            catch
            {

            }
        }
        private void TextBox_LostFocus(object sender, RoutedEventArgs e)
        {
            try
            {
                var TextBoxItem = sender as TextBox;
                string content = TextBoxItem.Name.ToString();
                switch (content)
                {
                    case "HeadAreaBox":
                        hatbv.VisibilityP = Visibility.Collapsed;
                        hacbv.VisibilityP = Visibility.Visible;
                        CheckNewItem(ref headareacollection, TextBoxItem.Text,HeadAreaCombo);
                        break;
                    case "KeyColumnBox":
                        kctbv.VisibilityP = Visibility.Collapsed;
                        kccbv.VisibilityP = Visibility.Visible;
                        CheckNewItem(ref keycolumncollection, TextBoxItem.Text,KeyColumnCombo);
                        break;
                    case "IgnoreColumnBox":
                        ictbv.VisibilityP = Visibility.Collapsed;
                        iccbv.VisibilityP = Visibility.Visible;
                        CheckNewItem(ref ignorecolumncollection, TextBoxItem.Text,IgnoreColumnCombo);
                        break;
                    //case "DataAreaBox":
                    //    datbv.VisibilityP = Visibility.Collapsed;
                    //    dacbv.VisibilityP = Visibility.Visible;
                    //    CheckNewItem(ref dataareacollection, TextBoxItem.Text, DataAreaCombo);
                    //    break;
                }

            }
            catch
            {

            }
        }
        public void CheckNewItem(ref ObservableCollection<string> collection,string item,ComboBox combobox)
        {
            string tempitem = item.ToUpper();
            if (!collection.Contains(tempitem))
            {
                collection.Add(tempitem);
                combobox.SelectedIndex = collection.Count - 1;
            }
        }
        private void TextBox_PreviewDragOver(object sender, DragEventArgs e)
        {
            e.Effects = DragDropEffects.Copy;
            e.Handled = true;
        }
        private void TextBox_PreviewDrop(object sender, DragEventArgs e)
        {
            foreach (string f in (string[])e.Data.GetData(DataFormats.FileDrop))
            {
                oldadd.Text = f;
            }
        }

        private void Window_MouseDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            tb.Focus();
        }
    }


    public class TextBoxVisibility:INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private Visibility visibility=Visibility.Collapsed;
        public Visibility VisibilityP
        {
            get { return visibility; }
            set
            {
                visibility = value;
                if(this.PropertyChanged!=null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("VisibilityP"));
                }
            }
        }
    }

    public class ComboBoxiVisibility : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private Visibility visibility = Visibility.Visible;
        public Visibility VisibilityP
        {
            get { return visibility; }
            set
            {
                visibility = value;
                if (this.PropertyChanged != null)
                {
                    this.PropertyChanged.Invoke(this, new PropertyChangedEventArgs("VisibilityP"));
                }
            }
        }
    }

    public static class ExcelMgr
    {
        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int ProcessId);
        public static void KillExcelApp(this Excel.Application app)
        {
            app.Quit();
            IntPtr intptr = new IntPtr(app.Hwnd);
            GetWindowThreadProcessId(intptr, out int id);
            var p = Process.GetProcessById(id);
            if (p != null)
            p.Kill();
        }
    }

    public class Parameter
    {
        public List<string> Sheetlist { get; set; }
        public string OldAdd { get; set; }
        public string SavePath { get; set; }
        public string NewAdd { get; set; }
    }
}
