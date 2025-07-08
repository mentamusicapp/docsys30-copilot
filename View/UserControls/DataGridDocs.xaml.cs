using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlTypes;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Threading;
using System.Threading;
using WinForms = System.Windows.Forms;
using System.Globalization;
using System.IO;

namespace DocumentsModule.View.UserControls
{
    /// <summary>
    /// Interaction logic for DataGridDocs.xaml
    /// </summary>
    public partial class DataGridDocs : UserControl
    {
        public static DataTable dataTable;
        public static DataTable AllDataTable;
        public static DataGrid StaticDatagrid;
        public static DataGrid StaticDatagridHanhayot;
        public static DataGrid StaticDatagridPladot;
        public static DataGrid StaticDatagridAshala;
        public static DataGrid StaticDatagridBanam;
        public static DataGrid StaticDatagridPituaj;
        public static DataGrid StaticDatagridDocs;
        public static DataGrid CurrentDataGrid;
        internal DocumentHandling dh = null;
        public static int shotefToDelete=0;
        int akol = 0;
        int filterHanayot = 0;
        int filterPladot = 0;
        int filterHashala = 0;
        int filterbanam = 0;
        int filterpituaj = 0;
        int filterDocs = 0;
        public static string FilterOfNispah = "";
        public static List<int> ListOfNispah = new List<int>();
        public static string currentHeader="";
        public DataGridDocs()
        {
            InitializeComponent();
            StaticDatagrid = dataGridDocs;
            StaticDatagridHanhayot = dataGridDocs1;
            StaticDatagridPladot = dataGridDocs2;
            StaticDatagridAshala = dataGridDocs3;
            StaticDatagridBanam = dataGridDocs4;
            StaticDatagridPituaj = dataGridDocs5;
            StaticDatagridDocs = dataGridDocs6;
            _clickTimer = new DispatcherTimer
            {
                Interval = _doubleClickthreshold
            };
            _clickTimer.Tick += ClickTimer_Tick;
        }

        private void dataGridDocs_Loaded(object sender, RoutedEventArgs e)
        {
            if (akol != 0)
                return;
            else
                akol++;
            DateTime today = DateTime.Today;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            conn.Open();
            SqlCommand COMMAND = new SqlCommand("SP_GetDocsList_3", conn);
            COMMAND.CommandType = CommandType.StoredProcedure;

            COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMe", "0"));
            COMMAND.Parameters.Add(new SqlParameter("@P_ShotefAd", "0"));
            COMMAND.Parameters.Add(new SqlParameter("@P_TaarichAd", today.ToString("yyyyMMdd")));
            COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", today.AddYears(-1).ToString("yyyyMMdd")));
            COMMAND.Parameters.Add(new SqlParameter("@P_Proyekt", SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_Nadon", ""));
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahTeur", PublicFuncsNvars.curUser.getFullName()));
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahKod", PublicFuncsNvars.curUser.userCode));
            COMMAND.Parameters.Add(new SqlParameter("@P_Mehutav", ""));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsLePeula", SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsLoPail", SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsHufats", SqlInt32.Null));

            COMMAND.Parameters.Add(new SqlParameter("@P_IsDocFile", SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsNispachFile", SqlInt32.Null));

            //COMMAND.Parameters.Add(new SqlParameter("@P_SugMismach", "0"));
            COMMAND.Parameters.Add(new SqlParameter("@P_Top", 1000));//פעם הראשונה תמיד עד 1000 תוצאות
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahAnaf", '\0'));
            COMMAND.Parameters.Add(new SqlParameter("@P_Tik", "0"));
            //FilterString = "";
            try
            {
                using (SqlDataReader reader = COMMAND.ExecuteReader())
                {
                    /*while (reader.Read())
                    {
                        string shotef = reader.GetInt32(0).ToString();
                        MessageBox.Show(shotef);
                    }*/
                    dataTable = new DataTable();
                    dataTable.Load(reader);
                    conn.Close();
                    AllDataTable = dataTable;
                    try
                    {
                        dataGridDocs.ItemsSource = CleanDataTable(dataTable).Select("Nispah = 0").CopyToDataTable().DefaultView;
                    }
                    catch (Exception exx)
                    {
                        dataGridDocs.DataContext = CleanDataTable(dataTable).DefaultView;
                    }
                    DocumentsSearch.documents = dataTable.AsEnumerable().Select(row =>
                    {
                        return new KeyValuePair<int, int>(row.Field<int>("Shotef"), row.Field<int>("SholeahKod"));
                    }).ToList();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void dataGridDocs_MouseDoubleClick(object sender, MouseButtonEventArgs e)//not in use
        {

            if (getDataGrid().SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)getDataGrid().SelectedItem;

                int id = Convert.ToInt32(selectedRow["Shotef"]);
                int nispah = Convert.ToInt32(selectedRow["Nispah"]);

                if (nispah == 0)
                {
                    if (!PublicFuncsNvars.dhFormsOpen.Contains(id))
                    {
                        KeyValuePair<int, int> d = getDocById(id);
                        if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                        {
                            Thread docHandleThread = new Thread(openDocumentHandlingForm);
                            docHandleThread.SetApartmentState(ApartmentState.STA);
                            docHandleThread.Start(d.Key);
                        }
                        else
                        {
                            WinForms.MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information,
                                WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                        }
                    }
                    else
                    {
                        WinForms.MessageBox.Show("המסך של מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסך מספר פעמים", "מסך פתוח", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error,
                            WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RtlReading | WinForms.MessageBoxOptions.RightAlign);
                    }
                }
                else
                {
                    ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(id, nispah));// id));//rowIndex-e.RowIndex
                }

            }
        }
        private void openDocumentHandlingForm(object obj)
        {
            int d = (int)obj;
            dh = new DocumentHandling(d);
            dh.Activate();
            dh.ShowDialog();
        }
        private void viewAtt(object idObj)
        {
            KeyValuePair<int, int> ids = (KeyValuePair<int, int>)idObj;
            PublicFuncsNvars.viewAtt(ids.Key, ids.Value);
        }
        private KeyValuePair<int, int> getDocById(int id)
        {
            foreach (KeyValuePair<int, int> doc in DocumentsSearch.documents)
            {
                if (doc.Key == id)
                    return doc;
            }
            return new KeyValuePair<int, int>(-1, -1);
        }

        private void TextBlock_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)//one click in the nadon for open the טיפול מסמך
        {
            if (getDataGrid().SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)getDataGrid().SelectedItem;

                int id = Convert.ToInt32(selectedRow["Shotef"]);
                int nispah = Convert.ToInt32(selectedRow["Nispah"]);
                if (nispah == 0)
                {
                    if (!PublicFuncsNvars.dhFormsOpen.Contains(id))
                    {
                        KeyValuePair<int, int> d = getDocById(id);
                        if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                        {
                            Thread docHandleThread = new Thread(openDocumentHandlingForm);
                            docHandleThread.SetApartmentState(ApartmentState.STA);
                            docHandleThread.Start(d.Key);
                        }
                        else
                        {
                            WinForms.MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information,
                                WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                        }
                    }
                    else // emily lutvak - cancel the message when clicking the nadon and the window already open to avoid the message on double-click. 31/10/2024
                    {
                       /* WinForms.MessageBox.Show("המסך של מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסך מספר פעמים", "מסך פתוח", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error,
                            WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RtlReading | WinForms.MessageBoxOptions.RightAlign);*/
                    }
                }
                else
                {
                    ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(id, nispah));// id));//rowIndex-e.RowIndex
                }
            }
        }

        private DateTime _lastClickTime;
        private readonly TimeSpan _doubleClickthreshold = TimeSpan.FromMilliseconds(300);
        private DispatcherTimer _clickTimer;

        private void ClickTimer_Tick(object sender, EventArgs e)
        {
            /*_clickTimer.Stop();
            TextBlock_MouseLeftButtonDown(sender, e);*/
        }
        private void TextBlock_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)//double click in the rest of the row for open the טיפול מסמך
        {
            DateTime now = DateTime.Now;
            if (now - _lastClickTime <= _doubleClickthreshold)
            {
                _lastClickTime = DateTime.MinValue;
                if (getDataGrid().SelectedItem != null)
                {
                    DataRowView selectedRow = (DataRowView)getDataGrid().SelectedItem;

                    int id = Convert.ToInt32(selectedRow["Shotef"]);
                    int nispah = Convert.ToInt32(selectedRow["Nispah"]);

                    if (nispah == 0)
                    {
                        if (!PublicFuncsNvars.dhFormsOpen.Contains(id))
                        {
                            KeyValuePair<int, int> d = getDocById(id);
                            if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                            {
                                Thread docHandleThread = new Thread(openDocumentHandlingForm);
                                docHandleThread.SetApartmentState(ApartmentState.STA);
                                docHandleThread.Start(d.Key);
                            }
                            else
                            {
                                WinForms.MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information,
                                    WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                            }
                        }
                        else
                        {
                            WinForms.MessageBox.Show("המסך של מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסך מספר פעמים", "מסך פתוח", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Error,
                                WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RtlReading | WinForms.MessageBoxOptions.RightAlign);
                        }
                    }
                    else
                    {
                        ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(id, nispah));// id));//rowIndex-e.RowIndex
                    }
                }
            }
            else
                _lastClickTime = now;
        }
        public static int SearchDocs(int fromShoef, int toShotef, string fromDate, string toDate, string Nadon, string SholeahTeur, string Mehutav, int IsLePeula, int IsPail, int Tik, int IsHufatz, int SugMismach, int top, char anaf,int Proj, int? isMainDoc, int? isAttDoc)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            conn.Open();
            SqlCommand COMMAND = new SqlCommand("SP_GetDocsList_3", conn);
            COMMAND.CommandType = CommandType.StoredProcedure;

            COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMe", fromShoef));
            COMMAND.Parameters.Add(new SqlParameter("@P_ShotefAd", toShotef));
            COMMAND.Parameters.Add(new SqlParameter("@P_TaarichMe", fromDate));
            COMMAND.Parameters.Add(new SqlParameter("@P_TaarichAd", toDate));
            COMMAND.Parameters.Add(new SqlParameter("@P_Proyekt",(Proj!=0) ? Proj : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_Nadon", Nadon));
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahTeur", SholeahTeur));
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahKod", SholeahKod));
            COMMAND.Parameters.Add(new SqlParameter("@P_Mehutav", Mehutav));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsLePeula", (IsLePeula!=0) ? IsLePeula - 1 : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsLoPail",IsPail==1? 1:SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_Tik", Tik));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsHufats", (IsHufatz != 0) ? IsHufatz - 1 : SqlInt32.Null)); //IsHufatz));

            COMMAND.Parameters.Add(new SqlParameter("@P_IsDocFile", isMainDoc != null ? (int)isMainDoc : SqlInt32.Null));
            COMMAND.Parameters.Add(new SqlParameter("@P_IsNispachFile", isAttDoc != null ? (int)isAttDoc : SqlInt32.Null));

            //COMMAND.Parameters.Add(new SqlParameter("@P_SugMismach", SugMismach));
            COMMAND.Parameters.Add(new SqlParameter("@P_Top", top));
            COMMAND.Parameters.Add(new SqlParameter("@P_SholeahAnaf", anaf));

            
            //FilterString = "";
            using (SqlDataReader reader = COMMAND.ExecuteReader())
            {
                dataTable = new DataTable();
                dataTable.Load(reader);
                AllDataTable = dataTable;
                //dataGridDocs.DataContext = dataTable.DefaultView;
                conn.Close();
                DataTable datatt = new DataTable();
                try
                {
                    StaticDatagrid.ItemsSource = CleanDataTable(dataTable).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    StaticDatagrid.DataContext = CleanDataTable(dataTable).DefaultView;
                }

                try
                {
                    StaticDatagridHanhayot.ItemsSource = FilterHanhayot(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridHanhayot.ItemsSource = null;// = (new DataTable());
                        
                    /*else
                        StaticDatagridHanhayot.DataContext = FilterHanhayot(CleanDataTable(dataTable)).DefaultView;*/
                }
                try
                {
                    StaticDatagridAshala.ItemsSource = FilterAshala(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridAshala.ItemsSource = null;// = (new DataTable());

                    //else
                      //  StaticDatagridAshala.DataContext = FilterAshala(CleanDataTable(dataTable)).DefaultView;
                }
                try
                {
                    StaticDatagridPladot.ItemsSource = FilterPladot(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridPladot.ItemsSource = null;// = (new DataTable());

                    //else
                      //  StaticDatagridPladot.DataContext = FilterPladot(CleanDataTable(dataTable)).DefaultView;
                }
                try
                {
                    StaticDatagridBanam.ItemsSource = FilterBanam(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridBanam.ItemsSource = null;// = (new DataTable());

                    //else
                        //StaticDatagridBanam.DataContext = FilterBanam(CleanDataTable(dataTable)).DefaultView;
                }
                try
                {
                    StaticDatagridPituaj.ItemsSource = FilterPituah(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridPituaj.ItemsSource = null;// = (new DataTable());

                    //else
                      //  StaticDatagridPituaj.DataContext = FilterPituah(CleanDataTable(dataTable)).DefaultView;
                }
                try
                {
                    StaticDatagridDocs.ItemsSource = FilterDocs(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
                }
                catch (Exception exx)
                {
                    //if (exx.Message.Contains("Nispah"))
                        StaticDatagridDocs.ItemsSource = null;// = (new DataTable());

                    //else
                     //   StaticDatagridDocs.DataContext = FilterDocs(CleanDataTable(dataTable)).DefaultView;
                }
                try
                {
                    dataTable = dataTable.Select("Nispah = 0").CopyToDataTable();
                }
                catch { }
                DocumentsSearch.documents = dataTable.AsEnumerable().Select(row =>
                {
                    return new KeyValuePair<int, int>(row.Field<int>("Shotef"), row.Field<int>("SholeahKod"));
                }).ToList();
                FilterOfNispah = "";
                ListOfNispah = new List<int>();
                return DocumentsSearch.documents.Count;
            }
        }

        private void StackPanel_MouseEnter(object sender, MouseEventArgs e)
        {
            this.Cursor = Cursors.Hand;
        }

        public static DataTable CleanDataTable(DataTable dt)
        {
            if (dt== null) return null;
            try
            {
                dt.Columns.Add("NewTxt", typeof(string), "IIf(isRagish = 'true','',Txt)");
                dt.Columns.Add("NispahJustNumber", typeof(string), "IIf(Nispah = '0','',Nispah)");
                dt.AcceptChanges();
                dt.Columns.Add("Mehutav", typeof(string));
                dt.Columns["Mehutav"].ReadOnly = false;
                dt.AsEnumerable().ToList().ForEach(row => row["Mehutav"] = row["MechutavStr"].ToString().Replace(";", "\n"));
                return dt;
            }
            catch (Exception e)
            {
                return dt;
            }
        }
        
        private void viewDoc(object idObj)
        {
            //Mouse.OverrideCursor = Cursors.Wait;
            WinForms.Cursor.Current = WinForms.Cursors.WaitCursor;
            (new PublicFuncsNvars()).viewDoc((int)idObj);
            //Mouse.OverrideCursor = null;
        }

        private void ViewDocOnly_Click(object sender, RoutedEventArgs e)
        {
            DateTime now = DateTime.Now;
            
            if (getDataGrid().SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)getDataGrid().SelectedItem;

                int id = Convert.ToInt32(selectedRow["Shotef"]);
                int nispah = Convert.ToInt32(selectedRow["Nispah"]);
                //string hasAtt = selectedRow["HasAtt"].ToString();
                KeyValuePair<int, int> d = getDocById(id);
                if (PublicFuncsNvars.isAllowedToRagish(id) && (PublicFuncsNvars.isAuthorizedUser(d.Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id)))
                {

                    //if (PublicFuncsNvars.isAuthorizedUser(getDocById(id).Value, PublicFuncsNvars.curUser) || PublicFuncsNvars.isCurUserAllowedToWatchDoc(id))
                    if (nispah == 0)
                        ThreadPool.QueueUserWorkItem(viewDoc, id);
                    else
                        ThreadPool.QueueUserWorkItem(viewAtt, new KeyValuePair<int, int>(id, nispah));// id));//rowIndex-e.RowIndex
                }
                else
                {
                    WinForms.MessageBox.Show("אינך מורשה/ית לצפות במסמך זה.", "אין הרשאות", WinForms.MessageBoxButtons.OK, WinForms.MessageBoxIcon.Information,
                        WinForms.MessageBoxDefaultButton.Button1, WinForms.MessageBoxOptions.RightAlign | WinForms.MessageBoxOptions.RtlReading);
                }

                
            }
        }

        private void dataGridDocs_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key== Key.Delete && (Keyboard.Modifiers & ModifierKeys.Shift)==ModifierKeys.Shift&& shotefToDelete != 0)
             {
                MessageBoxResult x = MessageBox.Show("To Delete " + shotefToDelete + " ?", "מחיקת שוטף", MessageBoxButton.YesNo,MessageBoxImage.Warning);
                if (x == MessageBoxResult.Yes)
                {
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    SqlCommand cmd = new SqlCommand("dbo.SP_DeleteDocument", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add(new SqlParameter("@ShoteToDel", shotefToDelete));
                    conn.Open();
                    cmd.ExecuteNonQuery();
                    conn.Close();

                    MessageBox.Show("DELETED " + shotefToDelete);
                }
            }
        }

        private void dataGridDocs_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (getDataGrid().SelectedItem != null)
            {
                DataRowView selectedRow = (DataRowView)getDataGrid().SelectedItem;

                shotefToDelete = Convert.ToInt32(selectedRow["Shotef"]);
                
            }
        }
        public static DataTable FilterHanhayot(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 1 || f.Field<short>("SugMismach") == 99);
            DataTable d= filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
            return d;

        }
        public static DataTable FilterPladot(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 2 || f.Field<short>("SugMismach") == 99);
            return filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
        }
        
        public static DataTable FilterAshala(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 3 || f.Field<short>("SugMismach") == 99);
            return filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
        }
        public static DataTable FilterBanam(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 4 || f.Field<short>("SugMismach") == 99);
            return filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
        }
        public static DataTable FilterPituah(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 5 || f.Field<short>("SugMismach") == 99);
            return filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
        }
        public static DataTable FilterDocs(DataTable dt)
        {
            var filteredrows = dt.AsEnumerable()
                .Where(f => f.Field<short>("SugMismach") == 0 || f.Field<short>("SugMismach") == 99);
            return filteredrows.Any() ? filteredrows.CopyToDataTable() : new DataTable();
        }
        private void dataGridDocs1_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterHanayot != 0)
                return;
            else
                filterHanayot++;
            try
            {
                dataGridDocs1.ItemsSource = FilterHanhayot(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                string m = exx.Message;
                if (exx.Message.Contains("Nispah"))
                {
                    dataGridDocs1.DataContext = (new DataTable()).DefaultView;
                    return;
                }
                dataGridDocs1.DataContext = FilterHanhayot(CleanDataTable(dataTable)).DefaultView;
            }
        }

        private void dataGridDocs2_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterPladot != 0)
                return;
            else
                filterPladot++;
            try
            {
                dataGridDocs2.ItemsSource = FilterPladot(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                dataGridDocs2.DataContext = FilterPladot(CleanDataTable(dataTable)).DefaultView;
            }
        }

        private void dataGridDocs3_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterHashala != 0)
                return;
            else
                filterHashala++;
            try
            {
                dataGridDocs3.ItemsSource = FilterAshala(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                dataGridDocs3.DataContext = FilterAshala(CleanDataTable(dataTable)).DefaultView;
            }
        }

        private void dataGridDocs4_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterbanam != 0)
                return;
            else
                filterbanam++;
            try
            {
                dataGridDocs4.ItemsSource = FilterBanam(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                dataGridDocs4.DataContext = FilterBanam(CleanDataTable(dataTable)).DefaultView;
            }
        }

        private void dataGridDocs5_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterpituaj != 0)
                return;
            else
                filterpituaj++;
            try
            {
                dataGridDocs5.ItemsSource = FilterPituah(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                dataGridDocs5.DataContext = FilterPituah(CleanDataTable(dataTable)).DefaultView;
            }
        }

        private void dataGridDocs6_Loaded(object sender, RoutedEventArgs e)
        {
            if (filterDocs != 0)
                return;
            else
                filterDocs++;
            try
            {
                dataGridDocs6.ItemsSource = FilterDocs(CleanDataTable(dataTable)).Select("Nispah = 0").CopyToDataTable().DefaultView;
            }
            catch (Exception exx)
            {
                dataGridDocs6.DataContext = FilterDocs(CleanDataTable(dataTable)).DefaultView;
            }
        }
        private DataGrid getDataGrid()
        {
            TabItem selectedTab = TabControlDocs.SelectedItem as TabItem;
            if (selectedTab != null)
            {
                string selectedHeader = selectedTab.Header.ToString();
                switch (selectedHeader)
                {
                    case "הכל":
                        return dataGridDocs;
                    case "מסמכים":
                        return dataGridDocs6;
                    case "הנחיות":
                        return dataGridDocs1;
                    case "הקצאות (פלדות)":
                        return dataGridDocs2;
                    case "השאלות/הקצאות":
                        return dataGridDocs3;
                    case "בנ\"מ":
                        return dataGridDocs4;
                    case "השאלות (פיתוח)":
                        return dataGridDocs5;
                    default:
                        return dataGridDocs;
                }
            }
            return dataGridDocs;
        }
        private DataTable getDataTable()
        {
            TabItem selectedTab = TabControlDocs.SelectedItem as TabItem;
            if (selectedTab != null)
            {
                string selectedHeader = selectedTab.Header.ToString();
                switch (selectedHeader)
                {
                    case "הכל":
                        return AllDataTable;
                    case "מסמכים":
                        return FilterDocs(AllDataTable);
                    case "הנחיות":
                        return FilterHanhayot(AllDataTable);
                    case "הקצאות (פלדות)":
                        return FilterPladot(AllDataTable);
                    case "השאלות/הקצאות":
                        return FilterAshala(AllDataTable);
                    case "בנ\"מ":
                        return FilterBanam(AllDataTable);
                    case "השאלות (פיתוח)":
                        return FilterPituah(AllDataTable);
                    default:
                        return AllDataTable;
                }
            }
            return AllDataTable;
        }
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            TabControl tabcontrol = sender as TabControl;
             if (tabcontrol!= null)
            {
                
                TabItem selectedTab = TabControlDocs.SelectedItem as TabItem;
                if (selectedTab != null)
                {
                    string selectedHeader = selectedTab.Header.ToString();
                    switch (selectedHeader)
                    {
                        case "הכל":
                            CurrentDataGrid=StaticDatagrid;
                            break;
                        case "מסמכים":
                            CurrentDataGrid = StaticDatagrid;
                            break;
                        case "הנחיות":
                            CurrentDataGrid = StaticDatagrid;
                            break;
                        case "הקצאות (פלדות)":
                            CurrentDataGrid = StaticDatagrid;
                            break;
                        case "בנ\"מ":
                            CurrentDataGrid = StaticDatagrid;
                            break;
                        case "השאלות (פיתוח)":
                            CurrentDataGrid = StaticDatagrid;
                            break;
                    }
                    if (currentHeader != selectedHeader)
                    {
                        currentHeader = selectedHeader;
                        FilterOfNispah = "";
                        ListOfNispah = new List<int>();
                    }
                }
                
            }
        }

        private void OpenNispah_Click(object sender, RoutedEventArgs e)
        {

        }

        private void OpenNispah_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            DataGrid DataGridOpen = getDataGrid();
            DataTable DataTableOpen = getDataTable();
            if (DataGridOpen.SelectedItem != null)
            {
                
                DataRowView selectedRow = (DataRowView)DataGridOpen.SelectedItem;
                DataRowView rows = (DataRowView)DataGridOpen.Items[0];
                string sug = selectedRow["SugMismach"].ToString();
                sug = rows["SugMismach"].ToString();
                int id = Convert.ToInt32(selectedRow["Shotef"]);
                //if (int.Parse(selectedRow["Nispah"].ToString()) > 0) // ASAF MOR 23.09.24
                //{
                    
                //}
                if (selectedRow["HasAtt"].ToString() == "+")
                {
                    if (!ListOfNispah.Contains(id))
                    {
                        ListOfNispah.Add(id);
                        FilterOfNispah += FilterOfNispah == "" ? $"Nispah = 0 or Shotef='{id}'" : $" or Shotef='{id}'";
                        
                    }
                }
                else if (selectedRow["HasAtt"].ToString() == "-")
                {
                    FilterOfNispah = FilterOfNispah.Replace($"or Shotef='{id}'", "");
                    ListOfNispah.Remove(id);
                    FilterOfNispah += FilterOfNispah == "" ? "Nispah = 0" : "";

                }
                DataTable dt = DataTableOpen.Select(FilterOfNispah).CopyToDataTable();

                WinForms.BindingSource bs = new WinForms.BindingSource();
                bs.DataSource = dt;
                DataGridOpen.ItemsSource = bs;
                dt.DefaultView.Sort = "Shotef DESC, Nispah ASC";

                foreach (int idNispah in ListOfNispah)
                {
                    int indexRow = DataGridOpen.ItemsSource.Cast<DataRowView>()
                      .Select((rowv, idx) => new { RowView = rowv, Index = idx }).FirstOrDefault(x => x.RowView.Row.Field<int>("Shotef") == idNispah)?.Index ?? -1;
                    if (indexRow != -1)
                    {
                        DataRowView row = (DataRowView)DataGridOpen.Items[indexRow];
                        row["HasAtt"] = "-";
                    }
                }
                DataView datav = dt.DefaultView;

            }
        }



        private void StackPanel_MouseLeave(object sender, MouseEventArgs e)
        {
            this.Cursor = null;
        }
        /*public class MainWindowViewModel
{
public ICommand ShowMessageCommand { get; }
public ICommand DoubleClickCommand { get; }
public MainWindowViewModel(DataTable dataTable)
{
ShowMessageCommand = new RelayCommand(ShowMessageCommand);
DoubleClickCommand = new RelayCommand(DoubleClickCommand);
}
}
public class RelayCommand : ICommand
{
private readonly Action<object> _execute;
private readonly Func<object, bool> _canExecute;

public event EventHandler CanExecuteChanged;
public RelayCommand(Action<object> execute, Func<object,bool> canExecute = null)
{
_execute = execute ?? throw new ArgumentException(nameof(execute));
_canExecute = canExecute;
}

public bool CanExecute(object parameter) => _canExecute?.Invoke(parameter) ?? true;
public void Execute(object parameter) => _execute(parameter);
public void RaiseCanExecuteChanged() => CanExecuteChanged?.Invoke(this, EventArgs.Empty);
}*/
    }
    
}

