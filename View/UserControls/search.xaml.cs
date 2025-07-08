using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using Forms = System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data.SqlTypes;

namespace DocumentsModule.View.UserControls
{
    /// <summary>
    /// Interaction logic for search.xaml
    /// </summary>
    public partial class search : UserControl
    {
        
        List<User> users;
        Dictionary<int, string> projects = PublicFuncsNvars.projects;
        public ObservableCollection<string> Items { get; set; }
        private ICollectionView _collectionView;
        Dictionary<string, int> tkufaKV = new Dictionary<string, int>();
        public search()
        {
            users = PublicFuncsNvars.users;
            InitializeComponent();
            Init_tkufaKV();
            anafcombobox.SelectedIndex = 0;
            SugMehutav.SelectedIndex = 0;
            hufatz.SelectedIndex = 0;
            tkufa.SelectedIndex = tkufaKV["y"]; // too hard coded
            Tokzaot.SelectedIndex = 1;
            foreach (User u in users)
                if (u.getFullName().Trim() != "")
                {
                    usercombobox.Items.Add(u.getFullName());
                    mehutavcombobox.Items.Add(u.getFullName()+" - "+u.job);
                }
                    
            var sorteditems = usercombobox.Items
                .Cast<object>()
                .Select(item => item as ComboBoxItem ?? new ComboBoxItem { Content = item })
                .OrderBy(item => item.Content.ToString(), new HebrewStringComparer())
                .ToList();
            usercombobox.Items.Clear();
            //mehutavcombobox.Items.Clear();
            foreach (var item in sorteditems)
            {
                usercombobox.Items.Add(item);
            }
            usercombobox.Text = PublicFuncsNvars.curUser.getFullName();
            var sorteditemsM = mehutavcombobox.Items
                .Cast<object>()
                .Select(item => item as ComboBoxItem ?? new ComboBoxItem { Content = item })
                .OrderBy(item => item.Content.ToString(), new HebrewStringComparer())
                .ToList();
            mehutavcombobox.Items.Clear();
            //mehutavcombobox.Items.Clear();
            foreach (var item in sorteditemsM)
            {
                mehutavcombobox.Items.Add(item);
            }

            List<Folder> directories = PublicFuncsNvars.folders;
            tikcombobox.ItemsSource=(directories.Select(item => item.id).ToList());
            
            foreach (KeyValuePair<int, string> p in projects)
                procombobox.Items.Add(p.Key.ToString());
        }

        private void Init_tkufaKV()
        {
            for (int i = 0; i < tkufa.Items.Count; i++)
                tkufaKV.Add(((ComboBoxItem)tkufa.Items[i]).Tag.ToString(), i);
        }

        private void tbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbSearch.Text))
            {
                tbPlaceHolder.Visibility = Visibility.Visible;
                //if (FromShoteftx.Text.Length <= 1 && toShoteftx.Text.Length <= 1)
                if (FromShoteftx.Text.Length == 1 && toShoteftx.Text.Length == 1)
                {
                    FromShoteftx.Text = "";
                    toShoteftx.Text = "";
                }
                else if (FromShoteftx.Text == toShoteftx.Text && (FromShoteftx.Text.Length > 1))
                {
                    FromShoteftx.Text = "";
                    toShoteftx.Text = "";
                }
                else
                {
                    Nadontx.Text = tbSearch.Text;
                }
            }
              
            else
            {
                tbPlaceHolder.Visibility = Visibility.Hidden;
                usercombobox.Text = string.Empty;
                tkufa.SelectedIndex = tkufaKV["all"];
                int res1;
                bool ok1 = int.TryParse(tbSearch.Text, out res1);
                if (ok1)// כתבנו מספר
                {
                    FromShoteftx.Text = tbSearch.Text;
                    toShoteftx.Text = tbSearch.Text;
                    ispail.IsChecked = true;
                    Nadontx.Text = "";
                    isMainDoc.IsChecked = null;
                    isAttDoc.IsChecked = null;
                }
                else // כתבנו מחרוזת
                {
                    Nadontx.Text = tbSearch.Text;
                    if (FromShoteftx.Text == toShoteftx.Text)
                    {
                        FromShoteftx.Text = "";
                        toShoteftx.Text = "";
                    }
                }
                    
            }
        }

        internal void DefaultSearch()
        {
            tbSearch.Clear();
            tbSearch.Focus();
            Nadontx.Text = "";
            FromShoteftx.Text = "";
            toShoteftx.Text = "";
            tkufa.SelectedIndex = tkufaKV["y"];
            usercombobox.Text = PublicFuncsNvars.curUser.getFullName();
            isMainDoc.IsChecked = null;
            isAttDoc.IsChecked = null;
            SearchButton_Click(null, null);
        }

        private void btnClear_Click(object sender, RoutedEventArgs e) // x button
        {
            ResetSearch();
            SearchButton_Click(null, null);
            //ClearSearchAdvanced_Click(sender, e);
        }

        private void ResetSearch()
        {
            anafcombobox.SelectedIndex = 0;
            tkufa.SelectedIndex = tkufaKV["all"];
            Nadontx.Text = "";
            FromShoteftx.Text = "";
            toShoteftx.Text = "";
            usercombobox.SelectedIndex = -1;
            tikcombobox.Text = "";
            procombobox.Text = "";
            SugMehutav.SelectedIndex = 0;
            mehutavcombobox.Text = "";
            ispail.IsChecked = false;
            hufatz.SelectedIndex = 0;
            isMainDoc.IsChecked = null;
            isAttDoc.IsChecked = null;
            tbSearch.Text = string.Empty;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            anafcombobox.SelectedIndex = 0;
            //tkufa.SelectedIndex = 6;
            AdvancedSearchPopup.Visibility = Visibility.Visible;
            //ads.IsOpen = true;
            AdvancedSearchPopup.IsOpen = true;
        }

        private void UserListBox_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {

        }
        

        private void ClearSearchAdvanced_Click(object sender, RoutedEventArgs e)
        {
            ResetSearch();

            //UserTextBox.Text = "";
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            AdvancedSearchPopup.Visibility = Visibility.Hidden;
            AdvancedSearchPopup.IsOpen = false;
            Mouse.OverrideCursor = Cursors.Wait;
            int Proj = 0;
            int FromShotef = 0;
            int toShotef = 0;
            string fromDate = "";
            string toDate = "";
            string Nadon = "";
            
            string SholeahTeur = usercombobox.Text;// UserTextBox.Text;
            /*int SholeahKod = 0;
            
            bool isSholeahKod = int.TryParse(usercombobox.Text, out SholeahKod);
            if (isSholeahKod)
            {
                SholeahTeur = "";
            }*/
            string Mehutav = mehutavcombobox.Text;
            int IsLePeula = SugMehutav.SelectedIndex;
            int IsPail = (ispail.IsChecked==true) ?1: 0;// IsPail. ? 1 : SqlInt32.Null;
            int? IsMainDoc = GetNullableCB(isMainDoc.IsChecked);
            int? IsAttDoc = GetNullableCB(isAttDoc.IsChecked);
            int Tik;
            try
            {
                Tik = int.Parse(tikcombobox.Text);
            }
            catch
            {
                Tik = 0;
            }
            int IsHufatz = hufatz.SelectedIndex;
            int SugMismach = 0;
            int top = Tokzaot.Text=="הכל" ? 999999999 : int.Parse(Tokzaot.Text);
            char anaf = '\0';
            if (!string.IsNullOrEmpty(FromShoteftx.Text) && !string.IsNullOrEmpty(toShoteftx.Text))
            {
                int res1 = 0, res2;
                if (int.TryParse(FromShoteftx.Text, out res1) || FromShoteftx.Text == "")
                    FromShotef = res1;
                else
                    res1 = -1;

                if (int.TryParse(toShoteftx.Text, out res2))
                    toShotef = res2;
                else if (toShoteftx.Text == "")
                    res2 = 0;
                else
                    res2 = -1;

                if (res1 != -1 && res2 != -1)
                {
                    if (res1 > res2)
                    {
                        Forms.MessageBox.Show("לא ניתן להזין שוטף סיום קטן יותר משוטף התחלה", "חיפוש שוטפים שגוי", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information,
                            Forms.MessageBoxDefaultButton.Button1, Forms.MessageBoxOptions.RightAlign | Forms.MessageBoxOptions.RtlReading);
                        return;// -1;
                    }
                }
                else
                {
                    Forms.MessageBox.Show("שוטף יכול להכיל רק מספרים", "חיפוש שוטפים שגוי", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information,
                            Forms.MessageBoxDefaultButton.Button1, Forms.MessageBoxOptions.RightAlign | Forms.MessageBoxOptions.RtlReading);
                    return;// -1;
                }
            }
            else
            {
                //COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMe", "0"));
                //COMMAND.Parameters.Add(new SqlParameter("@P_ShotefAd", "0"));
            }
            short res14 = -1;
            if (short.TryParse(procombobox.Text, out res14))
            {
                if (!projects.ContainsKey(res14))
                {
                    Forms.MessageBox.Show("אין פרויקט עם מספר מזהה זה", "חיפוש מפרויקט שגוי", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information,
                            Forms.MessageBoxDefaultButton.Button1, Forms.MessageBoxOptions.RightAlign | Forms.MessageBoxOptions.RtlReading);
                    return ;
                }
                Proj=res14;
            }
            else
                Proj = 0;

            DateTime today = DateTime.Today;
            toDate= tkufa.SelectedIndex != tkufaKV["all"] ? today.ToString("yyyyMMdd") : "";
            if (tkufa.SelectedIndex != tkufaKV["all"])
            {
                if (tkufa.SelectedIndex == tkufaKV["d"])
                    fromDate= today.ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["w"])
                    fromDate = today.AddDays(-7).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["m"])
                    fromDate = today.AddMonths(-1).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["2m"])
                    fromDate = today.AddMonths(-2).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["3m"])
                    fromDate = today.AddMonths(-3).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["6m"])
                    fromDate = today.AddMonths(-6).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["y"])
                    fromDate = today.AddYears(-1).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["2y"])
                    fromDate = today.AddYears(-2).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == tkufaKV["3y"])
                    fromDate = today.AddYears(-3).ToString("yyyyMMdd");
                else
                    fromDate = today.AddYears(-5).ToString("yyyyMMdd");
            }
            else
                fromDate = "";
            Nadon = Nadontx.Text;
            anaf=!anafcombobox.Text.Equals("הכל") ? PublicFuncsNvars.getBranchByString(anafcombobox.Text) : '\0';
            
            int numberOfRows = DataGridDocs.SearchDocs(FromShotef, toShotef, fromDate, toDate, Nadon, SholeahTeur, Mehutav, IsLePeula, IsPail, Tik, IsHufatz, SugMismach, top, anaf,Proj, IsMainDoc, IsAttDoc);
            if (numberOfRows == 0)
            {
                Forms.MessageBox.Show("לא נמצאו מסמכים התואמים את נתוני החיפוש", "", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information,
                        Forms.MessageBoxDefaultButton.Button1, Forms.MessageBoxOptions.RightAlign | Forms.MessageBoxOptions.RtlReading);
            }
            else
            {
                results.Content = "נמצאו " + numberOfRows + " מסמכים";
            }
            Mouse.OverrideCursor = null;
        }

        private int? GetNullableCB(bool? isChecked)
        {
            if (isChecked == null) return null;
            return Convert.ToInt32(isChecked);
        }

        private void FromShoteftx_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                int res1;
                bool ok1 = int.TryParse(FromShoteftx.Text, out res1);


                if (ok1)
                {
                    toShoteftx.Text = FromShoteftx.Text;
                    if (FromShoteftx.Text == toShoteftx.Text)
                        tbSearch.Text = toShoteftx.Text;
                    else
                        tbSearch.Text = Nadontx.Text;
                    //comboBox8.SelectedIndex = 10;
                }
                else
                {
                    if (FromShoteftx.Text == toShoteftx.Text)
                        tbSearch.Text = toShoteftx.Text;
                    else
                        tbSearch.Text = Nadontx.Text;
                    FromShoteftx.Text = "";
                    toShoteftx.Text = "";
                }
                //else
//comboBox8.SelectedIndex = 6;
            }

            catch { }
            
        }

        private void Nadontx_TextChanged(object sender, TextChangedEventArgs e)
        {
            int res1;
            bool ok1 = int.TryParse(FromShoteftx.Text, out res1);
            if (ok1)
            {
                if (FromShoteftx.Text != toShoteftx.Text)
                    tbSearch.Text = Nadontx.Text;
                /*toShoteftx.Text = FromShoteftx.Text;
                if (FromShoteftx.Text != toShoteftx.Text)
                    tbSearch.Text = Nadontx.Text;*/
            }
            else
                tbSearch.Text = Nadontx.Text;
        }

        private void Grid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                AdvancedSearchPopup.Visibility = Visibility.Hidden;
                AdvancedSearchPopup.IsOpen = false;
                SearchButton_Click(sender, e);
            }
        }

        /*private void UserTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var searchText = UserTextBox.Text;
            _collectionView.Filter = item =>
             {
                 return ((string)item).ToLower().Contains(searchText);
             };
            usercombobox.IsDropDownOpen = true;
        }
        */
        private void usercombobox_Loaded(object sender, RoutedEventArgs e)
        {
            /*var combobox = sender as ComboBox;
            var sorteditems = combobox.Items
                .Cast<object>()
                .Select(item => item as ComboBoxItem ?? new ComboBoxItem { Content = item })
                .OrderBy(item => item.Content.ToString(), new HebrewStringComparer())
                .ToList();
            combobox.Items.Clear();
            foreach (var item in sorteditems)
                combobox.Items.Add(item);*/
            /*var textbox = combobox.Template.FindName("PART_EditableTextBox", combobox) as TextBox;
            if (textbox != null)
            {
                textbox.TextChanged += ComboBox_TextChanged;
            }*/
        }
        private void ComboBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            var textbox = sender as TextBox;
            var combobox = (ComboBox)textbox.Parent;
            string searchText = combobox.Text;
            foreach(ComboBoxItem item in combobox.Items)
            {
                if (item is ComboBoxItem )
                {
                    item.Visibility = item.Content.ToString().Contains(searchText) ? Visibility.Visible : Visibility.Collapsed;

                }
            }
        }

        private void IsPail_Checked(object sender, RoutedEventArgs e)
        {

        }

        private void tikcombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (procombobox.Text != "" && tikcombobox.Text!="")
                procombobox.Text = "";
        }

        private void procombobox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (procombobox.Text != "" && tikcombobox.Text != "")
                tikcombobox.Text = "";
        }
    }
    public class HebrewStringComparer : IComparer<string>
    {
        public int Compare(string x, string y)
        {
            return string.Compare(x, y, new System.Globalization.CultureInfo("he-IL"), System.Globalization.CompareOptions.None);
        }
    }
}
