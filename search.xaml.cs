using System;
using System.Collections.Generic;
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
        public search()
        {
            InitializeComponent();
            anafcombobox.SelectedIndex = 0;
            tkufa.SelectedIndex = 6;
            Top.SelectedIndex = 1;
        }

        private void tbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (string.IsNullOrEmpty(tbSearch.Text))
            {
                tbPlaceHolder.Visibility = Visibility.Visible;
                if (FromShoteftx.Text.Length <= 1 && toShoteftx.Text.Length <= 1)
                {
                    FromShoteftx.Text = "";
                    toShoteftx.Text = "";
                }
            }
                
            else
            {
                tbPlaceHolder.Visibility = Visibility.Hidden;
                int res1;
                bool ok1 = int.TryParse(tbSearch.Text, out res1);
                if (ok1)
                {
                    FromShoteftx.Text = tbSearch.Text;
                    toShoteftx.Text = tbSearch.Text;
                }
                else
                    Nadontx.Text = tbSearch.Text;
            }
                
        }

        private void btnClear_Click(object sender, RoutedEventArgs e)
        {
            tbSearch.Clear();
            tbSearch.Focus();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            anafcombobox.SelectedIndex = 0;
            tkufa.SelectedIndex = 6;
            Top.SelectedIndex = 1;
            AdvancedSearchPopup.IsOpen = true;
        }

        private void UserListBox_MouseRightButtonDown(object sender, MouseButtonEventArgs e)
        {

        }

        private void UserListBox_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (UserListBox.SelectedItem != null)
            {
                UserTextBox.Text = (UserListBox.SelectedItem as ListBoxItem)?.Content.ToString();
                UserSelectionPopup.IsOpen = false;
            }
        }

        private void UserListBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!UserSelectionPopup.IsKeyboardFocusWithin)
            {
                UserSelectionPopup.IsOpen = false;
            }
        }
        

        private void ClearSearchAdvanced_Click(object sender, RoutedEventArgs e)
        {
            anafcombobox.SelectedIndex = 0;
            tkufa.SelectedIndex = 6;
            Nadontx.Text = "";
            FromShoteftx.Text = "";
            toShoteftx.Text = "";
        }

        private void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            int FromShotef = 0;
            int toShotef = 0;
            string fromDate = "";
            string toDate = "";
            string Nadon = "";
            string SholeahTeur = UserTextBox.Text;
            int SholeahKod = 0;
            string Mehutav = "";
            int IsLePeula = 0;
            int IsPail = 0;
            int Tik = 0;
            int IsHufatz = 0;
            int SugMismach = 0;
            int top = 1000;
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

            DateTime today = DateTime.Today;
            toDate= tkufa.SelectedIndex != 10 ? today.ToString("yyyyMMdd") : "";
            if (tkufa.SelectedIndex != 10)
            {
                if (tkufa.SelectedIndex == 0)
                    fromDate= today.ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 1)
                    fromDate = today.AddDays(-7).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 2)
                    fromDate = today.AddMonths(-1).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 3)
                    fromDate = today.AddMonths(-2).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 4)
                    fromDate = today.AddMonths(-3).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 5)
                    fromDate = today.AddMonths(-6).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 6)
                    fromDate = today.AddYears(-1).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 7)
                    fromDate = today.AddYears(-2).ToString("yyyyMMdd");
                else if (tkufa.SelectedIndex == 8)
                    fromDate = today.AddYears(-3).ToString("yyyyMMdd");
                else
                    fromDate = today.AddYears(-5).ToString("yyyyMMdd");
            }
            else
                fromDate = "";

            short res13 = -1, res14 = -1;
            /*if (short.TryParse(textBox14.Text, out res14))
            {
                if (!projects.ContainsKey(res14))
                {
                    MessageBox.Show("אין פרויקט עם מספר מזהה זה", "חיפוש מפרויקט שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjMe", res14));
            }
            else
            COMMAND.Parameters.Add(new SqlParameter("@P_ProjMe", "0"));
            if (short.TryParse(textBox13.Text, out res13))
            {
                if (!projects.ContainsKey(res13))
                {
                    MessageBox.Show("אין פרויקט עם מספר מזהה זה", "חיפוש עד פרויקט שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
                COMMAND.Parameters.Add(new SqlParameter("@P_ProjAd", res13));
            }
            else
            COMMAND.Parameters.Add(new SqlParameter("@P_ProjAd", "0"));*/

            /*if (res13 != -1 && res14 != -1 && res14 > res13)
            {
                Forms.MessageBox.Show("לא ניתן לחפש פרויקט סיום שקטן מפרויקט התחלה", "חיפוש פרויקט שגוי", Forms.MessageBoxButtons.OK, Forms.MessageBoxIcon.Information,
                        Forms.MessageBoxDefaultButton.Button1, Forms.MessageBoxOptions.RightAlign | Forms.MessageBoxOptions.RtlReading);
                return -1;
            }*/

            /*int rc;
            if (textBox7.Text != "0")
            {
                if (int.TryParse(textBox7.Text, out rc))
                {
                    if (PublicFuncsNvars.getUserNameByUserCode(rc) == null)
                    {
                        MessageBox.Show("אין שולח עם מספר משתמש זה.", "חיפוש שולח שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        return -1;
                    }
                }
                else if (textBox7.Text != "" && !textBox7.Text.Equals("קוד"))
                {
                    MessageBox.Show("מספר שולח לא יכול להכיל תווים שאינם ספרות.", "חיפוש שולח שגוי", MessageBoxButtons.OK, MessageBoxIcon.Information,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                    return -1;
                }
            }*/
            Nadon = Nadontx.Text;
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahTeur", !textBox8.Text.Equals("") && !textBox8.Text.Equals("שם / תפקיד") ? textBox8.Text : ""));
            //COMMAND.Parameters.Add(new SqlParameter("@P_SholeahKod", int.TryParse(textBox7.Text, out rc) ? rc : 0));
            //COMMAND.Parameters.Add(new SqlParameter("@P_Mehutav", (!textBox9.Text.Equals("") && !textBox9.Text.Equals("שם / תפקיד")) ? textBox9.Text : ""));
            //COMMAND.Parameters.Add(new SqlParameter("@P_IsLePeula", (!comboBox1.Text.Equals("הכל")) ? comboBox1.SelectedIndex - 1 : SqlInt32.Null));
            //IsLePeula = SqlInt32.Null;
            //COMMAND.Parameters.Add(new SqlParameter("@P_IsPail", checkBox1.Checked ? 1 : SqlInt32.Null));
            //COMMAND.Parameters.Add(new SqlParameter("@P_Tik", (!textBox4.Text.Equals("") && !textBox4.Text.Equals("שם תיק")) ? long.Parse(textBox12.Text) : 0));
            //COMMAND.Parameters.Add(new SqlParameter("@P_IsHufats", (!comboBox4.Text.Equals("הכל")) ? comboBox4.SelectedIndex - 1 : SqlInt32.Null));
            //COMMAND.Parameters.Add(new SqlParameter("@P_SugMismach", (!comboBox2.Text.Equals("הכל")) ? comboBox2.SelectedIndex - 1 : 0));
            top=(!Top.Text.Equals("הכל")) ? int.Parse(Top.Text) : 999999999;
            anaf=!anafcombobox.Text.Equals("הכל") ? PublicFuncsNvars.getBranchByString(anafcombobox.Text) : '\0';


            int numberOfRows = DataGridDocs.SearchDocs(FromShotef, toShotef, fromDate, toDate, Nadon, SholeahTeur, SholeahKod, Mehutav, IsLePeula, IsPail, Tik, IsHufatz, SugMismach, top, anaf);
            
        }

        private void UserTextBox_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            UserSelectionPopup.PlacementTarget = UserTextBox;
            UserSelectionPopup.Placement = System.Windows.Controls.Primitives.PlacementMode.Bottom;
            UserSelectionPopup.IsOpen = true;
            e.Handled = true;
        }

        private void FromShoteftx_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                int res1, res2;
                bool ok1 = int.TryParse(FromShoteftx.Text, out res1), ok2 = int.TryParse(toShoteftx.Text, out res2);


                if (ok1)
                {
                    toShoteftx.Text = FromShoteftx.Text;
                    if (FromShoteftx.Text == toShoteftx.Text)
                        tbSearch.Text = toShoteftx.Text;
                    else
                        tbSearch.Text = Nadontx.Text;
                    //comboBox8.SelectedIndex = 10;
                }
                //else
//comboBox8.SelectedIndex = 6;
            }

            catch { }
            
        }

        private void Nadontx_TextChanged(object sender, TextChangedEventArgs e)
        {
            int res1, res2;
            bool ok1 = int.TryParse(FromShoteftx.Text, out res1), ok2 = int.TryParse(toShoteftx.Text, out res2);
            if (ok1)
            {
                toShoteftx.Text = FromShoteftx.Text;
                if (FromShoteftx.Text != toShoteftx.Text)
                    tbSearch.Text = Nadontx.Text;
            }
            else
                tbSearch.Text = Nadontx.Text;
        }

        private void Grid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                SearchButton_Click(sender, e);
            }
        }
    }
}
