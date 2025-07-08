using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Sockets;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Data;

//using Microsoft.Office.Core;

namespace DocumentsModule
{
    public enum Branch { office = '1', archive = '2', tiger = '3', projects = '4', budgets = '5', computers = '6', organization = '7', manufacturing = '8', development = '9', sayaf = 'ס', chelka = 'ח', other = 'א' };
    //יערה שינתה ב-19.03.23 כדי שסיווגים סודי ומעלה יופיעו בסגולה
    public enum Classification { unclassified = 1, restricted = 2, confidetial = 3, secret = 4, topSecret = 5, secret_shos = 8, topSecret_shos = 9, sensitivePersonal =10, unknown = 0 };
    public enum RecipientListsLevel { personal = 1, branch = 2, unit = 3 };
    public enum RoleType { none = 0/*משתמש רגיל*/, clerk = 1/*פקידה*/, computers = 2/*צוות מחשוב*/, departmentHead = 3/*רמ"ד*/, branchHead = 4/*רע"ן*/ };
    public enum FileType { organization = 'א', spares = 'ח', steelPlates = 'ל', occasional = 'מ', subject = 'נ', serial = 'ס', kit = 'ע', development = 'פ', tzbm = 'צ', regular = 'ר', shosh = 'ש' };
    public enum DocType { normal, directive, steelsAllotment, lendingOrAllotment, bnm, developmentLending };

    class CompareRecipientLists : IEqualityComparer<KeyValuePair<short, short>>
    {
        public bool Equals(KeyValuePair<short, short> list1, KeyValuePair<short, short> list2)
        {
            return list1.Key == list2.Key && list1.Value == list2.Value;
        }

        public int GetHashCode(KeyValuePair<short, short> list)
        {
            return int.Parse(list.Key.ToString() + "0" + list.Value.ToString());
        }
    }

    public class PublicFuncsNvars
    {
        internal static List<int> openDocs = new List<int>();//this holds a list of the open ducuments so the user can't open the same document more than once
        internal static List<string> openVerDocs = new List<string>();
        internal static List<int> dhFormsOpen = new List<int>();//this holds a list of the open DocumentHandling forms so the user can"t open the same one more than once
        internal static List<int> docsHeldInDB = new List<int>();//this holds a list of the documents held by the current user in the database
        internal static List<KeyValuePair<int, int>> openAtts = new List<KeyValuePair<int, int>>();//same as the documents one, for attachments
        internal static List<User> users;//list of users
        //30.10.22 יערה הורידה על מנת להשתמש בכל התוכנה במשתנה המוגדר בהגדרות
        //internal static string conStr = "Data Source=modsql6p;Initial Catalog=MantakDB;Integrated Security=True";//connection to our database
        //internal static string conStr = Properties.Settings.Default.MantakDBPConnectionString;
        //internal static string conStr = "Data Source=" + Global.P_SQL_SRV+ ";Initial Catalog="+Global.P_SQL_DB+";Integrated Security=True";
        internal static List<Folder> folders;//all folders
        internal static User curUser;//the current user using the system
        internal static Dictionary<int, string> projects;
        internal static Dictionary<KeyValuePair<short, short>, RecipientList> recipientsLists;//existing recipient lists to use as a group
        internal static List<Recipient> interDist = new List<Recipient>();//ת.פ.
        internal static string userLogin = Environment.UserName.ToUpper();//the current user's username
        private static bool iSave = false;
        // ASAF MOR
        private static bool afterSave = true;

        private static string newName = null;
        internal static bool isRegularDocument = false;

        /*an attempt to speed sql queries up*/
        internal static void setArithabortOn()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SET ARITHABORT ON", conn);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
        }


        public static string getCurUser()
        {
            return curUser.userLogin + " - " + curUser.getFullName();
        }
        internal static void getInterDist()
        {
            // 5.11.2023 יערה - ביטול סינון לי תפקידים ברשימת תפ
            foreach (User u in users.Where(x => x.isActive))
                interDist.Add(new Recipient(u.userCode, -1, u.job, false, true, u.email));
        }

        private static Recipient getInterDistMByIdFromDB(int id)
        {
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT doal, taor_tpkid FROM dbo.tmtafkidu WHERE kod_tpkid=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            Recipient r = new Recipient(id, -1, sdr.GetString(1).Trim(), false, true, sdr.GetString(0).Trim());
            conn.Close();
            return r;
        }

        internal static User getUserFromLogIn(string user)
        {
            foreach (User u in users)
            {
                if (u.userLogin.ToUpper().Equals(user.ToUpper()))//ToLower() == user.ToLower());//.
                    return u;
            }

            MessageBox.Show($@"המשתמש {user} לא קיים במערכת!", "לא קיים", MessageBoxButtons.OK, MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            return null;
        }

        internal static string getBranchString(Branch branch)
        {
            switch (branch)
            {
                case Branch.office:
                    return "לשכה";
                case Branch.projects:
                    return "פרוייקטים";
                case Branch.budgets:
                    return "תקציבים";
                case Branch.organization:
                    return "ארגון";
                case Branch.manufacturing:
                    return "ייצור";
                case Branch.development:
                    return "פיתוח";
                case Branch.sayaf:
                    return "סייף";
                case Branch.chelka:
                    return "חלקה";
                case Branch.other:
                    return "אחר";
                case Branch.archive:
                    return "ארכיון";
                case Branch.computers:
                    return "מחשוב";
                case Branch.tiger:
                    return "נמר חו\"ל";
            }
            return null;
        }

        /*
         * this function gets 3 textBoxes and puts the right info in the name and job textboxes
         * according to the user code in the code textbox.
         */
        public static void nameNjobByCode(ref TextBox codeTB, ref TextBox firstNameTB, ref TextBox lastNameTB, ref TextBox jobTB)//לא בשימוש
        {
            if (codeTB.Text.Equals(""))
            {
                jobTB.Text = "";
                firstNameTB.Text = "";
                lastNameTB.Text = "";
            }
            else if (codeTB.Text.Equals("קוד"))
            {
                jobTB.Text = "תפקיד";
                firstNameTB.Text = "שם פרטי";
                lastNameTB.Text = "שם משפחה";
            }
            else
            {
                int res;
                if (int.TryParse(codeTB.Text, out res))
                    foreach (User u in users)
                    {
                        if (u.userCode == res)
                        {
                            jobTB.Text = u.job;
                            firstNameTB.Text = u.firstName;
                            lastNameTB.Text = u.lastName;
                        }
                    }
            }
        }

        /*
         * this function uses one textbox that is used for folder info and puts the right info in the others according to it
         */
        public static void directoryByCode(ref TextBox codeTB, ref TextBox shortTB, ref TextBox nameTB, ref TextBox numTB, string check,
            string replace, string command, string paramName, Type t)
        {
            if (codeTB.Text.Equals(""))
            {
                shortTB.Text = "";
                nameTB.Text = "";
                if (numTB != null)
                    numTB.Text = "";
            }
            else if (codeTB.Text.Equals(check))
            {
                shortTB.Text = replace;
                nameTB.Text = "שם תיק";
                if (numTB != null)
                    numTB.Text = "מס' בתיק";
            }
            else
            {
                int res = -1;
                if ((t.Equals(typeof(int)) && int.TryParse(codeTB.Text, out res)) || t.Equals(typeof(string)))
                {
                    SqlConnection conn = new SqlConnection((Global.ConStr));
                    SqlCommand comm = new SqlCommand(command, conn);
                    if (res != -1)
                        comm.Parameters.AddWithValue(paramName, (long)res);
                    else
                        comm.Parameters.AddWithValue(paramName, codeTB.Text);
                    conn.Open();
                    SqlDataReader sdr = comm.ExecuteReader();
                    if (sdr.Read())
                    {
                        if (res != -1)
                            shortTB.Text = sdr.GetString(0).Trim()+" - "+sdr.GetString(1).Trim();
                        else
                            shortTB.Text = sdr.GetString(0).Trim() + " - " + sdr.GetInt32(1).ToString();
                        nameTB.Text = sdr.GetString(0).Trim();
                        conn.Close();
                        conn.Open();
                        if (numTB != null)
                        {
                            comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(mispar_in_tik) IS NULL THEN 0" + Environment.NewLine +
                                "ELSE MAX(mispar_in_tik)" + Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.tiukim WHERE mispar_nose=@id", conn);
                            if (res == -1)
                                comm.Parameters.AddWithValue("@id", int.Parse(shortTB.Text));
                            else
                                comm.Parameters.AddWithValue("@id", res);
                            numTB.Text = (int.Parse(comm.ExecuteScalar().ToString()) + 1).ToString();
                        }
                    }
                    else
                    {
                        shortTB.Text = "";
                        nameTB.Text = "";
                        if (numTB != null)
                            numTB.Text = "";
                    }
                    conn.Close();
                }
            }
        }

        internal static void getUsers()
        {
            users = new List<User>();
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT kod_tpkid, taor_tpkid, anp_hbrh_ihidh, shm_mshtmsh_MAGIC, User_Login, anp_mshtmsh_bmntk, doal, shm_prti, shm_mshphh, commanderId, roleTypeCode, isActive, isAllowedToOpenFolders FROM dbo.tmtafkidu WHERE shm_mshtmsh_MAGIC<>'' --ORDER BY kod_tpkid", conn);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                try
                {
                    PublicFuncsNvars.users.Add(new User(sdr.GetInt32(0), int.Parse(sdr.GetString(3).Trim()), sdr.GetString(7).Trim(), sdr.GetString(8).Trim(),
                        sdr.GetString(6).Trim(), sdr.GetString(4).Trim(), sdr.GetString(1).Trim(), sdr.GetSqlChars(5).Buffer[0], sdr.GetInt16(2), sdr.GetInt32(9),
                        sdr.GetInt16(10), sdr.GetBoolean(11), sdr.GetBoolean(12)));
                }
                catch (Exception e)
                {
                    saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                }
            }
            conn.Close();
        }

        internal static void getFolders()
        {
            folders = new List<Folder>();
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT ms_mshimh, shm_mshimh, shm_mkotzr, anp, is_tik_pail, sog_mshimh, ms_archh_shosh, kod_sioog FROM dbo.tm_mesimot WHERE shm_mkotzr<>'' and is_tik_pail=1 --ORDER BY ms_mshimh", conn);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                if (sdr.GetSqlDouble(6).Value != 0)
                    PublicFuncsNvars.folders.Add(new ShoshFolder(sdr.GetInt32(0), sdr.GetString(1).Trim(), sdr.GetString(2).Trim(), true,
                        (Branch)(sdr.GetSqlChars(3).Value[0]), sdr.GetBoolean(4), (FileType)(sdr.GetSqlChars(5).Value[0]), getClassification(sdr.GetInt16(7)),
                        (int)sdr.GetSqlDouble(6).Value));

                else
                    PublicFuncsNvars.folders.Add(new Folder(sdr.GetInt32(0), sdr.GetString(1).Trim(), sdr.GetString(2).Trim(), true,
                        (Branch)(sdr.GetSqlChars(3).Value[0]), sdr.GetBoolean(4), (FileType)(sdr.GetSqlChars(5).Value[0]), getClassification(sdr.GetInt16(7))));

            }
            conn.Close();
        }

        internal static void getProjects()
        {
            projects = new Dictionary<int, string>();
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT nam_proyykt, num_proyykt From dbo.tm_project where nam_proyykt<>'' --ORDER BY num_proyykt", conn);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
                PublicFuncsNvars.projects.Add(sdr.GetInt16(1), sdr.GetString(0).Trim());
            conn.Close();
        }

        internal static void getRecipientsLists()
        {
            recipientsLists = new Dictionary<KeyValuePair<short, short>, RecipientList>(new CompareRecipientLists());
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT cod_sys, cod_lst_tpotzh, nam_lst_tpotzh, anf_tpotzh, creating_user, personalObranchOunit FROM dbo.tm_tfuza --ORDER BY cod_lst_tpotzh", conn);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                short systemCode = (short)sdr.GetByte(0);
                short id = sdr.GetInt16(1);
                recipientsLists.Add(new KeyValuePair<short, short>(systemCode, id), new RecipientList(id, sdr.GetInt32(4), sdr.GetInt16(5), sdr.GetSqlChars(3).Value[0], sdr.GetString(2).Trim(),
                    getRecipientsByList(id, systemCode)));
            }
            conn.Close();
        }

        private static List<Recipient> getRecipientsByList(short id, short sysCode)
        {
            List<Recipient> recipients = new List<Recipient>();
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT onum, cod_mcotb, is_actn_ydyah, tiur_tafkid, ktovet_mail FROM dbo.tm_tfuz_res WHERE cod_lst_tpotzh=@listCode AND cod_sys=@sysCode", conn);
            comm.Parameters.AddWithValue("@listCode", id);
            comm.Parameters.AddWithValue("@sysCode", (byte)sysCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                string role;
                string email;
                User u;
                if (null != (u = getUserByCode(sdr.GetInt32(1))))
                {
                    role = u.job;
                    email = u.email;
                }
                else
                {
                    role = sdr.GetString(3).Trim();
                    email = sdr.GetString(4).Trim();
                }
                recipients.Add(new Recipient(sdr.GetInt32(1), sdr.GetInt16(0), role, !sdr.GetBoolean(2), true, email));
            }
            conn.Close();
            return recipients;
        }

        internal static string removeNansButLetters(string p)
        {
            string newString = "";
            foreach (char c in p)
            {
                if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z') || (c >= 'א' && c <= 'ת'))
                    newString += c;
            }
            return newString;
        }

        internal static Classification getClassification(short c)
        {
            switch (c)
            {
                case 1:
                    return Classification.unclassified;
                case 2:
                    return Classification.restricted;
                case 3:
                    return Classification.confidetial;
                case 4:
                case 7:
                    return Classification.secret;
                case 5:
                case 6:
                    return Classification.topSecret;
                //יערה שינתה ב-19.03.23 כדי שסיווגים סודי ומעלה יופיעו  בסגולה
                case 8:
                    return Classification.secret_shos;
                case 9:
                    return Classification.topSecret_shos;
                case 10:
                    return Classification.sensitivePersonal;
                default:
                    return Classification.unknown;
            }
        }

        internal static short getClassificationCode(string c)
        {
            switch (c)
            {
                case "בלמ\"ס":
                    return 1;
                case "מוגבל":
                    return 2; //Obsolete
                case "שמור":
                    return 3;
                case "סודי":
                    return 4;
                case "סודי ביותר":
                    return 5;
                case "סודי לשו\"ס ":
                    return 8;
                case "סודי ביותר לשו\"ס":
                    return 9;
                default:
                    return 0;
            }
        }

         


        internal static void initializeCurrentUser()
        {
            curUser = getUserFromLogIn(userLogin);
        }

        internal static List<Folder> getDirectoriesForDoc(int id)
        {
            List<int> ids = new List<int>();
            List<Folder> toReturn = new List<Folder>();
            SqlConnection conn = new SqlConnection((Global.ConStr));
            SqlCommand comm = new SqlCommand("SELECT mispar_nose FROM dbo.tiukim WHERE shotef_klali=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while (sdr.Read())
            {
                ids.Add(sdr.GetInt32(0));
            }
            foreach (Folder d in folders)
            {
                if (ids.Contains(d.id))
                    toReturn.Add(d);
            }
            conn.Close();
            return toReturn;
        }

        internal static string getEmailByUserCode(int code)
        {
            foreach (User u in users)
            {
                if (u.userCode == code)
                    return u.email;
            }
            return null;
        }

        /*
         * opening the document specified for viewing without editing
         */
        internal void viewDoc(int id)
        {
            if (MyGlobals.afterDelete == true)
            {
                Application.UseWaitCursor = true;
                Cursor.Current = Cursors.WaitCursor;
                Thread.Sleep(7000);
                MyGlobals.afterDelete = false;
            }

            MyGlobals.afterViewOnly = true;
            Cursor.Current = Cursors.WaitCursor;
            if (!openDocs.Contains(id))
            {
                SqlConnection conn = new SqlConnection((Global.ConStr));
                SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                comm.Parameters.AddWithValue("@id", id);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    openDocs.Add(id);
                    byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                    string fileExt = sdr.GetString(1).Trim();
                    string filePath = Program.folderPath + "\\" + id + "." + fileExt;


                    if (File.Exists(filePath))
                    {
                        try
                        {
                            string archiveFolder = Program.archiveFolder;// Path.GetDirectoryName(filePath) + "/Archive/";
                            Directory.CreateDirectory(archiveFolder);
                            string copyTo = archiveFolder + "_" + Path.GetFileNameWithoutExtension(filePath) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(filePath);
                            File.Move(filePath, copyTo);
                        }

                        catch { }
                    }

                    if (!File.Exists(filePath))
                    {
                        File.WriteAllBytes(filePath, fileData);


                        if (fileExt.ToLower() == "doc" || fileExt.ToLower() == "docx" || fileExt.ToLower().Contains("doc"))
                        {
                            /*ProcessStartInfo startinfo = new ProcessStartInfo
                            {
                                FileName = @"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
                                Arguments = $"/r \"{filePath}\"",
                                UseShellExecute = true
                            };
                            Process.Start(startinfo);*/
                            docViewOnly(filePath);

                            // BringToFront(id);
                        }
                        else
                            Process.Start(filePath);
                        while (true)
                        {
                            Thread.Sleep(5000);
                            try
                            {
                                File.Delete(filePath);
                                openDocs.Remove(id);
                                break;
                            }
                            catch (Exception e) { }
                        }
                    }
                    else
                    {
                        openDocs.Remove(id);
                    }
                }
                else
                {
                    int attId = GetFirstAtt(id);
                    if (attId > -1)
                        viewAtt(id, attId);

                    else
                    {
                        MessageBox.Show("מסמך זה לא קיים במערכת", "לא קיים", MessageBoxButtons.OK, MessageBoxIcon.Error,
                      MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                }

                   
                conn.Close();
            }
            else
            {
                MessageBox.Show("מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסמך מספר פעמים", "מסמך פתוח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);


            }
            Cursor.Current = Cursors.Default;
            Application.UseWaitCursor = false;
        }

        private int GetFirstAtt(int idObj)
        {
            int attId = -1;
            using (SqlConnection conn = new SqlConnection(Global.ConStr))
            {
                conn.Open();
                SqlCommand COMMAND = new SqlCommand(@"select * from dbo.F_GetDocFiles(@P_ShotefMismach)", conn);
                COMMAND.CommandType = CommandType.Text;

                COMMAND.Parameters.Add(new SqlParameter("@P_ShotefMismach", idObj));

                using (SqlDataReader reader = COMMAND.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        attId = reader.GetInt32(1);
                        if (attId > 0) break;
                    }
                }
            }
            if (attId == 0) attId = -1;
            return attId;
        }

        private void BringToFront(int id)
        {
            foreach (Process clsProcess in Process.GetProcesses())
                if (clsProcess.MainWindowTitle.Contains(id.ToString()))
                {
                    try
                    {
                        // MessageBox.Show(clsProcess.ProcessName);
                        var hndl = clsProcess.Handle;
                       bool r = SetForegroundWindow(hndl);
                        if (!r)
                        { }
                    }
                    catch (Exception ex)
                    { }
                }
        }

        internal static void docViewOnly(string filePath)
        {
            Word.Application wapp;
            try
            {
                wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
            }
            catch
            {
                wapp = new Word.Application();
            }
            wapp.Visible = true;
            object missing = Type.Missing;
            try
            {
                Word.Document doc = wapp.Documents.Open(filePath, missing, true);
                doc.Activate();
                wapp.Activate();

                ////wapp.Activate();
                //Word.Window window = wapp.ActiveWindow;
                //window.SetFocus();
                //window.Activate();
                //if (window != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(window);
            }
            catch
            {

            }
        }

        /*
         * opening the attachment specified for viewing without editing
         */
        internal static void viewAtt(int docId, int attId)
        {
            Cursor.Current = Cursors.WaitCursor;
            KeyValuePair<int, int> attToOpen = new KeyValuePair<int, int>(docId, attId);
            if (!openAtts.Contains(attToOpen))
            {
                SqlConnection conn = new SqlConnection((Global.ConStr));
                SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.docnisp WHERE shotef_mchtv=@docId AND shotef_nisph=@id AND datalength(file_data)>0", conn);
                comm.Parameters.AddWithValue("@docId", docId);
                comm.Parameters.AddWithValue("@id", attId);
                conn.Open();
                SqlDataReader sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    openAtts.Add(attToOpen);
                    byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                    string fileExt = sdr.GetString(1).Trim().Replace(".","");
                    string filePath = Program.folderPath + "\\" + docId + "_" + attId + "." + fileExt;
                    File.WriteAllBytes(filePath, fileData);
                    if (fileExt.ToLower() == "doc" || fileExt.ToLower() == "docx" || fileExt.ToLower().Contains("doc"))
                    {
                        docViewOnly(filePath);
                    }
                    else
                        Process.Start(filePath);
                    while (true)
                    {
                        Thread.Sleep(5000);
                        try
                        {
                            File.Delete(filePath);
                            openAtts.Remove(attToOpen);
                            break;
                        }
                        catch (Exception e) { }
                    }
                }
                else
                {
                    MessageBox.Show("נספח זה לא קיים במערכת. בדקו הבעיה עם צוות מחשוב", "לא קיים", MessageBoxButtons.OK, MessageBoxIcon.Error,
                        MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסמך מספר פעמים", "מסמך פתוח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

            }
            Cursor.Current = Cursors.Default;
        }

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        /*
         * opening the document specified for viewing with editing
         */
        internal static void viewDocForEdit(int id)
        {
            Cursor.Current = Cursors.WaitCursor;
            if (MyGlobals.afterViewOnly == true)
            {
                Thread.Sleep(7000);
                MyGlobals.afterViewOnly = false;
            }
            SqlConnection conn = new SqlConnection((Global.ConStr));
            if (openDocs.Contains(id))
            {
                Thread.Sleep(4000);
            }
            if (!openDocs.Contains(id))
            {
                int holdingUser = whoHoldsThisDoc(id);
                int holdingHours;
                DateTime? holdingSince = null;
                if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode) sinceWhenDocHold(id);
                if (holdingSince == null) holdingHours = 0;
                else holdingHours = Convert.ToInt32((DateTime.Now - sinceWhenDocHold(id)).Value.TotalHours);
                if (holdingUser == 0 || holdingUser == PublicFuncsNvars.curUser.userCode)
                {
                    updateDB(id, true);
                    SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                    comm.Parameters.AddWithValue("@id", id);
                    conn.Open();
                    SqlDataReader sdr = comm.ExecuteReader();
                    if (sdr.Read())
                    {
                        openDocs.Add(id);
                        byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                        string fileExt = sdr.GetString(1).Trim();
                        string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                        newName = filePath;

                        string lowerCaseFileExt = fileExt.ToLower();
                        if (lowerCaseFileExt == "docx" || lowerCaseFileExt == "doc")
                        {
                            //openViewDoc(id);
                            OpenDocForEditAndNotClose(id);
                            openDocs.Remove(id);
                            releaseHeldDoc(id);
                            docsHeldInDB.Remove(id);
                        }
                            
                        else
                        {
                            if (File.Exists(filePath))
                            {
                                try
                                {
                                    string archiveFolder = Program.archiveFolder;
                                    Directory.CreateDirectory(archiveFolder);
                                    string copyTo = archiveFolder + "_" + Path.GetFileNameWithoutExtension(filePath) + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + Path.GetExtension(filePath);
                                    File.Move(filePath, copyTo);
                                }

                                catch (Exception e)
                                {
                                    MessageBox.Show("פתיחה נכשלה: " + e.Message);
                                }

                                try{}
                                catch (Exception)
                                {
                                    throw;
                                }
                            }
                            if (!File.Exists(filePath))
                            {
                                dynamic document = new object();
                                File.WriteAllBytes(filePath, fileData);
                                lowerCaseFileExt = fileExt.ToLower();
                                if (lowerCaseFileExt == "docx" || lowerCaseFileExt == "doc")
                                    lowerCaseFileExt = lowerCaseFileExt;
                                
                                else if (lowerCaseFileExt == "xlsx" || lowerCaseFileExt == "xls")
                                {
                                    Excel.Application eApp = new Excel.Application();
                                    document = eApp.Workbooks.Open(filePath);
                                    try { document.unprotected("pikachu"); }
                                    catch { };
                                    eApp.Visible = true;
                                }
                                else
                                    Process.Start(filePath);
                                int timeToUpdate = 300;
                                while (true)
                                {
                                    Thread.Sleep(5000);
                                    timeToUpdate -= 5;
                                    try
                                    {
                                        try
                                        {
                                            if (timeToUpdate == 0 && (lowerCaseFileExt == "docx" || lowerCaseFileExt == "doc" || lowerCaseFileExt == "xlsx" || lowerCaseFileExt == "xls"))
                                            {
                                                iSave = true;
                                                Word.WdProtectionType pt = ((Word.Document)document).ProtectionType;
                                                document.Save();
                                                timeToUpdate = 300;
                                                if (pt == Word.WdProtectionType.wdAllowOnlyReading && ((Word.Document)document).ProtectionType != pt)
                                                    //A.M   document.Protect(Word.WdProtectionType.wdAllowOnlyReading, Type.Missing, "pikachu");
                                                    iSave = iSave;
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            saveLogError("publicFuncsNvars", ex.ToString(), ex.Message);
                                        }
                                        if (timeToUpdate % 60 == 0 || !isFileOpen(newName))
                                        {
                                            while (!afterSave) ;
                                            FileStream fs = File.Open(newName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                                            BinaryReader reader = new BinaryReader(fs);
                                            fileData = reader.ReadBytes((int)fs.Length);
                                            reader.Close();
                                            fs.Close();
                                            //(גרסאות)
                                            DocumentHandling.SaveVersion(id);
                                            saveDocToDB(ref fileData, id, filePath, ref comm, ref conn,"");
                                            
                                            SqlCommand comm2 = new SqlCommand("INSERT INTO dbo.documentsModuleSavingLog(dateNtime, userCode, fileLocation)"
                                                + " VALUES(@dateNtime, @userCode, @fileLocation)", conn);
                                            comm2.Parameters.AddWithValue("@dateNtime", DateTime.Now);
                                            comm2.Parameters.AddWithValue("@userCode", curUser.userCode);
                                            comm2.Parameters.AddWithValue("@fileLocation", newName);
                                            conn.Open();
                                            comm2.ExecuteNonQuery();
                                            conn.Close();
                                        }
                                        try
                                        {
                                            FileStream stream = File.Open(newName, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                                            stream.Close();
                                            while (!afterSave) ;
                                            MyGlobals.afterDelete = true;
                                            File.Delete(newName);
                                            if (File.Exists(filePath))
                                                File.Delete(filePath);
                                            openDocs.Remove(id);
                                            releaseHeldDoc(id);
                                            docsHeldInDB.Remove(id);
                                            break;
                                        }
                                        catch (Exception ex) { }
                                    }
                                    catch (SqlException e)
                                    {
                                        if (document is Word.Document)
                                            e = e; //  document.Protect(Word.WdProtectionType.wdAllowOnlyReading, Type.Missing, "pikachu");
                                        else if (document is Excel.Workbook)
                                           e = e; //    document.Protect("pikachu");
                                        MessageBox.Show("המסמך שלך לא הצליח להישמר בניסיון השמירה האחרון, ליתר ביטחון מומלץ לשמור עותק ברשת.", "שמירת מסמך נכשלה",
                                            MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1,
                                             MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                        try { document.Unprotect("pikachu"); }
                                        catch { }
                                        saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                                    }
                                }
                            }
                            else
                            {
                                openDocs.Remove(id);
                                releaseHeldDoc(id);
                                docsHeldInDB.Remove(id);
                            }
                        }
                        
                        
                    }
                    else
                    {
                        MessageBox.Show("מסמך זה לא קיים במערכת, נסו לפתוח את הנספחים שלו", "לא קיים", MessageBoxButtons.OK, MessageBoxIcon.Error,
                            MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }
                    conn.Close();
                }
                else // A.M  Document is already open on somone else computer
                {


                    // if opened for more then 24h -> change the ownership , else -> do nothing. 
                    if (holdingHours > 24)
                    {
                        DialogResult res = MessageBox.Show("מסמך זה בבעלות " + users.First(x => x.userCode == holdingUser).getFullName() + "  " + " \nהאם ברצונך להעביר את הבעלות על המסמך אליך?", "מסמך פתוח", MessageBoxButtons.YesNo,
                                         MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);

                        if (res == DialogResult.Yes)
                        {
                            if (releaseDoc(id) == true)
                            {
                                viewDocForEdit(id);
                            }
                        }
                    }

                    else
                    {
                        MessageBox.Show("מסמך זה כבר פתוח אצל " + users.First(x => x.userCode == holdingUser).getFullName() + " ", "מסמך פתוח", MessageBoxButtons.OK,
                        MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                    }

                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("מסמך זה כבר פתוח אצלך, לא ניתן לפתוח את אותו מסמך מספר פעמים", "מסמך פתוח", MessageBoxButtons.OK, MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                MyGlobals.afterEdit = true;
                return;
            }
            Cursor.Current = Cursors.Default;
        }

        private static void wApp_DocumentBeforeClose(Word.Document Doc, ref bool Cancel)
        {
            iSave = true;
        }
        
        private static void wApp_DocumentAfterSave(Word.Document Doc, bool isCanceled)
        {
            newName = Doc.FullName;
            afterSave = true;
        }
        

        private static bool isFileOpen(string filePath)
        {
            try
            {
                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
                stream.Close();
                return false;
            }
            catch (Exception ex)
            {
                return true;
            }
        }

        internal static string getUserEmail(int user)
        {
            foreach (User u in users)
            {
                if (user == u.userCode || (u.userCode == 100 && user == 110))
                {
                    return u.email;
                }
            }
            return null;
        }

        internal static string getClassificationByEnum(Classification classification)
        {
            switch (classification)
            {
                case Classification.unclassified:
                    return "בלמ\"ס";
                case Classification.restricted:
                    return "מוגבל"; //Obsolete
                case Classification.secret:
                    return "סודי";
                case Classification.topSecret:
                    return "סודי ביותר";
                //יערה שינתה ב-19.03.23 כדי שסיווגים סודי ומעלה יופיעו  בסגולה
                case Classification.secret_shos:
                    return "סודי לשו\"ס";
                case Classification.topSecret_shos:
                    return "סודי ביותר לשו\"ס";
                case Classification.confidetial:
                default:
                    return "שמור";
            }
        }

        internal static void changeControlsVisiblity(bool b, List<Control> controls)
        {
            foreach (Control c in controls)
                while (true)
                {
                    try
                    {
                        c.Visible = b;
                        break;
                    }
                    catch { }
                }
        }

        internal static void makeControlsEnDisabled(bool b, Control[] controls)
        {
            foreach (Control c in controls)
                while (true)
                {
                    try
                    {
                        c.Enabled = b;
                        break;
                    }
                    catch { }
                }
        }

        internal static string getUserNameByUserCode(int userCode)
        {
            foreach (User u in users)
                if (u.userCode == userCode || (u.userCode == 100 && userCode == 110))
                    return u.getFullName();
            return null;
        }
        internal static int getUserCodeByUserTafkid(string userT)
        {
            foreach (User u in users)
                if (u.job==userT.Trim()||u.getFullName().Contains(userT.Trim())|| userT.Trim().Contains(u.getFullName()))
                    return u.userCode;
            return 0;//return null that is int
        }
        internal static User getUserByCode(int code)
        {
            foreach (User u in users)
                if (u.userCode == code || (u.userCode == 100 && code == 110))
                    return u;
            return null;
        }

        internal static KeyValuePair<int, string> getProjectById(int id)
        {
            foreach (KeyValuePair<int, string> p in projects)
            {
                if (p.Key == id)
                    return p;
            }
            return new KeyValuePair<int, string>(-1, "");
        }

        internal static void printExcel(string filePath)
        {
            Excel.Application eApp = new Excel.Application();
            Excel.Workbook wb = eApp.Workbooks.Open(filePath);
            foreach (Excel.Worksheet ws in wb.Worksheets)
                ws.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            wb.PrintOutEx();
            Thread.Sleep(5000);
            wb.Close(false);
            eApp.Quit();
        }

        internal static void printWord(string filePath)
        {
            Word.Application wApp = new Word.Application();
            Word.Document doc = wApp.Documents.Open(filePath);
            doc.PrintOut();
            while (wApp.BackgroundPrintingStatus > 0) ;
            doc.Close(false);
            wApp.Quit(false);
        }

        internal static void print(string filePath, string fileExt, byte[] fileData)
        {
            if (!File.Exists(filePath))
            {
                File.WriteAllBytes(filePath, fileData);
            }
            if (fileExt.Equals("xlsx") || fileExt.Equals("xls"))
            {
                printExcel(filePath);
            }
            else if (fileExt.Equals("docx") || fileExt.Equals("doc"))
            {
                printWord(filePath);
            }
            else
            {
                ProcessStartInfo toPrintStartInfo = new ProcessStartInfo(filePath);
                toPrintStartInfo.Verb = "Print";
                toPrintStartInfo.CreateNoWindow = true;
                toPrintStartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                Process.Start(toPrintStartInfo);
            }
        }

        internal static char getBranchByString(string branch)
        {
            switch (branch)
            {
                case "לשכה":
                    return '1';
                case "ארכיון":
                    return '2';
                case "נמר חו\"ל":
                    return '3';
                case "פרוייקטים":
                    return '4';
                case "תקציבים":
                    return '5';
                case "מחשוב":
                    return '6';
                case "ארגון":
                    return '7';
                case "ייצור":
                    return '8';
                case "פיתוח":
                    return '9';
                case "סייף":
                    return 'ס';
                case "חלקה":
                    return 'ח';
                default:
                    return 'א';
            }
        }

        internal static bool updateRecipientsInWordDoc(Document curDoc, bool updateLocation)
        {
            bool toReturn = true;
            if (!updateLocation)
            {


                int holdingUser = whoHoldsThisDoc(curDoc.getID());
                if (holdingUser != 0 && holdingUser != PublicFuncsNvars.curUser.userCode)
                {
                    MessageBox.Show("מסמך זה כבר פתוח לעריכה אצל " + PublicFuncsNvars.getUserNameByUserCode(holdingUser) + ", לא ניתן לעדכן מכותבים.",
                                        "מכותבים", MessageBoxButtons.OK, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                }
                else
                {
                    updateDB(curDoc.getID(), false);
                    SqlConnection conn = new SqlConnection(Global.ConStr);
                    string forAct = "", forKnow = "";
                    SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                    comm.Parameters.AddWithValue("@id", curDoc.getID());
                    conn.Open();
                    SqlDataReader sdr = comm.ExecuteReader();
                    if (sdr.Read())
                    {
                        byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                        string fileExt = sdr.GetString(1).Trim();
                        string filePath = Program.folderPath + "\\" +
                            curDoc.getID() + "." + fileExt;
                        int lengt = curDoc.getRecipients().Count;
                        foreach (Recipient r in curDoc.getRecipients())
                        {
                            string toAdd = r.getRole() + Environment.NewLine;
                            if (r.getIFA())
                                forAct += toAdd;
                            else
                                forKnow += toAdd;
                        }


                        if (forAct.EndsWith(Environment.NewLine))
                            forAct = forAct.Remove(forAct.Length - 2);
                        if (forKnow.EndsWith(Environment.NewLine))
                            forKnow = forKnow.Remove(forKnow.Length - 2);
                        bool itscreated = false;
                        if (!File.Exists(filePath))
                        {
                            File.WriteAllBytes(filePath, fileData);
                            itscreated = true;
                        }
                        else
                        {
                            try
                            {
                                fileData = File.ReadAllBytes(filePath);
                                //saveDocToDB(ref fileData, curDoc.getID(), filePath, ref comm, ref conn);
                                itscreated = true;
                            }
                            catch
                            {
                                MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את המכותבים.");
                            }
                        }

                        try
                        {
                            if (itscreated)
                            {
                                object missing = Type.Missing;
                                Word.Application wapp;
                                try
                                {
                                    wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                                }
                                catch
                                {
                                    wapp = new Word.Application();
                                }
                                bool iswAppVisible = wapp.Visible;
                                if (iswAppVisible)
                                    wapp.Visible = false;
                                Word.Document doc = wapp.Documents.Open(filePath);

                                string documentText = doc.Content.Text;
                                bool textExists = documentText.Contains("רשימת תפוצה");
                                dynamic customProperties = doc.CustomDocumentProperties;
                                Word.Range entireRange = doc.Content;

                                int partsAct = CustomParts(forAct, "נמענים_לפעולה", customProperties, doc,false);
                                int partsKnow = CustomParts(forKnow, "נמענים_לידיעה", customProperties, doc,false);
                                if (forAct != "")
                                    updateCustomPropertiesInWordDoc(doc, forAct, "נמענים_לפעולה",false);
                                else
                                    updateCustomPropertiesInWordDoc(doc, " ", "נמענים_לפעולה",false);

                                if (forKnow == "")
                                    //doc.BuiltInDocumentProperties["Title"].Value = forKnow;
                                    updateCustomPropertiesInWordDoc(doc, " ", "נמענים_לידיעה",false);
                                else
                                {
                                    //doc.BuiltInDocumentProperties["Title"].Value = forKnow;
                                    updateCustomPropertiesInWordDoc(doc, forKnow, "נמענים_לידיעה",false);
                                    /*if (partsKnow > 1)
                                    {

                                        entireRange = doc.Content;
                                        Word.Range foundRange = FindTextInRange(entireRange, customProperties["נמענים_לידיעה"].Value);
                                        if (foundRange != null)
                                        {
                                            Word.Range insertrange = doc.Range(foundRange.End, foundRange.End);
                                            insertrange.InsertAfter(customProperties["נמענים_לידיעה_1"].Value);
                                        }

                                        foreach (dynamic property in customProperties)
                                        {
                                            if (property.Name == "נמענים_לידיעה")
                                            {
                                                Word.Range proprange = doc.Range(property.LinkToContent ? property.LinkSource : property.LinkSource.Parent);
                                                if (proprange != null)
                                                {
                                                    Word.Range insertrange = doc.Range(proprange.End, proprange.End);
                                                    insertrange.InsertAfter(customProperties["נמענים_לידיעה_1"].Value);
                                                }
                                            }
                                        }
                                    }*/
                                }


                                if (lengt > 8 && !textExists)
                                {

                                    bool IsDocNewVersion = false;
                                    try
                                    {
                                        dynamic existingProperty = customProperties["סימוכין"];
                                        IsDocNewVersion = true;
                                    }
                                    catch
                                    { }
                                    if (IsDocNewVersion)
                                    {
                                        object unit = Word.WdUnits.wdStory;
                                        object extend = Word.WdMovementType.wdMove;
                                        wapp.Selection.EndKey(ref unit, ref extend);
                                        object breakType = Word.WdBreakType.wdPageBreak;
                                        doc.Application.Selection.InsertBreak(ref breakType);


                                        Word.Range rng = doc.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast, missing, missing);
                                        rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                        rng.Text = Environment.NewLine + "רשימת תפוצה" + Environment.NewLine + Environment.NewLine;
                                        rng.Underline = Word.WdUnderline.wdUnderlineSingle;
                                        rng.BoldBi = -1;

                                        rng = doc.GoTo(Word.WdGoToItem.wdGoToLine, Word.WdGoToDirection.wdGoToLast, missing, missing);
                                        rng.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                        Word.Table finalTable = doc.Tables[1];
                                        bool IsTable = false;

                                        string cust = customProperties["נמענים_לפעולה"].Value;
                                        string[] existingLines = cust.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                        string last = existingLines[0];
                                        foreach (Word.Table table in doc.Tables)
                                        {
                                            foreach (Word.Cell cell in table.Range.Cells)
                                            {
                                                bool containText = cell.Range.Text.Contains(last);
                                                if (customProperties["נמענים_לפעולה"] != null && containText)
                                                {
                                                    finalTable = table;
                                                    IsTable = true;
                                                    break;
                                                }
                                            }

                                        }

                                        if (IsTable)
                                        {
                                            if (finalTable.Rows.Count == 2)
                                            {
                                                finalTable.Range.Copy();
                                                Word.Range rangeTable = finalTable.Range;
                                                finalTable.Delete();
                                                rangeTable.Text = "רשימת תפוצה - ראה בנספח" + Environment.NewLine;
                                                rangeTable.Underline = Word.WdUnderline.wdUnderlineSingle;
                                                rangeTable.BoldBi = -1;
                                                rng.Paste();
                                                Word.Table lTable = doc.Tables[doc.Tables.Count];
                                                lTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                                                foreach (Word.Cell cell in lTable.Range.Cells)
                                                {
                                                    cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                                }
                                            }
                                        }
                                        else
                                        {

                                            Word.Table newTable = doc.Tables.Add(rng, 2, 1);
                                            newTable.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent);
                                            Word.Cell cell1 = newTable.Cell(1, 1);
                                            cell1.Range.Text = "cell1";
                                            Word.Cell cell2 = newTable.Cell(2, 1);
                                            cell2.Range.Text = "cell2";

                                            Word.Table lTable = doc.Tables[doc.Tables.Count];
                                            lTable.Rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;

                                            cell1.Borders[Word.WdBorderType.wdBorderBottom].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                            cell1.Borders[Word.WdBorderType.wdBorderBottom].LineWidth = Word.WdLineWidth.wdLineWidth025pt;
                                            cell2.Borders[Word.WdBorderType.wdBorderTop].LineStyle = Word.WdLineStyle.wdLineStyleSingle;
                                            cell2.Borders[Word.WdBorderType.wdBorderTop].LineWidth = Word.WdLineWidth.wdLineWidth025pt;

                                            foreach (Word.Cell celll in lTable.Range.Cells)
                                            {
                                                celll.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                            }


                                            bool propertyExists = false;
                                            foreach (dynamic property in customProperties)
                                            {
                                                if (property.Name == "נמענים_לפעולה")
                                                {
                                                    propertyExists = true;
                                                    break;
                                                }
                                            }
                                            if (!propertyExists)
                                            {
                                                customProperties.Add("נמענים_לפעולה", false, 4, forAct);
                                            }
                                            string nameprop = "DOCPROPERTY " + "נמענים_לפעולה";
                                            entireRange = doc.Content;
                                            Word.Range lr = FindTextInRange(entireRange, "cell1");
                                            if (lr != null)
                                                doc.Fields.Add(Range: lr, Type: Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);

                                            Word.Cell cell = newTable.Cell(1, 1);
                                            Word.Range cellRange = cell.Range;
                                            cell.Range.BoldBi = -1;
                                            propertyExists = false;
                                            foreach (dynamic property in doc.CustomDocumentProperties)
                                            {
                                                if (property.Name == "נמענים_לידיעה")
                                                {
                                                    propertyExists = true;
                                                    break;
                                                }
                                            }
                                            if (!propertyExists)
                                            {
                                                doc.CustomDocumentProperties.Add("נמענים_לידיעה", false, 4, forKnow);
                                            }
                                            nameprop = "DOCPROPERTY " + "נמענים_לידיעה";
                                            entireRange = doc.Content;
                                            lr = FindTextInRange(entireRange, "cell2");
                                            if (lr != null)
                                                doc.Fields.Add(Range: lr, Type: Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("שימו לב! מסמך זה הינו מסמך ישן, ולכן לא ניתן לעדכן את המכותבים אוטומטית." + Environment.NewLine +
                                        "אנא פנו לצוות מחשוב", "מסמך ישן", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                                        MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                                    }

                                }
                                string shotef = curDoc.getID().ToString();
                                doc.Save();
                                string Text = docToTxt(doc, filePath);
                                doc.Close(true);
                                Marshal.ReleaseComObject(doc);
                                if (iswAppVisible)
                                    wapp.Visible = true;
                                //wapp.Quit();

                                //bool isWappOpen = true;
                                /*while (isWappOpen)
                                {
                                    try
                                    {
                                    
                                        Marshal.FinalReleaseComObject(wapp);
                                        wapp = null;
                                    }
                                    catch
                                    {
                                        isWappOpen = false;
                                    }
                                }*/
                                fileData = File.ReadAllBytes(filePath);
                                //גרסאות
                                DocumentHandling.SaveVersion(curDoc.getID());
                                saveDocToDB(ref fileData, curDoc.getID(), filePath, ref comm, ref conn, Text);
                                
                            }

                        }
                        catch (Exception e)
                        {
                            saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                        }
                        try
                        {
                            File.Delete(filePath);
                        }
                        catch { }


                    }
                    releaseHeldDoc(curDoc.getID());
                    docsHeldInDB.Remove(curDoc.getID());
                }
            }
            return toReturn;
        }

        internal static void updateCustomPropertiesInWordDoc(Word.Document doc, object newRefferences, string v)
        {
            throw new NotImplementedException();
        }

        private static void killProcessByProcID(int procID) 
        {
            if (procID != int.MaxValue)
            {
                Process wordProccess = null;
                try {
                    wordProccess = Process.GetProcessById(procID);
                }

                catch
                {
                    return;
                }
                wordProccess.Kill();
                while (!wordProccess.HasExited)
                {
                    wordProccess.Refresh();
                    Thread.Sleep(500);
                }

                Thread.Sleep(2000);

            }


        }
        private static void killProcess(object id) 
        {

            int procID = GetProccessIdByWindowTitle(id.ToString());
            if (procID != int.MaxValue)
            {
                Process wordProccess = Process.GetProcessById(procID);
                wordProccess.Kill();
                while (!wordProccess.HasExited)
                {
                    wordProccess.Refresh();
                    Thread.Sleep(500);
                }
            }
        }

        private static int GetProccessIdByWindowTitle(string appID)
        {
            Process[] P_CESSES = Process.GetProcesses();
            for (int p_count = 0; p_count < P_CESSES.Length; p_count++)
            {
                if (P_CESSES[p_count].MainWindowTitle.Contains(appID) && !P_CESSES[p_count].MainWindowTitle.Contains("טיפול במסמך שוטף")) // <--- added "shotef" to not close the documentHandling form instead of the word document (both have the ShotefID on their title)
                {
                    return P_CESSES[p_count].Id;
                }
            }

            return int.MaxValue;
        }

        internal static void saveDocToDB(ref byte[] fileData, int id, string filePath, ref SqlCommand comm, ref SqlConnection conn, string TextString)
            //קריאה לפונקציה ב: עדכון פרטים, שינוי במכותב וסימוכין, שכפול ועריכה של המסמך.
        {
            
            int holdsIt = whoHoldsThisDoc(id);
            if (holdsIt != 0 && holdsIt != PublicFuncsNvars.curUser.userCode) return;
            //comm = new SqlCommand("UPDATE dbo.documents SET file_data=@data WHERE shotef_mismach=@id", conn);
            comm = new SqlCommand("UPDATE dbo.documents SET file_data=@data, Txt=@TextString, LastTxtUpdateDate=@dateTime WHERE shotef_mismach=@id", conn);

            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@data", fileData);
            comm.Parameters.AddWithValue("@TextString", TextString);
            comm.Parameters.AddWithValue("@dateTime", DateTime.Now);
            
            conn.Close();
            conn.Open();
            try
            {
                comm.ExecuteNonQuery();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            
            conn.Close();
        }
        internal static DateTime? sinceWhenDocHold(int id) // A.M , cant get NULL from GetDateTime
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT OpenedForEditDTime FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            DateTime? holdDate = new DateTime?();
            if (sdr.Read())
                try { holdDate = sdr.GetDateTime(0); }
                catch { holdDate = null; }

            else
                holdDate = null;
            conn.Close();
            return holdDate;
        }
        internal static int whoHoldsThisDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT whoOpenedForEdit FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            int isOpened;
            if (sdr.Read())
                isOpened = sdr.GetInt32(0);
            else
                isOpened = 0;
            conn.Close();
            return isOpened;
        }

        internal static void replaceInWordDoc(Word.Application Wapp, string toReplace, string replaceWith) 
        {
            Wapp.Selection.Find.ClearFormatting();
            Word.Find findObject = Wapp.Selection.Find;
            findObject.ClearFormatting();
            findObject.Text = toReplace;
            findObject.Replacement.ClearFormatting();
            findObject.Replacement.Text = replaceWith;
            object missing = Type.Missing;
            findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, true, Word.WdFindWrap.wdFindContinue, ref missing,
                ref missing, true, ref missing, ref missing, ref missing, ref missing);
        }

        internal static void replaceTextInHeaderFooter(Word.Document doc,string wordToReplace, string replacementWord)//Ahava 19-02-2024 עדכון של השוטף והסווג במסמכים ישנים.
        {
            Word.Section section = doc.Sections[1];
            Word.HeaderFooter header = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            Word.Range headerRange = header.Range;
            object findText = wordToReplace;
            object replaceText = replacementWord;
            object missing = Type.Missing;
            headerRange.Find.Execute(ref findText, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceText, ref missing, ref missing, ref missing, ref missing, ref missing);

            Word.HeaderFooter footer = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary];
            Word.Range footerRange = footer.Range;
            footerRange.Find.Execute(ref findText, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref replaceText, ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        internal static void updateRefferencesInWordDoc(int id, string oldRefferences, string newRefferences) 
        {

            byte[] fileData = null;
            string fileExt = null;
            string filePath = null;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            updateDB(id, false);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                fileData = sdr.GetSqlBytes(0).Buffer;
                fileExt = sdr.GetString(1).Trim();
                filePath = Program.folderPath + "\\" + id + "." + fileExt;
                bool itscreated = false;
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                    itscreated = true;
                }
                else
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        //saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                        itscreated = true;
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את התיקים.");
                    }
                }
                
                /*if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                }*/
                if (itscreated)
                {                 
                    try
                    {
                        Word.Application wordApp;
                        try
                        {
                            wordApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                        }
                        catch
                        {
                            wordApp = new Word.Application();
                        }
                        bool iswAppVisible = wordApp.Visible;
                        if (iswAppVisible)
                            wordApp.Visible = false;
                        Word.Document doc = wordApp.Documents.Open(filePath);
                        
                        updateCustomPropertiesInWordDoc(doc, newRefferences, "סימוכין");
                        //killProcessByProcID(id);
                        string Text = docToTxt(doc, filePath);
                        doc.Close();
                        if (iswAppVisible)
                            wordApp.Visible = true;
                        //wordApp.Quit();
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, Text);
                        fileData = File.ReadAllBytes(filePath);
                        wordApp.Quit();
                    }
                    catch (Exception e)
                    {
                        saveLogError("publicFuncsNvars", e.ToString(), e.Message);

                    }                    
                }
            }
            releaseHeldDoc(id);
            docsHeldInDB.Remove(id);
        }
        
        /*internal static void updateSubjectAndClassificationInWordDoc(int id, string oldSubject, string newSubject, Classification oldClassification, Classification newClassification)
        {
            SqlConnection conn = new SqlConnection(conStr);
            updateDB(id, false);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();
                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                bool itscreated = false;
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                    itscreated = true;
                }
                else
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                        itscreated = true;
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את המכותבים.");
                    }
                }
                if (itscreated)
                {

                    try
                    {
                        Word.Application wApp;
                        try
                        {
                            wApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                        }
                        catch
                        {
                            wApp = new Word.Application();
                        }
                        wApp.Visible = false;
                        Word.Document doc = wApp.Documents.Open(filePath);
                        object missing = Type.Missing;
                        string oc = getClassificationByEnum(oldClassification);
                        string nc = getClassificationByEnum(newClassification);
                        
                        
                        doc.BuiltInDocumentProperties["Title"].Value = newSubject;
                        doc.BuiltInDocumentProperties["Category"].Value = nc;
                        
                        doc.Fields.Update();
                        doc.Save();
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                        wApp.Visible = true;
                        //wApp.Quit();
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                    }
                    catch (Exception e)
                    {
                        saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                    }
                    File.Delete(filePath);
                }
            }
            releaseHeldDoc(id);
            docsHeldInDB.Remove(id);
        }*/

        /*internal static void updateSignatureInWordDoc(int id, int newUser)
        {
            SqlConnection conn = new SqlConnection(conStr);
            updateDB(id, false);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();
                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                conn.Close();
                comm = new SqlCommand("SELECT hatimh FROM dbo.tmtafkidu WHERE kod_tpkid=@userId", conn);
                comm.Parameters.AddWithValue("@userId", newUser);
                conn.Open();
                sdr = comm.ExecuteReader();
                string[] linesNew = { "", "", "" };
                if (sdr.Read())
                {
                    string[] existingLines = sdr.GetString(0).Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                    int max = 0;
                    if (existingLines.Length > 3) max = 3; else max = existingLines.Length;
                    for (int i = 0; i < max; i++)
                        linesNew[i] = existingLines[i];
                }
                conn.Close();
                bool itscreated = false;
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                    itscreated = true;
                }
                else
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                        itscreated = true;
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את המכותבים.");
                    }
                }
                if (itscreated)
                {
                    try
                    {
                        object missing = Type.Missing;
                        Word.Application wApp;
                        try
                        {
                            wApp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                        }
                        catch
                        {
                            wApp = new Word.Application();
                        }
                        wApp.Visible = false;
                        Word.Document doc = wApp.Documents.Open(filePath);
                        try
                        {
                            wApp = (Word.Application)Marshal.GetActiveObject($"Word.Application");
                        }
                        catch (Exception error)
                        {
                            MessageBox.Show(error.Message);
                        }
                        bool itsOpenn = false;
                        foreach (Word.Document docc in wApp.Documents)
                        {
                            if (docc.FullName == filePath)
                            {
                                itsOpenn = true;
                                doc = docc;
                                break;
                            }
                        }
                        if (!itsOpenn)
                        {
                            doc = wApp.Documents.Open(filePath);
                        }
                        updateCustomPropertiesInWordDoc(doc, linesNew[0], "חתימה_שורה_א");
                        updateCustomPropertiesInWordDoc(doc, linesNew[1], "חתימה_שורה_ב");
                        updateCustomPropertiesInWordDoc(doc, linesNew[2], "חתימה_שורה_ג");
                        wApp.WindowState = Word.WdWindowState.wdWindowStateMinimize;
                        //int procID = GetProccessIdByWindowTitle(id.ToString());
                        doc.Close(true);
                        Marshal.ReleaseComObject(doc);
                        wApp.Visible = true;
                        //wApp.Quit();
                        //killProcessByProcID(id);
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                    }
                    catch (Exception e)
                    {
                        saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                    }
                }
                
                File.Delete(filePath);

            }
            releaseHeldDoc(id);
            docsHeldInDB.Remove(id);
        }*/

        private static string getUserRoleById(int id)
        {
            foreach (User u in users)
                if (u.userCode == id || (u.userCode == 100 && id == 110))
                    return u.job;
            return null;
        }

        internal static string getDocExt(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                return sdr.GetString(0);
            }
            return null;
        }


        internal static bool isAuthorizedUser(int owner, User checkCommander)
        {
            if (checkCommander.userCode == owner)
            {
                return true;
            }
            else if (owner == 0 || owner == 99999)
            {
                return true;
            }
            else
            {

                User u = getUserByCode(owner);
                if (u == null) return true; // if user not exist, allow anyone to edit it.

                if ((u.permissionsBranch == checkCommander.permissionsBranch && checkCommander.roleType == RoleType.clerk) ||
                    (Branch.office == checkCommander.permissionsBranch && checkCommander.roleType == RoleType.clerk) ||
                    ((u.permissionsBranch == checkCommander.permissionsBranch || checkCommander.permissionsBranch == Branch.office)
                    && checkCommander.roleType == RoleType.branchHead) ||
                    checkCommander.roleType == RoleType.computers || checkCommander.userCode == u.commanderCode)
                {
                    return true;
                }
                else
                {
                    List<User> subordinates = getUserDirectSubordinates(checkCommander.userCode);
                    foreach (User sub in subordinates)
                    {
                        if (PublicFuncsNvars.curUser.userCode != checkCommander.userCode && isAuthorizedUser(owner, sub))
                            return true;
                    }
                    return false;
                }
            }
        }

        private static List<User> getUserDirectSubordinates(int userCode)
        {
            List<User> subs = new List<User>();
            foreach (User u in users)
            {
                if ( u.commanderCode == userCode || (u.commanderCode == 100 && userCode == 110))
                    subs.Add(u);
            }
            return subs;
        }

        internal static void textBox_Enter()
        {
            Application.CurrentInputLanguage = InputLanguage.FromCulture(new System.Globalization.CultureInfo("he-IL"));
        }

        internal static bool shortDescExists(string newShortDesc)
        {
            foreach (Folder f in folders)
                if (f.shortDescription.Equals(newShortDesc))
                    return true;
            return false;
        }

        internal static string getfileTypeString(FileType fileType)
        {
            switch (fileType)
            {
                case FileType.development:
                    return "פיתוח";
                case FileType.kit:
                    return "ערכה";
                case FileType.occasional:
                    return "מזדמן";
                case FileType.organization:
                    return "ארגון";
                case FileType.regular:
                    return "רגילה";
                case FileType.serial:
                    return "סדרתי";
                case FileType.shosh:
                    return "שו\"ש";
                case FileType.spares:
                    return "חלפים";
                case FileType.steelPlates:
                    return "לוחות פלדה";
                case FileType.subject:
                    return "נושא";
                case FileType.tzbm:
                    return "צב\"מ";
            }
            return null;
        }

        internal static bool valueExistsInObjectProp<T1, T2>(List<T1> objList, string prop, T2 searchValue)
        {
            var type1 = objList[0].GetType();

            PropertyInfo pi = type1.GetProperty(prop, BindingFlags.NonPublic | BindingFlags.Public | BindingFlags.Instance);

            foreach (T1 obj in objList)
            {
                dynamic value = pi.GetValue(obj, null);

                if (value != null && value.Equals(searchValue))
                    return true;
            }
            return false;
        }

        internal static bool userLoginExists(string login)
        {
            return valueExistsInObjectProp(users, "userLogin", login);
        }

        internal static bool userCodeExists(int userCode)
        {
            return valueExistsInObjectProp(users, "userCode", userCode);
        }

        internal static bool validEmail(string p)
        {
            string[] firstSplit = p.Split('@');
            if (firstSplit.Length != 2)
                return false;
            string[] secondSplit = firstSplit[1].Split('.');
            if (secondSplit.Length < 2)
                return false;
            if (secondSplit[1] == "")
                return false;
            return true;
        }

        internal static void exportToXL(string fileName, string headLine, string[] headers, List<string[]> values)
        {
            Excel.Application xl = new Excel.Application();
            xl.Visible = false;
            xl.DisplayAlerts = false;
            Excel.Workbook wb = xl.Workbooks.Add(Type.Missing);
            Excel.Worksheet ws = (Excel.Worksheet)wb.ActiveSheet;
            ws.Name = fileName;
            ws.Range[ws.Cells[3, 1], ws.Cells[3, 3]].Merge();
            ws.Cells[1, 1] = "הופק ע\"י";
            ws.Cells[1, 2] = "בתאריך";
            ws.Cells[1, 3] = "בשעה";
            ws.Cells[2, 1] = PublicFuncsNvars.curUser.userCode.ToString();
            ws.Cells[2, 2] = DateTime.Today.ToShortDateString();
            ws.Cells[2, 3] = DateTime.Now.ToLongTimeString();
            ws.Cells[3, 1] = headLine;

            ws.Cells.Font.Name = "Arial";
            ws.Cells.Font.Size = 11;
            ws.Cells.Font.Color = Color.Black;
            for (int i = 1; i <= headers.Length; i++)
                ws.Cells[4, i] = headers[i - 1];
            ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, ws.Range[ws.Cells[4, 1], ws.Cells[4, headers.Length]], Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing);



            int dataRow = 5;

            foreach (string[] valRow in values)
            {
                for (int i = 1; i <= valRow.Length; i++)
                    ws.Cells[dataRow, i] = valRow[i - 1];

                dataRow++;
            }

            Excel.Range range = ws.Range[ws.Cells[1, 1], ws.Cells[dataRow, headers.Length]];
            range.EntireColumn.AutoFit();
            range.EntireRow.AutoFit();
            range = ws.Range[ws.Cells[5, 1], ws.Cells[dataRow, headers.Length]];
            range.Style.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
            range = ws.Range[ws.Cells[4, 1], ws.Cells[dataRow, headers.Length]];
            Excel.Borders b = range.Borders;
            b.LineStyle = Excel.XlLineStyle.xlContinuous;
            b.Weight = 2d;

            range = ws.Range["A1", "C3"];
            b = range.Borders;
            b.LineStyle = Excel.XlLineStyle.xlContinuous;
            b.Weight = 2d;

            range = ws.Range[ws.Cells[1, 1], ws.Cells[4, headers.Length]];
            range.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            range = ws.Range[ws.Cells[1, 1], ws.Cells[1, 3]];
            range.Interior.Color = Color.Silver;

            range = ws.Range[ws.Cells[3, 1], ws.Cells[3, 1]];
            range.Interior.Color = Color.MistyRose;

            range = ws.Range[ws.Cells[4, 1], ws.Cells[4, headers.Length]];
            range.Interior.Color = Color.LightGoldenrodYellow;

            xl.Visible = true;
            xl.DisplayAlerts = true;
        }

        internal static bool userloginMatchesUserCode(string login, int code)
        {
            return users.Count(x => (x.userLogin == login && x.userCode == code)) == 1;
        }

        internal static RecipientListsLevel getRecipientListLevelByString(string level)
        {
            switch (level)
            {
                case "אישית":
                    return RecipientListsLevel.personal;
                case "ענפית":
                    return RecipientListsLevel.branch;
                case "יחידתית":
                    return RecipientListsLevel.unit;
            }
            return RecipientListsLevel.unit;
        }

        static string[] ranks = { "טור'", "רב''ט", "סמל", "סמ''ר", "רס''ל", "רס''ר", "רס''ם", "רס''ב", "רנ''ג", "סג''ם", "סגן", "סרן", "רס''ן", "סא''ל", "אל''ם",
                             "תא''ל", "אלוף", "רא''ל", "אע''צ", "בח''ג", "קמ''א","קא''ב" };
        //adiing the func as done on the internet - from the dll...
        [DllImport("USER32.DLL")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy, uint uFlags);
        [DllImport("USER32.DLL")]
        //public static extern bool ShowWindow(IntPtr hWnd, uint windowStyle);
        //[DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);


        internal static List<string> CtrlK()//List<Tuple<string, string, string, bool>>
        {
            List<string> allRecipients = new List<string>();
            //SetWindowPos(Process.GetCurrentProcess().MainWindowHandle, new IntPtr(0), 50, 50, 0, 0, 0);
            //Outlook.Account acc = new Outlook.Account();
            Outlook.Application oApp = new Outlook.Application(); ;
            Process[] pr = Process.GetProcessesByName("OUTLOOK");
            Outlook.MailItem oMilItem = (Outlook.MailItem)oApp.CreateItem(Outlook.OlItemType.olMailItem);

            Outlook.NameSpace space = oApp.GetNamespace("MAPI");

            Outlook.AddressList AddKList = space.GetGlobalAddressList();
            Outlook.AddressLists adls = space.Session.AddressLists;

            Outlook.MAPIFolder f = AddKList.GetContactsFolder();
            Outlook.SelectNamesDialog snd = oApp.Session.GetSelectNamesDialog();
            try
            {
                int prI = 0;
                pr = Process.GetProcesses();
                for (int i = 0; i < pr.Length; i++)
                {
                    if (pr[i].ProcessName.Contains("OUTLOOK"))
                    {
                        prI = i;
                        break;
                    }

                }
                SetForegroundWindow(pr[prI].MainWindowHandle);
                snd.Display();
                Outlook.Recipients res = snd.Recipients;
                Outlook.ExchangeUser eu;
                
                for (int i = 1; i <= res.Count; i++)
                {
                    eu = res[i].AddressEntry.GetExchangeUser();
                    string name = eu.Name;
                    string address = eu.PrimarySmtpAddress;
                    string jobTitle = trimRank(eu.JobTitle);
                    int type = eu.Parent.Type;
                    char delimiter = (char)242;
                    allRecipients.Add(name + delimiter + address + delimiter + jobTitle + delimiter + type);
                }
                //System.IO.File.WriteAllLines(@"c:\temp\address.txt", allRecipients);
            }
            catch (Exception e)
            {
                Application.Exit();
                
            }
            return allRecipients;
        }
        internal static string trimRank(string title)
        {
            for (int i = 0; i < ranks.Length; i++)
                if (title.StartsWith(ranks[i]))
                    return title.Substring(title.IndexOf('-') + 2);
            return title;
        }
                        
        internal static List<Tuple<string, string, string, bool>> getCtrlKRecipients() //חדש CtrlK
        {
            List<string> rec = CtrlK();
            List<Tuple<string, string, string, bool>> recTuple = new List<Tuple<string, string, string, bool>>(); 
            foreach (string r in rec)
            {
                string[] rParts = r.Split((char)242);
                bool ifa = rParts[3] == "1";
                recTuple.Add(new Tuple<string, string, string, bool>(rParts[0], rParts[1], rParts[2], ifa));
            }
            return recTuple;
        }

        internal static bool isCurUserAllowedToWatchDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT roleCode FROM dbo.doc_Authorizations WHERE docId=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                conn.Close();
                comm = new SqlCommand("SELECT roleCode FROM dbo.doc_Authorizations WHERE docId=@id AND roleCode=@userCode", conn);
                comm.Parameters.AddWithValue("@id", id);
                comm.Parameters.AddWithValue("@userCode", curUser.userCode);
                conn.Open();
                sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    conn.Close();
                    return true;
                }
                else
                {
                    conn.Close();
                    return false;
                }
            }
            return true;
        }


        internal static bool isAllowedToRagish(int shotef)
        {

            List<string> allowedUsers = new List<string>();
            bool isRagish = false;
            SqlConnection conn2 = new SqlConnection(Global.ConStr);
            SqlCommand comm2 = new SqlCommand("SELECT isRagish from dbo.documents (nolock) WHERE shotef_mismach=@shotef", conn2);
            comm2.Parameters.AddWithValue("@shotef", shotef);
            conn2.Open();
            SqlDataReader sdr2 = comm2.ExecuteReader();

            while (sdr2.Read())
            {
                isRagish = sdr2.GetBoolean(0);
            }

            if (!isRagish) return true;

            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT kod_sholeah as Tafkid from dbo.documents (nolock) WHERE shotef_mismach=@shotef union select kod_mechutav from doc_mech where shotef_klali = @shotef", conn);
            comm.Parameters.AddWithValue("@shotef", shotef);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();

            while (sdr.Read())
            {
                allowedUsers.Add(sdr.GetInt32(0).ToString());
            }

            if (allowedUsers.Contains(curUser.userCode.ToString())) return true;
            return false;

        }
        internal static bool isCurUserAllowedToEditDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT roleCode FROM dbo.doc_Authorizations WHERE docId=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                conn.Close();
                comm = new SqlCommand("SELECT roleCode FROM dbo.doc_Authorizations WHERE docId=@id AND roleCode=@userCode AND isForEdit=@true", conn);
                comm.Parameters.AddWithValue("@id", id);
                comm.Parameters.AddWithValue("@userCode", curUser.userCode);
                comm.Parameters.AddWithValue("@true", true);
                conn.Open();
                sdr = comm.ExecuteReader();
                if (sdr.Read())
                {
                    conn.Close();
                    return true;
                }
                else
                {
                    conn.Close();
                    return false;
                }
            }
            return true;
        }

        internal static bool isNormalDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            // בדיקה האם יש סוג לנספח , אזי הוא מגיע ממערכת משימות, ויש לקחת את המכותבים שלו
            SqlCommand comm_tm_nig_mis = new SqlCommand("SELECT typ_rcrd FROM dbo.tm_nig_mis WHERE sn_doc_gnrl=@id", conn);
            comm_tm_nig_mis.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr_tm_nig_mis = comm_tm_nig_mis.ExecuteReader();

            bool IsReadTm_nig_mis = sdr_tm_nig_mis.Read();
            conn.Close();

            // בדיקה אם יש מסמך וורד או אקסל  בתוך הרשומה
            SqlCommand comm_documents = new SqlCommand("SELECT file_data FROM dbo.documents (nolock) WHERE shotef_mismach=@id", conn);
            comm_documents.Parameters.AddWithValue("@id", id);
            conn.Open();

            SqlDataReader sdr_documents = comm_documents.ExecuteReader();

            bool IsReadDocuments = sdr_documents.Read();

            isRegularDocument = !IsReadDocuments;

            conn.Close();

            if (IsReadTm_nig_mis && IsReadDocuments)
            {
                conn.Close();
                return false;
            }
            else
            {
                conn.Close();
                return true;
            }
        }

        internal static bool isValidEmail(string email)
        {
            Regex r = new Regex(@".+@[^\.]+(\.[^\.]+)+");
            return r.IsMatch(email);
        }

        internal static bool releaseDoc(int id)
        {

            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=0 WHERE shotef_mismach=@id", conn);
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            bool ok;
            try
            {
                comm.ExecuteNonQuery();
                ok = true;

            }
            catch (Exception e)
            {
                saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                MessageBox.Show("שחרור המסמך נכשל. פנו לצוות מחשוב." + Environment.NewLine + Environment.NewLine + e.Message,
                                    "שחרור מסמך", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                ok = false;
            }
            finally
            {
                conn.Close();
            }
            return ok;

        }
        internal static void releaseAllHeldDocs()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=0 WHERE shotef_mismach=@id", conn);
            comm.CommandTimeout = 0;
            foreach (int id in docsHeldInDB)
            {
                comm.Parameters.Clear();
                comm.Parameters.AddWithValue("@id", id);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }
        }


       


        private static bool isOutlookLoaded(ref Microsoft.Office.Interop.Outlook.Application app)
        {
            bool retValue = true;
            try
            {
                app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");

            }

            catch
            {
                retValue = false;
            }

            return retValue;
        }


        public static string CreateShortcut(int id)
        {
            try
            {
                string dirPath = Global.P_APP  + "\\Shortcuts\\";//Application.StartupPath
                string filename = Global.P_APP + "\\Shortcuts\\" + id + ".cmd";//Application.StartupPath
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                using (StreamWriter writer = System.IO.File.CreateText(filename))
                {
                    string command = Global.P_APP + "\\MantakDocumentsShortcut2.cmd -s " + id; //Application.StartupPath
                    writer.WriteLine("@echo off");
                    writer.WriteLine(command);
                    return filename;
                }
            }

            catch (Exception e)
            {
                MessageBox.Show(e.Message);
                return "";
            }
        }

        public static void sendShareMail(string from, string to, string cc, string bcc, string subject, string body, List<Tuple<byte[], string>> attsMs)
        {
            try
            {
                Encoding enc = Encoding.GetEncoding(1255);
                string[] lines = { "FROM: " + from,
                                 "TO: " + to,
                                 "CC: " + cc,
                                 "BCC: " + bcc,
                                 "SUBJECT: " + subject,
                                 "BODY:" + Environment.NewLine + body };
                string dirPath = Program.folderPath;
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                string filePath = dirPath + "\\sendmail.txt";
                File.WriteAllLines(filePath, lines);



                Microsoft.Office.Interop.Outlook.Application app = null;

                try
                {
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    var processStartInfo = new ProcessStartInfo() { FileName = "outlook", WindowStyle = ProcessWindowStyle.Minimized };
                    Process.Start(processStartInfo);
                    while (!isOutlookLoaded(ref app)) ;

                }

                if (app == null) return;

                Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;
                mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                mailItem.SentOnBehalfOfName = from;
                mailItem.Subject = subject;
                mailItem.To = to;
                mailItem.CC = cc;
                mailItem.BCC = bcc;
                body = "<a href='" + body + "'>" + "קישור למסמך" + "</a>";
                body = "<p DIR=\"RTL\">" + body + "</p>";
                mailItem.HTMLBody = body;
                if (attsMs != null)
                {
                    foreach (var item in attsMs)
                    {
                        string fixedFilename = item.Item2;
                        foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                        {
                            fixedFilename = fixedFilename.Replace(c, '_');
                        }

                        string outputFile = "C:\\temp\\" + fixedFilename; // Path.GetTempPath() + item.Item2;

                        System.IO.File.WriteAllBytes(outputFile, item.Item1);
                        mailItem.Attachments.Add(outputFile, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, item.Item2);
                    }
                }
                System.Threading.Thread.Sleep(2000);
                mailItem.Display(true);
            }
            catch (Exception ex)
            {
                if (ex.Message.Contains("שיח"))
                {
                    MessageBox.Show("נראה שיש כבר חלון שליחת הודעה פתוח באווטלוק, אנא סגור אותו ונסה שנית", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign);
                }

                else
                    MessageBox.Show(ex.GetType().ToString() + ":\n " + ex.Message + "\n\n, Stacktrace: " + ex.StackTrace);
            }
        }

        public static object[] sivug_by_reshet()   //יערה הוסיפה ב-19.03.23 כדי שהסיווגים סודי ומעלה יופיעו רק ברשת הסגולה
        {
            object[] sivug_comboBox = new object[]  {
            "בלמ\"ס",
            "מוגבל", //Obsolete
            "שמור"};
            //if (Global.ConStr.ToUpper().Contains("TOP")) // || true) // Remove by Danny 16/05/2024
            if (int.Parse(Global.P_MAX_CLASS) == 4) // Danny Add at 16/05/2024
            {
                object[] sivug_comboBox_top = new object[]  {
                 "סודי" };
                sivug_comboBox = sivug_comboBox.Concat(sivug_comboBox_top).ToArray();
            }
            else if (int.Parse(Global.P_MAX_CLASS)>4) // Danny Add at 16/05/2024
                {
                object[] sivug_comboBox_top = new object[]  {
                 "סודי",
                 "סודי ביותר",
                "סודי לשו\"ס",
                "סודי ביותר לשו\"ס" };
                sivug_comboBox = sivug_comboBox.Concat(sivug_comboBox_top).ToArray();
            }
            return sivug_comboBox;
        }
        public static void sendAdvancedMail(string from, string to, string cc, string bcc, string subject, string body, List<Tuple<byte[], string, bool>> attsMs)
        {
            try
            {
                Encoding enc = Encoding.GetEncoding(1255);
                string[] lines = { "FROM: " + from,
                                 "TO: " + to,
                                 "CC: " + cc,
                                 "BCC: " + bcc,
                                 "SUBJECT: " + subject,
                                 "BODY:" + Environment.NewLine + body };
                string dirPath = Program.folderPath;
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                string filePath = dirPath + "\\sendmail.txt";
                File.WriteAllLines(filePath, lines);



                Microsoft.Office.Interop.Outlook.Application app = null;

                try
                {
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    var processStartInfo = new ProcessStartInfo() { FileName = "outlook", WindowStyle = ProcessWindowStyle.Minimized };
                    Process.Start(processStartInfo);
                    while (!isOutlookLoaded(ref app)) ;

                }

                if (app == null) return;

                Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;
                mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                mailItem.SentOnBehalfOfName = from;
                mailItem.Subject = subject;
                mailItem.To = to;
                mailItem.CC = cc;
                mailItem.BCC = bcc;
                body = "<p DIR=\"RTL\">" + body + "</p>";
                mailItem.HTMLBody = body;
                foreach (var item in attsMs)
                {
                    string fixedFilename = item.Item2;
                    foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                    {
                        fixedFilename = fixedFilename.Replace(c, '_');
                    }

                    string outputFile = "C:\\temp\\" + fixedFilename; // Path.GetTempPath() + item.Item2;

                    System.IO.File.WriteAllBytes(outputFile, item.Item1);
                    if (item.Item3) // convert To PDF
                    {
                        outputFile = ConvertToPDF(outputFile);
                    }
                    mailItem.Attachments.Add(outputFile, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, item.Item2);
                }
                System.Threading.Thread.Sleep(2000);
                mailItem.Display(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetType().ToString() + ": " + ex.Message + ", Stacktrace: " + ex.StackTrace);
            }
        }

        private static string ConvertToPDF(string filename)
        {
            bool isWord = false;
            bool isExcel = false;
            bool isPPT = false;
            List<string> WordTypes = new List<string> { ".doc", ".docx", ".dot", ".dotx" };
            List<string> ExcelTypes = new List<string> { ".xl", ".xlsx", ".xlsm", ".xlsb", ".xlam", ".xls" };
            List<string> PptTypes = new List<string> { ".ppt", ".pptm", ".pptx", ".ptx" };
            string extention = Path.GetExtension(filename);
            string name = Path.GetFileNameWithoutExtension(filename);
            if (WordTypes.Any(s => extention.ToLower().Equals(s))) isWord = true;
            if (ExcelTypes.Any(s => extention.ToLower().Equals(s))) isExcel = true;
            if (PptTypes.Any(s => extention.ToLower().Equals(s))) isPPT = true;

            try
            {

                if (isWord)
                {
                    var appWord = new Microsoft.Office.Interop.Word.Application();
                    var wordDocument = appWord.Documents.Open(filename);
                    appWord.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                    wordDocument.ExportAsFixedFormat("c:\\temp\\" + name + ".pdf", Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF);
                    wordDocument.Close();
                    appWord.Quit();
                    appWord = null;
                    return ("c:\\temp\\" + name + ".pdf");
                }


                else if (isExcel)
                {
                    var appExcel = new Microsoft.Office.Interop.Excel.Application();
                    var workbooks = appExcel.Workbooks;
                    var workbook = workbooks.Open(filename);

                    appExcel.DisplayAlerts = false;
                    workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, "c:\\temp\\" + name + ".pdf");
                    workbook.Close();
                    appExcel.Quit();
                    Marshal.ReleaseComObject(workbook);
                    Marshal.ReleaseComObject(workbooks);
                    Marshal.ReleaseComObject(appExcel);
                    return ("c:\\temp\\" + name + ".pdf");
                }


                else if (isPPT)
                {
                    var appPPT = new Microsoft.Office.Interop.PowerPoint.Application();
                    var pres = appPPT.Presentations;
                    var pptDocument = pres.Open(filename, WithWindow: 0);
                    appPPT.DisplayAlerts = Microsoft.Office.Interop.PowerPoint.PpAlertLevel.ppAlertsNone;
                    pptDocument.ExportAsFixedFormat("c:\\temp\\" + name + ".pdf", Microsoft.Office.Interop.PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);
                    pptDocument.Close();
                    appPPT.Quit();
                    Marshal.ReleaseComObject(pptDocument);
                    Marshal.ReleaseComObject(pres);
                    Marshal.ReleaseComObject(appPPT);
                    return ("c:\\temp\\" + name + ".pdf");
                }

            }

            catch
            {
                return filename;
            }

            return filename;


        }

        public static void sendMail(string from, string to, string cc, string bcc, string subject, string body, List<Tuple<byte[], string>> attsMs)
        {
            try
            {
                Encoding enc = Encoding.GetEncoding(1255);
                string[] lines = { "FROM: " + from,
                                 "TO: " + to,
                                 "CC: " + cc,
                                 "BCC: " + bcc,
                                 "SUBJECT: " + subject,
                                 "BODY:" + Environment.NewLine + body };
                string dirPath = Program.folderPath;
                if (!Directory.Exists(dirPath))
                    Directory.CreateDirectory(dirPath);
                string filePath = dirPath + "\\sendmail.txt";
                File.WriteAllLines(filePath, lines);



                Microsoft.Office.Interop.Outlook.Application app = null;

                try
                {
                    app = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
                }
                catch
                {
                    var processStartInfo = new ProcessStartInfo() { FileName = "outlook", WindowStyle = ProcessWindowStyle.Minimized };
                    Process.Start(processStartInfo);
                    while (!isOutlookLoaded(ref app)) ;

                }

                if (app == null) return;

                Microsoft.Office.Interop.Outlook.MailItem mailItem = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;
                mailItem.BodyFormat = Microsoft.Office.Interop.Outlook.OlBodyFormat.olFormatHTML;
                mailItem.SentOnBehalfOfName = from;
                mailItem.Subject = subject;
                mailItem.To = to;
                mailItem.CC = cc;
                mailItem.BCC = bcc;
                body = "<p DIR=\"RTL\">" + body + "</p>";
                mailItem.HTMLBody = body;
                if (attsMs != null)
                {
                    foreach (var item in attsMs)
                    {
                        string fixedFilename = item.Item2;
                        foreach (char c in System.IO.Path.GetInvalidFileNameChars())
                        {
                            fixedFilename = fixedFilename.Replace(c, '_');
                        }

                        string outputFile = "C:\\temp\\" + fixedFilename; // Path.GetTempPath() + item.Item2;

                        System.IO.File.WriteAllBytes(outputFile, item.Item1);
                        mailItem.Attachments.Add(outputFile, Microsoft.Office.Interop.Outlook.OlAttachmentType.olByValue, Type.Missing, item.Item2);
                    }
                }
                System.Threading.Thread.Sleep(2000);
                mailItem.Display(true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.GetType().ToString() + ": " + ex.Message + ", Stacktrace: " + ex.StackTrace);
            }
        }


        internal static bool releaseHeldDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=0 WHERE shotef_mismach=@id", conn);
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            bool ok;
            try
            {
                comm.ExecuteNonQuery();
                ok = true;
                docsHeldInDB.Remove(id);
            }
            catch (Exception e)
            {
                saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                MessageBox.Show("שחרור המסמך נכשל. פנו לצוות מחשוב." + Environment.NewLine + Environment.NewLine + e.Message,
                                    "שחרור מסמך", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                ok = false;
            }
            finally
            {
                conn.Close();
            }
            return ok;
        }

        internal static void saveLogError(string formName, string exceptionError, string exceptionMessage)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("INSERT INTO dbo.documentsModuleErrorLog(dateNtime, formName, exceptionError, exceptionMessage)"
                + " VALUES(@dateNtime, @formName, @exceptionError, @exceptionMessage)", conn);
            comm.Parameters.AddWithValue("@dateNtime", DateTime.Now);
            comm.Parameters.AddWithValue("@formName", formName);
            comm.Parameters.AddWithValue("@exceptionError", exceptionError);
            comm.Parameters.AddWithValue("@exceptionMessage", exceptionMessage);
        }

        internal static void openDocumentHandlingForm(int docId)
        {
            if (!PublicFuncsNvars.dhFormsOpen.Contains(docId))
            {
                DocumentHandling dh = new DocumentHandling(docId);
                dh.Activate();
                dh.ShowDialog();
            }
        }

        internal static bool signDoc(int docId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            updateDB(docId, false);
            SqlCommand comm = new SqlCommand("SELECT hanadon, file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@docId AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@docId", docId);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            string subject = sdr.GetString(0).Trim();
            byte[] fileData = sdr.GetSqlBytes(1).Buffer;
            string fileExt = sdr.GetString(2).Trim();
            conn.Close();
            string filePath = Program.folderPath + "\\" + docId + "." + fileExt;
            string pdfPath = Program.folderPath + "\\" + docId + ".pdf";

            /*comm = new SqlCommand("SELECT signature FROM dbo.tmtafkidu WHERE kod_tpkid=@userCode AND signature IS NOT NULL", conn);
            comm.Parameters.AddWithValue("@userCode", curUser.userCode);
            conn.Open();
            sdr = comm.ExecuteReader();
            if (!sdr.Read())
            {
                conn.Close();
                MessageBox.Show("קיימת בעיה בקובץ החתימה שלך." + Environment.NewLine + "המסמך לא נחתם.", "חתימה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                return false;
            }
            string picPath = Program.folderPath + "\\" + docId + ".png";
            File.WriteAllBytes(picPath, sdr.GetSqlBytes(0).Buffer);
            conn.Close();*/

            string picPath = Program.folderPath + "\\" + PublicFuncsNvars.curUser.userCode.ToString() + ".png";
            Cursor.Current = Cursors.WaitCursor;
            Signature signature = new Signature();
            signature.Activate();
            signature.ShowDialog();
            Cursor.Current = Cursors.Default;


            bool itscreated = false;
            if (!File.Exists(filePath))
            {
                File.WriteAllBytes(filePath, fileData);
                itscreated = true;
            }
            else
            {
                try
                {
                    fileData = File.ReadAllBytes(filePath);
                    //saveDocToDB(ref fileData, docId, filePath, ref comm, ref conn);
                    itscreated = true;
                }
                catch
                {
                    MessageBox.Show("המסמך פתוח לכן לא ניתן לחתום על המסמך.");
                    return false;
                }
            }
            if (itscreated)
            {
                try
                {
                    //Word.Application wapp = new Word.Application();
                    Word.Application wapp;
                    try
                    {
                        wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                    }
                    catch
                    {
                        wapp = new Word.Application();
                    }
                    bool iswAppVisible = wapp.Visible;
                    if (iswAppVisible)
                        wapp.Visible = false;
                    object missing = Type.Missing;
                    Word.Document doc = wapp.Documents.Open(filePath, missing, missing, missing, missing, missing, missing, missing,
                        missing, missing, missing, false);
                    doc.Select();
                    dynamic customProperties = doc.CustomDocumentProperties;
                    bool IsDocNewVersion = false;
                    try
                    {
                        dynamic existingProperty = customProperties["סימוכין"];
                        IsDocNewVersion = true;
                    }
                    catch
                    { }
                    if (IsDocNewVersion)
                    {
                        
                        Word.Table finalTable = doc.Tables[1];
                        bool IsTable = false;
                        foreach (Word.Table table in doc.Tables)
                        {
                            foreach(Word.Cell cell in table.Range.Cells)
                            {
                               
                                if (customProperties["חתימה_שורה_א"] != null && cell.Range.Text.Contains(customProperties["חתימה_שורה_א"].Value.ToString()))
                                {
                                    finalTable = table;
                                    IsTable = true;
                                    break;
                                }
                            }
                            
                        }
                        if (IsTable)
                        {
                            Word.InlineShape pic = finalTable.Range.InlineShapes.AddPicture(picPath);
                            pic.Height = 55;
                            pic.Width = 158;

                            Word.Shape shape = pic.ConvertToShape();

                            shape.WrapFormat.Type = Word.WdWrapType.wdWrapBehind;
                            DialogResult result = MessageBox.Show(" המסמך נחתם בהצלחה . האם ברצונך להציגו ?", " חתימה ", MessageBoxButtons.YesNo, MessageBoxIcon.Question,
                                    MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            if (DialogResult.Yes == result)
                            {
                                doc.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF, true);
                            }
                            else
                                doc.ExportAsFixedFormat(pdfPath, Word.WdExportFormat.wdExportFormatPDF, false);
                            

                            
                            wapp.WindowState = Word.WdWindowState.wdWindowStateMinimize;
                            int procID = GetProccessIdByWindowTitle(docId.ToString());
                        }

                        doc.Save();
                        doc.Close();
                        //File.Delete(filePath);
                        if (iswAppVisible)
                            wapp.Visible = true;
                        Marshal.ReleaseComObject(doc);
                        //wapp.Quit();
                        
                        //killProcessByProcID(docId);
                    }
                    else
                    {
                        MessageBox.Show("שימו לב! מסמך זה הינו מסמך ישן, ולכן לא ניתן לשנות בו את חתימת המשתמש אוטומטית." + Environment.NewLine +
                            "אנא פנו לצוות מחשוב", "מסמך ישן", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                        if (iswAppVisible)
                            wapp.Visible = true;
                            //wapp.Visible = false;  // 18_10_21 A.M
                        wapp.WindowState = Word.WdWindowState.wdWindowStateMinimize;
                        int procID = GetProccessIdByWindowTitle(docId.ToString());
                        doc.Save();
                        doc.Close();
                        File.Delete(filePath);
                        Marshal.ReleaseComObject(doc);
                        //wapp.Quit();
                        //killProcessByProcID(docId);
                        return false;
                    }

                    fileData = File.ReadAllBytes(pdfPath);
                    string command = "INSERT INTO dbo.docnisp (shotef_mchtv, shotef_nisph, kod_marcht, kod_sug_nsph, msd_sruk, msd_df, prtim, tarich," +
                        "shm_kovtz, is_pail, shotf_mmh, kod_sivug_bithoni, is_yetzu, is_sodi, bealim, is_ishi, is_anafi, kod_kvatzaim," +
                        " user_sorek, tarich_srika, is_letzaref_mail, mail_id, ocr, colorscan, Txt, LastTxtUpdateDate, file_data, file_extension)" +
                        Environment.NewLine + "output inserted.shotef_nisph" + Environment.NewLine +
                        " VALUES (@docId, (SELECT MAX(shotef_nisph) FROM dbo.docnisp)+1, 1, 0, @msdsruk+1," +
                        " 0, @name, @date, @name, 1, 0, (SELECT kod_sivug_bitchoni FROM MantakDB.dbo.documents WHERE shotef_mismach=@docId), 1, 0, @owner," +
                        " 0, 0, 0, '', '00000000', 0, '', 0, 0, NULL, NULL, @data, @ext)";
                    comm = new SqlCommand("SELECT CASE" + Environment.NewLine + "WHEN MAX(msd_sruk) IS NULL THEN 0" + Environment.NewLine + "ELSE MAX(msd_sruk)" +
                        Environment.NewLine + "END" + Environment.NewLine + "FROM dbo.docnisp WHERE shotef_mchtv=@docId", conn);
                    comm.Parameters.AddWithValue("@docId", docId);
                    conn.Open();
                    int msd = (int)comm.ExecuteScalar();
                    conn.Close();
                    comm = new SqlCommand(command, conn);
                    comm.Parameters.AddWithValue("@docId", docId);
                    comm.Parameters.AddWithValue("@name", subject);
                    string date = DateTime.Today.ToString("yyyyMMdd");
                    comm.Parameters.AddWithValue("@date", date);
                    comm.Parameters.AddWithValue("@owner", curUser.userCode);
                    comm.Parameters.AddWithValue("@data", fileData);
                    comm.Parameters.AddWithValue("@ext", "pdf");
                    comm.Parameters.AddWithValue("@msdsruk", msd);
                    conn.Open();
                    sdr = comm.ExecuteReader();
                    sdr.Read();
                    int id = sdr.GetInt32(0);
                    conn.Close();

                    comm = new SqlCommand("UPDATE dbo.documents SET signedAttId=@attId, isSigned=@true, dateSigned=@date WHERE shotef_mismach=@docId", conn);
                    comm.CommandTimeout = 0;
                    comm.Parameters.AddWithValue("@docId", docId);
                    comm.Parameters.AddWithValue("@attId", id);
                    comm.Parameters.AddWithValue("@true", true);
                    comm.Parameters.AddWithValue("@date", DateTime.Today);
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                }
                catch (Exception e)
                {
                    saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                }
                File.Delete(filePath);
                File.Delete(picPath);
                File.Delete(pdfPath);
            }

            releaseHeldDoc(docId);
            docsHeldInDB.Remove(docId);
            return true;
        }

        internal static void abortSignDoc(int docId)
        {
            try
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("SELECT signedAttId FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
                comm.CommandTimeout = 0;
                comm.Parameters.AddWithValue("@docId", docId);
                conn.Open();
                int signedAttId = (int)comm.ExecuteScalar();
                conn.Close();

                comm = new SqlCommand("DELETE FROM dbo.docnisp WHERE shotef_nisph=@id", conn);
                comm.Parameters.AddWithValue("@id", signedAttId);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();

                comm = new SqlCommand("UPDATE dbo.documents SET signedAttId=@attId, isSigned=@false WHERE shotef_mismach=@docId", conn);
                comm.CommandTimeout = 0;
                comm.Parameters.AddWithValue("@docId", docId);
                comm.Parameters.AddWithValue("@attId", -1);
                comm.Parameters.AddWithValue("@false", false);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }
            catch (Exception e)
            {
                saveLogError("publicFuncsNvars", e.ToString(), e.Message);
            }
        }

        internal static bool inUsers(int id)
        {
            foreach (User u in users)
                if (u.userCode == id)
                    return true;
            return false;
        }

        internal static void beginToPublishDoc(int docId, string subject)
        {

            Dictionary<int, Tuple<string, bool>> toPublish = new Dictionary<int, Tuple<string, bool>>();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT signedAttId FROM dbo.documents (nolock) WHERE (shotef_mismach = @docId)", conn);
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@docId", docId);
            conn.Open();
            int attId = (int)comm.ExecuteScalar();
            conn.Close();
            toPublish.Add(attId, new Tuple<string, bool>(subject + ".pdf", false));
            publishDoc(createDoc(docId), false, toPublish, false);
            MessageBox.Show("המסמך הופץ בהצלחה.", "חתימה", MessageBoxButtons.OK, MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
        }






        internal static void publishDoc(Document doc, bool publishOriginal, Dictionary<int, Tuple<string, bool>> toPublish, bool convertOriginal)
        {
            List<Recipient> civilionRecs = new List<Recipient>();
            string to = "";
            string cc = "";
            string bcc = "";
            List<Tuple<byte[], string, bool>> attachments = new List<Tuple<byte[], string, bool>>();
            User sender = getUserByCode(doc.getSenderRole());// SENDER = kod_sholeah החותם
            User creator = getUserByCode(doc.getCreatorCode());//CREATOR = user_metaiek יוצר המסמך
            string email = sender == null ? null : sender.email;
            string from = curUser.email + "(" + curUser.job + " - " + curUser.getFullName() + ")";//יערה שינתה 19.7 שולח אוטומטית מהמשתמש הפעיל
            string classification = getClassificationByEnum(doc.getClassification());
            string mailSubject = "הפצה - " + doc.getSubject() + " - שוטף:  " + doc.getID().ToString() + "@" + classification + "@";

            bool allRecAreMantak = true;
            cc += sender.email + ";"; //יערה שינתה 19.07.23  - חותם יקבל מייל כעותק (לידיעה) ולא לפעולה
            if (creator != null)
            {
                if (creator.email != null)
                    if (!from.Split(';')[0].Equals(creator.email))
                        cc += creator.email + ";";
                if (!from.Split(';')[0].Equals(curUser.email) && !curUser.email.Equals(creator.email))
                    bcc += curUser.email;
            }
            else
            {
                if (!from.Split(';')[0].Equals(curUser.email))
                    bcc += curUser.email;
            }
            foreach (Recipient r in doc.getRecipients())
            {
                if (r.getSendMail())
                {
                    bool inPub = r.getRole().StartsWith("ת.פ.");
                    if (isValidEmail(r.getEmail()) || inPub)
                    {
                        if (allRecAreMantak && !inUsers(r.getId()))
                            allRecAreMantak = false;
                        if (inPub)
                        {
                            string maToTemp;
                            bool ifa = r.getIFA();
                            string[] tps = r.getRole().Remove(0, 4).Split(',');
                            foreach (string tp in tps)
                            {
                                if (string.IsNullOrWhiteSpace(tp)) continue;
                                User tpUser = getUserByCode(int.Parse(tp));// יערה שינתה ב-08.23 כדי שת.פ ישלח לא רק לרע"ן
                                maToTemp = tpUser.email;
                                if (ifa)
                                    to += maToTemp + ";";
                                else
                                    cc += maToTemp + ";";
                            }
                        }
                        else
                        {
                            string maToTemp = r.getEmail();
                            if (r.getIFA())
                                to += maToTemp + ";";
                            else
                                cc += maToTemp + ";";
                        }
                    }
                    else
                    {
                        MessageBox.Show("כתובת המייל של " + r.getRole() + " לא תקינה." + Environment.NewLine + "המסמך לא הופץ.", "הפצה",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        return;
                    }
                }
                else if (r.getId() == 99999 && r.getEmail() == "")
                    civilionRecs.Add(r);

            }
            if (to.Length > 0)
                to.Remove(to.Length - 1);
            if (cc.Length > 0)
                cc.Remove(cc.Length - 1);
            string body = "";

            allRecAreMantak = false; // Asaf Mor 7/7/21 , Decided to always send the physical document and not bat file link.
            if (allRecAreMantak)
            {
                string text;
                if (publishOriginal)
                {
                    text = "ECHO OFF" + Environment.NewLine + "SET arg1=%1" + Environment.NewLine
                           + "I:\\users\\Public\\docBats\\openFileLink.exe.lnk " + doc.getID().ToString();
                    if (!File.Exists("I:\\users\\Public\\docBats\\" + doc.getID().ToString() + ".bat"))
                        File.WriteAllText("I:\\users\\Public\\docBats\\" + doc.getID().ToString() + ".bat", text);
                    body += "<a href='I:\\users\\Public\\docBats\\" + doc.getID().ToString() + ".bat'>קישור לקובץ המסמך</a><br /><br />";
                }
                foreach (KeyValuePair<int, Tuple<string, bool>> tp in toPublish)
                {
                    text = "ECHO OFF" + Environment.NewLine + "SET arg1=%1 arg2=%2" + Environment.NewLine
                       + "I:\\users\\Public\\docBats\\openFileLink.exe.lnk " + doc.getID().ToString() + " " + tp.Key.ToString();
                    if (!File.Exists("I:\\users\\Public\\docBats\\" + doc.getID().ToString() + "_" + tp.Key.ToString() + ".bat"))
                        File.WriteAllText("I:\\users\\Public\\docBats\\" + doc.getID().ToString() + "_" + tp.Key.ToString() + ".bat", text);
                    body += "<a href='I:\\users\\Public\\docBats\\" + doc.getID().ToString() + "_" + tp.Key.ToString() + ".bat'>קישור לקובץ נספח: "
                        + tp.Value + "</a><br /><br />";
                }
            }
            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm;
                SqlDataReader sdr;
                string name;
                if (publishOriginal)
                {
                    comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
                    comm.Parameters.AddWithValue("@id", doc.getID());
                    conn.Open();
                    sdr = comm.ExecuteReader();
                    sdr.Read();
                    string fileExt = sdr.GetString(1).Trim();
                    name = doc.getSubject().Trim() + "." + fileExt;
                    attachments.Add(new Tuple<byte[], string, bool>(sdr.GetSqlBinary(0).Value, name, convertOriginal));
                    conn.Close();
                }
                foreach (KeyValuePair<int, Tuple<string, bool>> tp in toPublish)
                {
                    comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.docnisp WHERE shotef_mchtv=@docId AND shotef_nisph=@attId AND datalength(file_data)>0", conn);
                    comm.Parameters.AddWithValue("@docId", doc.getID());
                    comm.Parameters.AddWithValue("@attId", tp.Key);
                    conn.Open();
                    sdr = comm.ExecuteReader();
                    sdr.Read();
                    name = tp.Value.Item1 + "." + sdr.GetString(1).Trim();

                    attachments.Add(new Tuple<byte[], string, bool>(sdr.GetSqlBinary(0).Value, name, tp.Value.Item2));
                    conn.Close();
                }
            }
            body += "המסמך מופץ למכותבים בדואר אלקטרוני על ידי " + PublicFuncsNvars.curUser.getFullName() + ", " + curUser.job;

            PublicFuncsNvars.sendAdvancedMail(from, to, cc, bcc, mailSubject, body, attachments);
            if (civilionRecs.Count > 0)
            {

                string msg = "יש מכותבים אזרחיים אשר לא ניתן לשלוח אותו אליהם במייל, ויש להפיץ אליהם בנפרד:\n\n";
                foreach (Recipient r in civilionRecs)
                    msg += r.getRole() + " - " + (r.getIFA() ? "לפעולה" : "לידיעה") + "\n";


                MessageBox.Show(msg, "מכותבים חיצוניים", MessageBoxButtons.OK, MessageBoxIcon.Information,
    MessageBoxDefaultButton.Button1, MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
            }
            doc.published();
        }

        internal static Document createDoc(int id)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT dbo.documents.mismach_or_kovetz, dbo.documents.hanadon, dbo.documents.is_nichnas," +
                                             " dbo.documents.tarich_hamichtav, dbo.documents.tarich_hazana, dbo.documents.kod_sholeah," +
                                             " dbo.documents.teur_tafkid_sholeah, dbo.documents.simuchin, dbo.documents.kod_sivug_bitchoni," +
                                             " dbo.documents.is_hufatz, dbo.documents.is_pail, dbo.documents.is_rapat, dbo.documents.tarich_hafatza," +
                                             " dbo.documents.hearot, dbo.documents.isSigned, dbo.documents.dateSigned, dbo.tm_nig_mis.typ_rcrd FROM dbo.documents (nolock)" +
                                             " LEFT OUTER JOIN dbo.tm_nig_mis ON" +
                                             " dbo.documents.shotef_mismach = dbo.tm_nig_mis.sn_doc_gnrl WHERE dbo.documents.shotef_mismach=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            bool fileOrDoc = sdr.GetBoolean(0), inOrOut = sdr.GetBoolean(2), isPublished = sdr.GetBoolean(9), isActive = sdr.GetBoolean(10), isRapat = sdr.GetBoolean(11),
                isSigned = sdr.GetBoolean(14);
            DateTime dateSigned = sdr.GetDateTime(15);
            string subject = sdr.GetString(1).Trim(), creationDate = sdr.GetString(3).Trim(), entryDate = sdr.GetString(4).Trim(),
                sender = sdr.GetString(6).Trim(), refferences = sdr.GetString(7).Trim(), publishDate = sdr.GetString(12).Trim(), notes = sdr.GetString(13).Trim();
            int senderUser = sdr.GetInt32(5);
            short classification = sdr.GetInt16(8), docType;
            object o = sdr.GetValue(16);
            string s = o.ToString();
            if (s == "")
                docType = 0;
            else
                docType = sdr.GetInt16(16);
            conn.Close();


            Classification c = PublicFuncsNvars.getClassification(classification);
            List<Recipient> sentTo = new List<Recipient>();
            SqlConnection conn2 = new SqlConnection(Global.ConStr);
            string command2 = (PublicFuncsNvars.isNormalDoc(id) ?
                "SELECT kod_mechutav, msd, tiur_tafkid, is_lepeula, is_lishloh_mail, ktovet_mail FROM dbo.doc_mech WHERE shotef_klali=@id order by msd" :
                "SELECT cod_mcotb_bal_tpkyd, CONVERT(smallint,row), dscr_job, CONVERT(bit,CASE [htyyhsot] WHEN 'פ' THEN 1 ELSE 0 END), is_lshloh_doal, adrs_email FROM" +
                " dbo.tm_nig_mhu WHERE typ_rcrd=(SELECT typ_rcrd FROM dbo.tm_nig_mis WHERE sn_doc_gnrl=@id) AND num_rshomh=(SELECT num_rcrd FROM dbo.tm_nig_mis WHERE sn_doc_gnrl=@id)" +
                " AND num_mctb=(SELECT MAX(num_mctb) FROM dbo.tm_nig_mhu WHERE num_rshomh=(SELECT num_rcrd FROM dbo.tm_nig_mis WHERE sn_doc_gnrl=@id)) order by row");
            SqlCommand comm2 = new SqlCommand(command2, conn2);
            comm2.Parameters.AddWithValue("@id", id);
            conn2.Open();
            SqlDataReader sdr2 = comm2.ExecuteReader();
            while (sdr2.Read())
            {
                Recipient r = new Recipient(sdr2.GetInt32(0), sdr2.GetInt16(1), sdr2.GetString(2).Trim(), sdr2.GetBoolean(3),
                    sdr2.GetString(2).Trim().StartsWith("ת.פ.") ? true : sdr2.GetBoolean(4), sdr2.GetString(5).Trim());
                sentTo.Add(r);
            }
            sdr2.Close();

            List<Folder> folders = new List<Folder>();
            comm2 = new SqlCommand("SELECT mispar_nose, is_rashi FROM dbo.tiukim WHERE shotef_klali=@id", conn2);
            comm2.Parameters.AddWithValue("@id", id);
            sdr2 = comm2.ExecuteReader();
            while (sdr2.Read())
            {
                int directoryID = sdr2.GetInt32(0);
                SqlConnection conn3 = new SqlConnection(Global.ConStr);
                SqlCommand comm3 = new SqlCommand("SELECT shm_mshimh, shm_mkotzr, anp, is_tik_pail, sog_mshimh, ms_archh_shosh, kod_sioog FROM dbo.tm_mesimot WHERE ms_mshimh=@directoryID", conn3);
                comm3.Parameters.AddWithValue("@directoryID", directoryID);
                conn3.Open();
                SqlDataReader sdr3 = comm3.ExecuteReader();
                while (sdr3.Read())
                {
                    if (sdr3.GetDouble(5) != 0)
                        folders.Add(new ShoshFolder(directoryID, sdr3.GetString(0).Trim(), sdr3.GetString(1).Trim(), sdr2.GetBoolean(1),
                            (Branch)(sdr3.GetSqlChars(2).Value[0]), sdr3.GetBoolean(3), (FileType)(sdr3.GetSqlChars(4).Value[0]),
                            (Classification)sdr3.GetInt16(6), (int)sdr3.GetDouble(5)));
                    else
                        folders.Add(new Folder(directoryID, sdr3.GetString(0).Trim(), sdr3.GetString(1).Trim(), sdr2.GetBoolean(1),
                            (Branch)(sdr3.GetSqlChars(2).Value[0]), sdr3.GetBoolean(3), (FileType)(sdr3.GetSqlChars(4).Value[0]),
                            (Classification)sdr3.GetInt16(6)));
                }
                conn3.Close();
            }
            sdr2.Close();

            Dictionary<int, bool> authorizedUsers = new Dictionary<int, bool>();
            comm2 = new SqlCommand("SELECT roleCode, isForEdit FROM MantakDB.dbo.doc_Authorizations WHERE docId=@docId", conn2);
            comm2.Parameters.AddWithValue("@docId", id);
            sdr2 = comm2.ExecuteReader();
            while (sdr2.Read())
            {
                int userCode = sdr2.GetInt32(0);
                bool isForEdit = sdr2.GetBoolean(1);
                authorizedUsers.Add(userCode, isForEdit);
            }
            conn2.Close();

            Document d = new Document(id, fileOrDoc, subject, inOrOut, creationDate, entryDate, publishDate, senderUser, sender,
                refferences, c, isPublished, isActive, isRapat, sentTo, folders, authorizedUsers, notes, (DocType)docType, isSigned, dateSigned);
            return d;
        }

        internal static bool alreadySigned(int docId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT isSigned FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@docId", docId);
            conn.Open();
            bool isSigned = (bool)comm.ExecuteScalar();
            conn.Close();
            return isSigned;
        }

        internal static Image getSignatureImage(int userCode)
        {
            Image i = null;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT signature FROM dbo.tmtafkidu WHERE kod_tpkid=@userCode AND signature IS NOT NULL", conn);
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@userCode", userCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                SqlBytes s = sdr.GetSqlBytes(0);
                try
                {
                    i = (Image)(new ImageConverter()).ConvertFrom(s.Buffer);
                }
                catch (Exception e)
                {

                }
            }
            conn.Close();
            return i;
        }

        internal static bool docHasRecipients(int docId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT COUNT(msd) FROM dbo.doc_mech WHERE shotef_klali=@id", conn);
            comm.Parameters.AddWithValue("@id", docId);
            conn.Open();
            return (int)comm.ExecuteScalar() > 0;
        }

        internal static bool docExists(int existingId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT count(shotef_mismach) FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", existingId);
            conn.Open();
            int count = (int)comm.ExecuteScalar();
            if (count > 0)
                return true;
            return false;
        }

        
        internal static void updateCustomPropertiesInWordDoc(Word.Document doc, string newRefferences, string propertyName,bool saveAfterUpdate = true)//Ahava 10.01.2024 update a custom properties of the word document acording to the name of te custom properties.
        {
            try
            {
                dynamic customProperties = doc.CustomDocumentProperties;
                dynamic existingProperty = customProperties[propertyName];
                existingProperty.Value = newRefferences;
            }
            catch
            {}
            doc.Fields.Update();
            if (saveAfterUpdate)
            doc.Save();
        }

        internal static void updateDB(int id, bool updateTime)//Ahava 10.01.2024 update information in the Data Base according to the id.
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand();
            if (updateTime)
            {
                comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=@userCode, OpenedForEditDTime=@currentTime WHERE shotef_mismach=@id", conn);
            }

            else
            {
                comm = new SqlCommand("UPDATE dbo.documents SET whoOpenedForEdit=@userCode WHERE shotef_mismach=@id", conn);
            }
            
            comm.CommandTimeout = 0;
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@userCode", curUser.userCode);
            if (updateTime)
            {
                comm.Parameters.AddWithValue("@currentTime", DateTime.Now);
            }
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            docsHeldInDB.Add(id);
        }
        internal static void openViewDoc(int id)//Ahava 10/01/2024 open a file  to edit.
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            docsHeldInDB.Add(id);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();

            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();

                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                if (File.Exists(filePath))
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        //saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח במחשב." + Environment.NewLine + "המסמך לא נחתם.", "חתימה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        return;
                    } 
                }
                if (!File.Exists(filePath))
                    File.WriteAllBytes(filePath, fileData);
                Word.Application wapp;
                try
                {
                    wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wapp = new Word.Application(); 
                }
                Word.Document doc = new Word.Document();
                object missing = Type.Missing;
                try
                {
                    wapp = (Word.Application)Marshal.GetActiveObject($"Word.Application");
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
                bool itsOpenn = false;
                foreach (Word.Document docc in wapp.Documents)
                {
                    if (docc.FullName == filePath)
                    {
                        doc = docc;
                        itsOpenn = true;
                        break;
                    }
                }
                if (!itsOpenn)
                {
                    doc = wapp.Documents.Open(filePath);
                    
                    wapp.Visible = true;
                    wapp.Activate();
                };
                    
                try
                {
                    while (doc.ActiveWindow != null)
                    {
                        Thread.Sleep(1000);
                    }
                }
                catch
                {}
                bool IsDocSaved = false;
                string Text;
                while (!IsDocSaved)
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        Thread.Sleep(2000);
                        fileData = File.ReadAllBytes(filePath);
                        //)גרסאות)
                        Text = docToTxt(doc, filePath);
                        DocumentHandling.SaveVersion(id);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn ,Text);
                        IsDocSaved = true;
                    }
                    catch
                    {
                        
                    }
                }
                
                killProcessByProcID(id);
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                conn.Close();
                if (File.Exists(filePath))
                {
                    fileData = File.ReadAllBytes(filePath);
                    //)גרסאות)
                    Text = docToTxt(doc,filePath);
                    DocumentHandling.SaveVersion(id);
                    saveDocToDB(ref fileData, id, filePath, ref comm, ref conn ,Text);
                    MessageBox.Show("נשמר כי הקובץ היה פתוח");
                }
            }
        }
        internal static void OpenDocForEditAndNotClose(int id)//Ahava 06/08/2024 open a file  to edit without closing.
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            docsHeldInDB.Add(id);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();

            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();

                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                if (File.Exists(filePath))
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        //saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח במחשב." + Environment.NewLine + "המסמך לא נחתם.", "חתימה",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        return;
                    }
                }
                if (!File.Exists(filePath))
                    File.WriteAllBytes(filePath, fileData);
                Word.Application wapp;
                try
                {
                    wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                }
                catch
                {
                    wapp = new Word.Application();
                }
                Word.Document doc = new Word.Document();
                object missing = Type.Missing;
                try
                {
                    wapp = (Word.Application)Marshal.GetActiveObject($"Word.Application");
                }
                catch (Exception error)
                {
                    MessageBox.Show(error.Message);
                }
                bool itsOpenn = false;
                foreach (Word.Document docc in wapp.Documents)
                {
                    if (docc.FullName == filePath)
                    {
                        doc = docc;
                        itsOpenn = true;
                        break;
                    }
                }
                if (!itsOpenn)
                {
                    doc = wapp.Documents.Open(filePath);

                    wapp.Visible = true;
                    wapp.Activate();
                };

                /*try
                {
                    while (doc.ActiveWindow != null)
                    {
                        Thread.Sleep(1000);
                    }
                }
                catch (Exception e )
                { MessageBox.Show(e.Message); }*/
                /*bool IsDocSaved = false;
                string Text;
                while (!IsDocSaved)
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        Thread.Sleep(2000);
                        fileData = File.ReadAllBytes(filePath);
                        //)גרסאות)
                        Text = docToTxt(doc, filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, Text);
                        DocumentHandling.SaveVersion(id);
                        IsDocSaved = true;
                    }
                    catch
                    {

                    }
                }

                killProcessByProcID(id);
                try
                {
                    File.Delete(filePath);
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                conn.Close();
                if (File.Exists(filePath))
                {
                    fileData = File.ReadAllBytes(filePath);
                    //)גרסאות)
                    Text = docToTxt(doc, filePath);
                    saveDocToDB(ref fileData, id, filePath, ref comm, ref conn, Text);
                    DocumentHandling.SaveVersion(id);
                    MessageBox.Show("נשמר כי הקובץ היה פתוח");
                }*/
            }
        }
        internal static int CustomParts(string value, string name,dynamic customDocumentProperties,Word.Document doc, bool saveAfterUpdate=true)
        {
            int parts = value.Length / 255 + (value.Length % 255 != 0 ? 1 : 0);
            
            if (parts > 1)
            {
                for (int i = 1; i < parts; ++i)
                {
                    string part = value.Substring(i * 255, Math.Min(255, value.Length - i * 255));
                    try
                    {
                        updateCustomPropertiesInWordDoc(doc, part, name + "_" + i, saveAfterUpdate);
                    }
                    catch//לא הכנס לפה אף פעם 
                    {
                        customDocumentProperties.add(name + "_" + i, false, 4, part);
                        updateCustomPropertiesInWordDoc(doc, part, name + "_" + i, saveAfterUpdate);
                    }
                      
                }
                if (parts==2)
                    updateCustomPropertiesInWordDoc(doc, " ", name + "_" + 2, saveAfterUpdate);
            }
            else
            {
                updateCustomPropertiesInWordDoc(doc, " ", name + "_" + 1, saveAfterUpdate);
                updateCustomPropertiesInWordDoc(doc, " ", name + "_" + 2, saveAfterUpdate);
            }
            
            return parts;
        }
        /*internal static bool documentUpdate(int id)
        {
            SqlConnection conn = new SqlConnection(conStr);
            docsHeldInDB.Add(id);
            SqlCommand comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();
                string filePath = Program.folderPath + "\\" + id + "." + fileExt;
                bool itscre = false;
                if (!File.Exists(filePath))
                {
                    File.WriteAllBytes(filePath, fileData);
                    itscre = true;
                }
                else
                {
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                        saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                        itscre = true;
                    }
                    catch
                    {
                        return false;
                    }
                }
                if (itscre)
                {
                    Word.Application wapp;
                    try
                    {
                        wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                    }
                    catch
                    {
                        wapp = new Word.Application();
                    }
                    wapp.Visible = false;
                    Word.Document doc = new Word.Document();
                    doc = wapp.Documents.Open(filePath, ReadOnly: false);
                    dynamic customProperties = doc.CustomDocumentProperties;

                    try
                    {
                        dynamic existingProperty = customProperties["סימוכין"];
                        doc.Close();
                        wapp.Visible = true;
                        
                        conn.Close();

                        try
                        {
                            File.Delete(filePath);
                        }
                        catch
                        { }
                        return true;
                    }
                    catch
                    {
                        doc.Close();
                        
                        wapp.Visible = true;
                        conn.Close();
                        killProcessByProcID(id);
                        try
                        {
                            File.Delete(filePath);
                        }
                        catch
                        { }
                        return false;
                    }
                }
            }
            return false;
        }*/




        /*internal static bool IsWAppRunning()
        {
            try
            {
                Word.Application wapp = (Word.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Word.Application");
                return true;
            }
            catch
            {
                return false;
            }
        }*/

        /*internal static void GetBookmarkValue(Word.Document doc, string bookmarkvalue, string customPropertieName)
        {
            try
            {
                if (doc.Bookmarks.Exists(customPropertieName))
                {
                    Word.Bookmark bookmark = doc.Bookmarks[customPropertieName];
                    Word.Range bookmarkRange = bookmark.Range;
                    string bookmarkValue = bookmarkRange.Text;
                    bool propertyexists = false;
                    foreach (dynamic property in doc.CustomDocumentProperties)
                    {
                        if (property.Name == customPropertieName)
                        {
                            propertyexists = true;
                            string nameprop = "DOCPROPERTY " + customPropertieName;
                            doc.Fields.Add(Range: bookmarkRange, Type: Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                            break;
                        }
                    }
                    if (!propertyexists)
                    {
                        doc.CustomDocumentProperties.Add(customPropertieName, false, 4, bookmarkvalue);
                        string nameprop = "DOCPROPERTY " + customPropertieName;
                        doc.Fields.Add(Range: bookmarkRange, Type: Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                    }
                }
                else
                {
                    Word.Range entireRange = doc.Content;
                    Word.Range foundRange = FindTextInRange(entireRange, bookmarkvalue);
                    if (foundRange != null)
                    {

                        bool propertyexists = false;
                        foreach (dynamic property in doc.CustomDocumentProperties)
                        {
                            if (property.Name == customPropertieName)
                            {
                                propertyexists = true;
                                string nameprop = "DOCPROPERTY " + customPropertieName;
                                doc.Fields.Add(Range: foundRange, Type: Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                                break;
                            }
                        }
                        if (!propertyexists)
                        {
                            //doc.CustomDocumentProperties.Add(customPropertieName, false, Word.WdFieldType.wdFieldAutoText, bookmarkvalue);
                            doc.CustomDocumentProperties.Add(customPropertieName, false, 4, bookmarkvalue);
                            string nameprop = "DOCPROPERTY " + customPropertieName;
                            doc.Fields.Add(Range: foundRange, Type: Microsoft.Office.Interop.Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                        }

                    }
                    else
                    {
                        Console.WriteLine("not found");
                    }

                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }*/


        internal static Word.Range FindTextInRange(Word.Range range, string searchText)
        {
            Word.Find find = range.Find;
            find.Text = searchText;
            find.Forward = true;
            find.MatchWholeWord = true;
            bool found = find.Execute();
            return found ? range : null;
        }
        public static bool IniFileExists(string filename)
        {
            if (!File.Exists(filename))
            {
                MessageBox.Show("קובץ" + Global.IniFileName + "לא קיים");
                Application.Exit();
            }
            return true;
        }
        internal static string docToTxt(Word.Document doc,string filePath)
        {
            try
            {
                string text = doc.Content.Text;
                return TxtFromTitle(doc); //doc.Content.Text;
            }
            catch
            {
                try
                {
                    Word.Application wapp;

                    try
                    {
                        wapp = (Word.Application)Marshal.GetActiveObject("Word.Application");
                        wapp.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;
                        wapp.Visible = false;
                        wapp.ScreenUpdating = false;
                        wapp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                    }
                    catch
                    {
                        wapp = new Word.Application();
                    }
                   
                    
                    doc = wapp.Documents.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, false);
                    
                    string text = TxtFromTitle(doc); //doc.Content.Text;
                    doc.Close();
                    return text;
                }
                catch(Exception e)
                {
                    try
                    {
                        Thread.Sleep(1000);
                        doc.Close();
                    }
                    catch (Exception ex)
                    {

                    }
                    return "";
                }
                
            }
        }
        internal static string TxtFromTitle(Word.Document doc)
        {
            string title;
            try
            {
                title = doc.BuiltInDocumentProperties[Word.WdBuiltInProperty.wdPropertyTitle].Value as string;
            }
            catch(Exception e)
            {
                try
                {
                    string text = doc.Content.Text;
                    return ClearTxt(text);
                }
                catch
                {
                    return "";
                }
            }
            if (!string.IsNullOrEmpty(title))
            {
                foreach(Word.Range storyRange in doc.StoryRanges)
                {
                    Word.Range searchrange = storyRange.Duplicate;
                    Word.Find findObject = searchrange.Find;
                    findObject.Text = title;
                    if (findObject.Execute())
                    {
                        Word.Range textRange = doc.Range(searchrange.End, storyRange.End);
                        string textAfterTitle = textRange.Text;
                        return ClearTxt(textAfterTitle);
                    }
                    else
                    {
                        try
                        {
                            string text = doc.Content.Text;
                            return ClearTxt(text);
                        }
                        catch
                        {
                            return "";
                        }
                    }
                }
            }
            else
            {
                try
                {
                    string text = doc.Content.Text;
                    return ClearTxt(text);
                }
                catch
                {
                    return "";
                }
            }

            return "";
        }

        internal static string ClearTxt(string text)
        {
            text = text.Replace("\r", " ").Replace("\v", " ").Replace("\n", " ").Replace("\t", " ").Replace("\a", " ");
            text = text.Replace("(", " ").Replace(")", " ").Replace("[", " ").Replace("]", " ").Replace("{", " ").Replace("}", " ").Replace("<", " ").Replace(">", " ").Replace("/", " ").Replace("\\", " ");
            text = text.Replace("_", " ").Replace("-", " ").Replace("+", " ").Replace(":", " ").Replace(";", " ").Replace(",", " ").Replace(".", " ");
            text = text.Replace("*", "").Replace("~", "").Replace("!", "").Replace("?", "").Replace("|", "").Replace("'", "").Replace("\"", "");
            text = Regex.Replace(text, @"\s+", " ");
            return text;
        }
    }
}
