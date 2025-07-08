using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;

namespace DocumentsModule
{
    public class Document
    {
        int id;//מס' שוטף
        DateTime creationDate;//תאריך יצירת המסמך
        DateTime entryDate;//תאריך הזנת המסמך(לא ממש נמצא בשימוש)
        DateTime publishDate;//תאריך הפצה אחרונה
        string subject;//נדון
        string refferences;//סימוכין
        int senderUser;//קוד משתמש בעלים
        string sender;//שם/תפקיד השולח
        List<Recipient> recipients;//רשימת המכותבים למסמך זה
        Branch branch;//הענף בשמו נוצר המסמך
        internal bool isPublished { get; set; }//האם הופץ
        internal bool isSigned { get; set; }//האם נחתם
        internal DateTime dateSigned { get; set; }//תאריך חתימה
        bool isActive;//האם המסמך פעיל
        Classification classification;//סיווג המסמך
        bool inOrOut;//נכנס או יוצא
        List<Folder> folders;//תיקים בהם מתויק המסמך
        bool docOrFile;//מסמך חדש או קובץ שצורף לשוטף
        bool isRapat;//האם הגיע מרפ"ט
        string notes;//הערות
        private Dictionary<int, bool> authorizedUsers;
        internal DocType type { get; private set; }

        internal Document(int id, bool docOrFile, string subject, bool inOrOut, string creationDate, string entryDate, string publishDate,
            int senderUser, string sender, string refferences, Classification classification, bool isPublished, bool isActive,
            bool isRapat, List<Recipient> sentTo, List<Folder> folders, Dictionary<int, bool> authorizedUsers, string notes, DocType type,
            bool isSigned, DateTime dateSigned)
        {
            this.notes = notes;
            this.id = id;
            this.docOrFile = docOrFile;
            this.subject = subject;
            this.inOrOut = inOrOut;
            this.isSigned = isSigned;
            this.dateSigned = dateSigned;
            try
            {
                this.creationDate = DateTime.ParseExact(creationDate, "yyyyMMdd", new CultureInfo("he-IL"));
            }
            catch(Exception e)
            {
                PublicFuncsNvars.saveLogError("Document", e.ToString(), e.Message);
                this.creationDate = DateTime.MinValue;
            }
            try
            {
                this.entryDate = DateTime.ParseExact(entryDate, "yyyyMMdd", new CultureInfo("he-IL"));
            }
            catch(Exception e)
            {
                PublicFuncsNvars.saveLogError("Document", e.ToString(), e.Message);
                this.entryDate = DateTime.MinValue;
            }
            if (isPublished)
            {
                this.publishDate = DateTime.ParseExact(publishDate, "yyyyMMdd", new CultureInfo("he-IL"));
            }
            this.senderUser = senderUser;
            this.sender = sender;
            this.refferences = refferences;
            this.classification = classification;
            this.isPublished = isPublished;
            this.isActive = isActive;
            this.isRapat = isRapat;

            if (sentTo != null)
                this.recipients = sentTo;
            else
                this.recipients = new List<Recipient>();

            if (authorizedUsers != null)
                this.authorizedUsers = authorizedUsers;
            else
                this.authorizedUsers = new Dictionary<int, bool>();
            this.folders = folders;
            this.type = type;
        }

        public int getID()
        {
            return id;
        }

        public List<Recipient> getRecipients()
        {
            return recipients;
        }

        public List<Folder> getFolders()
        {
            return folders;
        }

        public Branch getBranch()
        {
            return branch;
        }

        internal string getSubject()
        {
            return subject;
        }

        internal string getSenderName()
        {
            return sender;
        }

        internal string getRefferences()
        {
            return refferences;
        }

        internal Classification getClassification()
        {
            return classification;
        }

        internal bool getInOrOut()
        {
            return inOrOut;
        }

        internal bool getIsRapat()
        {
            return isRapat;
        }

        internal bool getIsActive()
        {
            return isActive;
        }

        internal DateTime getCreationDate()
        {
            return creationDate;
        }

        internal void removeAllRecipients()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.doc_mech WHERE shotef_klali=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            recipients.Clear();
        }

         internal bool addRecipient(Recipient r)
        {
            if (recipientsContains(r))
                return false;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("INSERT INTO dbo.doc_mech(kod_marechet, shotef_klali, msd, kod_mechutav, tiur_tafkid, is_lepeula,"
                +" is_ishu_kabala, is_lishloh_mail, ktovet_mail) VALUES(2, @docId, @msd, @id, @roleDescr, @isForAct, 0, @sendMail, @email)", conn);
            comm.Parameters.AddWithValue("@docId", id);
            r.setNID((short)(recipients.Count>0?(recipients[recipients.Count - 1].getNID() + 1):1));
            comm.Parameters.AddWithValue("@msd", r.getNID());
            comm.Parameters.AddWithValue("@id", r.getId());
            comm.Parameters.AddWithValue("@roleDescr", r.getRole());
            comm.Parameters.AddWithValue("@isForAct", r.getIFA());
            comm.Parameters.AddWithValue("@sendMail", r.getSendMail());
            comm.Parameters.AddWithValue("@email", r.getEmail());
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            recipients.Add(r);
            return true;
        }

        private bool recipientsContains(Recipient r)
        {
            foreach(Recipient rec in recipients)
            {
                if ((r.getId() != 99999 && r.getId() == rec.getId()) || (r.getId() == 99999 && r.getRole() == rec.getRole()))
                    return true;
            }
            return false;
        }

        internal short getMaxRecipient()
        {
            if (recipients.Count > 0)
                return recipients[recipients.Count - 1].getNID();
            return 0;
        }

        internal int getSenderRole()
        {
            return senderUser;
        }

        internal List<int> getAttsIds()
        {
            List<int> atts = new List<int>();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT shotef_nisph FROM dbo.docnisp WHERE shotef_mchtv=@docId AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@docId", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            while(sdr.Read())
                atts.Add(sdr.GetInt32(0));
            conn.Close();
            return atts;
        }

        internal bool getIsRagish()
        {
            bool isRagish = false;
            SqlConnection conn2 = new SqlConnection(Global.ConStr);
            SqlCommand comm2 = new SqlCommand("SELECT isRagish from dbo.documents (nolock) WHERE shotef_mismach=@shotef", conn2);
            comm2.Parameters.AddWithValue("@shotef", id);
            conn2.Open();
            SqlDataReader sdr2 = comm2.ExecuteReader();

            while (sdr2.Read())
            {
                isRagish = sdr2.GetBoolean(0);
            }

            return isRagish;
        }

        internal int getNumInFolder(int FolderId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT mispar_in_tik FROM dbo.tiukim WHERE mispar_nose=@id AND shotef_klali=@docId", conn);
            comm.Parameters.AddWithValue("@id", FolderId);
            comm.Parameters.AddWithValue("@docId", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            int n = sdr.GetInt32(0);
            conn.Close();
            return n;
        }

        internal int getProject()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT msd_proiect FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            sdr.Read();
            int n = sdr.GetInt16(0);
            conn.Close();
            return n;
        }

        internal DateTime getEntryDate()
        {
            return entryDate;
        }

        internal bool getIsPublished()
        {
            return isPublished;
        }

        
        internal void setPublishDate(DateTime publishDate , bool isPublished)
        {
            if (isPublished)
            {
                string date = publishDate.ToString("yyyyMMdd");
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET is_hufatz=@true, tarich_hafatza=@date WHERE shotef_mismach=@docId", conn);
                comm.Parameters.AddWithValue("@docId", id);
                comm.Parameters.AddWithValue("@true", isPublished);
                comm.Parameters.AddWithValue("@date", date);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }

            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET is_hufatz=@true WHERE shotef_mismach=@docId", conn);
                comm.Parameters.AddWithValue("@docId", id);
                comm.Parameters.AddWithValue("@true", isPublished);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }

            
        }
        internal DateTime getPublishDate()
        {
            return publishDate;
        }

        internal string getNotes()
        {
            return notes;
        }

        internal void published()
        {
            string date = DateTime.Today.ToString("yyyyMMdd");
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET is_hufatz=@true, tarich_hafatza=@date WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            comm.Parameters.AddWithValue("@true", true);
            comm.Parameters.AddWithValue("@date",date);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
        }
                
        internal void removeRecipient(short nid)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.doc_mech WHERE shotef_klali=@id AND msd=@nid", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@nid", nid);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();

            // updating msd (running number)
            /*SqlCommand comm = new SqlCommand("Uodate running number command", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@nid", nid);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();*/

            foreach (Recipient r in recipients)
            {
                if (r.getNID() == nid)
                {
                    recipients.Remove(r);
                    break;
                }
            }
        }

        internal Folder changeMainFolder(int folderId)
        {
            Folder mainFolder= getMainFolder();
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.tiukim SET is_rashi=0 WHERE mispar_nose=@folderId AND shotef_klali=@id", conn);
            if (mainFolder != null)
            {
                comm.Parameters.AddWithValue("@id", id);
                comm.Parameters.AddWithValue("@folderId", mainFolder.id);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
                mainFolder.isMain = false;
            }


            foreach(Folder folder in folders)
            {
                if(folder.id==folderId)
                {
                    comm = new SqlCommand("UPDATE dbo.tiukim SET is_rashi=1 WHERE mispar_nose=@folderId AND shotef_klali=@id", conn);
                    comm.Parameters.AddWithValue("@id", id);
                    comm.Parameters.AddWithValue("@folderId", folderId);
                    conn.Open();
                    comm.ExecuteNonQuery();
                    conn.Close();
                    folder.isMain = true;
                    break;
                }
            }

            Folder f = getMainFolder();
            comm = new SqlCommand("SELECT mispar_in_tik FROM dbo.tiukim WHERE mispar_nose=@folderId AND shotef_klali=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@folderId", folderId);
            conn.Open();
            var numberInFolder = comm.ExecuteScalar();
            conn.Close();
            string newRefferences=f.shortDescription + " - " + numberInFolder + " - " + id;

            comm = new SqlCommand("UPDATE dbo.documents SET simuchin=@refferences WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            comm.Parameters.AddWithValue("@refferences", newRefferences);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (sdr.Read())
            {
                PublicFuncsNvars.updateRefferencesInWordDoc(id, refferences, newRefferences);
                refferences = newRefferences;
            }
            return mainFolder;
        }

        private Folder getMainFolder()
        {
            foreach (Folder folder in folders)
                if (folder.isMain)
                    return folder;
            return null;
        }

        internal bool isFiledInFolder(int id)
        {
            foreach (Folder folder in folders)
                if (folder.id==id)
                    return true;
            return false;
        }

        internal bool isRegularDoc()
        {
            return type == DocType.normal && !docOrFile;
        }

        internal void addFolder(Folder folder)
        {
            folders.Add(new Folder(folder.id, folder.description, folder.shortDescription, false, folder.branch, folder.isActive, folder.type, folder.classification));
        }

        /*
         * updates the is for act status of the nid recipient in the document
         */
        internal void updateRecipientIFAByNid(int nid, bool value)
        {
            recipients.Where(x => x.getNID() == nid).ToList()[0].setIFA(value);
        }

        /*internal void updateDetails(string subject, short classification, string notes, bool isActive, int newUser,bool isRagish)
        {
            string oldSubject = this.subject;
            Classification oldClassification = this.classification;
            this.subject = subject;
            this.classification = (Classification)classification;
            this.notes = notes;
            this.isActive = isActive;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET hearot=@notes, hanadon=@subject, kod_sivug_bitchoni=@classification, is_pail=@isActive, isRagish=@isRagish WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            comm.Parameters.AddWithValue("@notes", notes);
            comm.Parameters.AddWithValue("@subject", subject);
            comm.Parameters.AddWithValue("@classification", classification);
            comm.Parameters.AddWithValue("@isActive", isActive);
            comm.Parameters.AddWithValue("@isRagish", isRagish);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();

            if(subject!=oldSubject || oldClassification!=(Classification)classification)
            {
                PublicFuncsNvars.updateSubjectAndClassificationInWordDoc(id, oldSubject, subject, oldClassification, (Classification)classification);
            }

            if(newUser!=senderUser && PublicFuncsNvars.users.Any(x=>x.userCode==newUser))
            {
                PublicFuncsNvars.updateSignatureInWordDoc(id, newUser);
                comm = new SqlCommand("UPDATE dbo.documents SET kod_sholeah=@newUser, teur_tafkid_sholeah=@userRole WHERE shotef_mismach=@docId", conn);
                comm.Parameters.AddWithValue("@docId", id);
                comm.Parameters.AddWithValue("@newUser", newUser);
                comm.Parameters.AddWithValue("@userRole", PublicFuncsNvars.users.Find(x => x.userCode == newUser).job);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
                this.senderUser = newUser;0
                this.sender = PublicFuncsNvars.users.Find(x => x.userCode == newUser).job;
            }
        }*/

        internal bool NewUpdateDetails(string newSubject, string oldSubject, short oldClassification, short newClassification,int oldUser, int newUser, string notes, bool isActive, bool isRagish, Document document,bool hasBeenUpdated)
        {
           
            this.subject = newSubject;
            this.classification = (Classification)newClassification;
            this.notes = notes;
            this.isActive = isActive;

            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.documents SET hearot=@notes, hanadon=@subject, kod_sivug_bitchoni=@classification, is_pail=@isActive, isRagish=@isRagish WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            comm.Parameters.AddWithValue("@notes", notes);
            comm.Parameters.AddWithValue("@subject", newSubject);
            comm.Parameters.AddWithValue("@classification", newClassification);
            comm.Parameters.AddWithValue("@isActive", isActive);
            comm.Parameters.AddWithValue("@isRagish", isRagish);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();

            conn = new SqlConnection(Global.ConStr);
            PublicFuncsNvars.updateDB(id, false);
            comm = new SqlCommand("SELECT file_data, file_extension FROM dbo.documents (nolock) WHERE shotef_mismach=@id AND datalength(file_data)>0", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            
            if (sdr.Read())
            {
                byte[] fileData = sdr.GetSqlBytes(0).Buffer;
                string fileExt = sdr.GetString(1).Trim();
                conn.Close();
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
                        //PublicFuncsNvars.saveDocToDB(ref fileData, id, filePath, ref comm, ref conn);
                        itscreated = true;
                    }
                    catch
                    {
                        MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את המסמך.");
                        MessageBox.Show("המסמך פתוח לכן לא ניתן לעדכן את המסמך." + Environment.NewLine + "יש לסגור את המסמך ואז לעדכן פרטים.", "עדכון פרטים",
                    MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                    MessageBoxOptions.RtlReading | MessageBoxOptions.RightAlign);
                        return false;
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
                        bool iswAppVisible = wApp.Visible;
                        if (iswAppVisible)
                            wApp.Visible = false;
                        Word.Document doc = wApp.Documents.Open(filePath);
                        bool newVersion = false;
                        dynamic customProperties = doc.CustomDocumentProperties;
                        try
                        {
                            dynamic existingProperty = customProperties["סימוכין"];
                            newVersion = true;
                        }
                        catch { }
                        
                        if (oldSubject != newSubject)
                            //PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, newSubject, "נדון");
                            doc.BuiltInDocumentProperties["Title"].Value = newSubject;//שיניתי
                        if (oldClassification != newClassification)
                        {
                            doc.BuiltInDocumentProperties["Category"].Value = PublicFuncsNvars.getClassificationByEnum(PublicFuncsNvars.getClassification(newClassification));
                            if (!newVersion)
                                PublicFuncsNvars.replaceTextInHeaderFooter(doc, PublicFuncsNvars.getClassificationByEnum(PublicFuncsNvars.getClassification(oldClassification)), PublicFuncsNvars.getClassificationByEnum(PublicFuncsNvars.getClassification(newClassification)));
                        }

                        if (newUser != oldUser && PublicFuncsNvars.users.Any(x => x.userCode == newUser))
                        {
                            SqlConnection conn1 = new SqlConnection(Global.ConStr);
                            SqlCommand comm1 = new SqlCommand("UPDATE dbo.documents SET kod_sholeah=@newUser, teur_tafkid_sholeah=@userRole WHERE shotef_mismach=@docId", conn1);
                            comm1.Parameters.AddWithValue("@docId", id);
                            comm1.Parameters.AddWithValue("@newUser", newUser);
                            comm1.Parameters.AddWithValue("@userRole", PublicFuncsNvars.users.Find(x => x.userCode == newUser).job);
                            conn1.Open();
                            comm1.ExecuteNonQuery();
                            conn1.Close();
                            this.senderUser = newUser;
                            this.sender = PublicFuncsNvars.users.Find(x => x.userCode == newUser).job;
                            
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
                            if (newVersion)
                            {
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, linesNew[0], "חתימה_שורה_א",false);
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, linesNew[1], "חתימה_שורה_ב", false);
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, linesNew[2], "חתימה_שורה_ג", false);
                            }
                            
                        }
                        if (true) //!hasBeenUpdated)
                        {
                            string forAct = "", forKnow = "";
                            int lengt = document.getRecipients().Count;
                            foreach (Recipient r in document.getRecipients())
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

                            string documentText = doc.Content.Text;
                            bool textExists = documentText.Contains("רשימת תפוצה");
                            int partsAct = PublicFuncsNvars.CustomParts(forAct, "נמענים_לפעולה", customProperties, doc,false);
                            int partsKnow = PublicFuncsNvars.CustomParts(forKnow, "נמענים_לידיעה", customProperties, doc,false);
                            if (forAct != "")
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, forAct, "נמענים_לפעולה",false);
                            else
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, " ", "נמענים_לפעולה", false);

                            if (forKnow != "")
                                //doc.BuiltInDocumentProperties["Title"].Value = forKnow;
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, forKnow, "נמענים_לידיעה", false);
                            else
                                //doc.BuiltInDocumentProperties["Title"].Value = " ";
                                PublicFuncsNvars.updateCustomPropertiesInWordDoc(doc, " ", "נמענים_לידיעה", false);

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
                                    object missing = Type.Missing;
                                    object unit = Word.WdUnits.wdStory;
                                    object extend = Word.WdMovementType.wdMove;
                                    wApp.Selection.EndKey(ref unit, ref extend);
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
                                        Word.Range entireRange = doc.Content;
                                        Word.Range lr = PublicFuncsNvars.FindTextInRange(entireRange, "cell1");
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
                                        {//להוסיף
                                            doc.CustomDocumentProperties.Add("נמענים_לידיעה", false, 4, forKnow);
                                        }
                                        nameprop = "DOCPROPERTY " + "נמענים_לידיעה";
                                        entireRange = doc.Content;
                                        lr = PublicFuncsNvars.FindTextInRange(entireRange, "cell2");
                                        if (lr != null)
                                            doc.Fields.Add(Range: lr, Type: Word.WdFieldType.wdFieldEmpty, Text: nameprop, PreserveFormatting: true);
                                    }
                                }
                            }
                        }
                        doc.Fields.Update();
                        doc.Save();
                        string Text =PublicFuncsNvars.docToTxt(doc,filePath); //1
                        doc.Close();
                        if(iswAppVisible)
                            wApp.Visible = true;
                        wApp.Quit(); // 23.9.24 - ASAF MOR
                        Marshal.ReleaseComObject(doc);
                        fileData = File.ReadAllBytes(filePath);
                        //לפני שמירה לDB )גרסאות)

                        DocumentHandling.SaveVersion(id);
                        PublicFuncsNvars.saveDocToDB(ref fileData, id, filePath, ref comm, ref conn,Text);
                        
                        if (!newVersion)
                        {
                            MessageBox.Show("המסמך בתבנית ישנה." + Environment.NewLine + "יש לערוך ידנית את המסמך", "מסמך ישן", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1,
                            MessageBoxOptions.RightAlign | MessageBoxOptions.RtlReading);
                            return false;
                        }
                    }
                    catch (Exception e)
                    {
                        PublicFuncsNvars.saveLogError("publicFuncsNvars", e.ToString(), e.Message);
                    }
                    try
                    {
                        fileData = File.ReadAllBytes(filePath);
                    }
                    catch
                    {
                    }
                    
                }
            }
            
            PublicFuncsNvars.releaseHeldDoc(id);
            PublicFuncsNvars.docsHeldInDB.Remove(id);
            return true;
        }

        internal bool addAuthorization(int userCode, bool isForEdit)
        {
            if (authorizedUsers.Keys.Contains(userCode))
                return false;
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("INSERT INTO dbo.doc_Authorizations(docId, roleCode, isForEdit) VALUES (@docId, @userCode, @isForEdit)", conn);
            comm.Parameters.AddWithValue("@docId", id);
            comm.Parameters.AddWithValue("@userCode", userCode);
            comm.Parameters.AddWithValue("@isForEdit", isForEdit);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            authorizedUsers.Add(userCode, isForEdit);
            return true;
        }

        internal void removeAllAuthorizations()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.doc_Authorizations WHERE docId=@id", conn);
            comm.Parameters.AddWithValue("@id", id);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            authorizedUsers.Clear();
        }

        internal void updateAuthorizationIFEByCode(int userCode, bool isForEdit)
        {
            authorizedUsers[userCode] = isForEdit;
        }

        internal void removeAuthorization(int userCode)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.doc_Authorizations WHERE docId=@id AND roleCode=@userCode", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@userCode", userCode);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            authorizedUsers.Remove(userCode);
        }

        internal void removeAttachment(int attId)
        {
            if (attId == this.getSignedAtt())
            {
                PublicFuncsNvars.abortSignDoc(id);
            }
            else
            {
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("DELETE FROM dbo.docnisp WHERE shotef_mchtv=@id AND shotef_nisph=@attId", conn);
                comm.Parameters.AddWithValue("@id", id);
                comm.Parameters.AddWithValue("@attId", attId);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }
        }

        internal void removeFolder(int folderId)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.tiukim WHERE shotef_klali=@id AND mispar_nose=@folderId", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@folderId", folderId);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            folders.Remove(folders.Find(x => x.id == folderId));
        }

        internal Dictionary<int,bool> getAuthorizedUsers()
        {
            return authorizedUsers;
        }

        internal bool isCurUserAuthorizedToEdit()
        {
            if(authorizedUsers.ContainsKey(PublicFuncsNvars.curUser.userCode))
                return authorizedUsers[PublicFuncsNvars.curUser.userCode];
            return authorizedUsers.Count == 0;
        }

        internal Recipient getRecipient(short nid)
        {
            foreach (Recipient r in recipients)
            {
                if (r.getNID() == nid)
                {
                    return r;
                }
            }
            return null;
        }

        internal bool moveRecipient(short toMove, short newLocation)
        {
            Recipient r = getRecipient((short)toMove);
            int index = recipients.IndexOf(r);

            Recipient r2 = getRecipient((short)newLocation);
            int index2 = recipients.IndexOf(r2);

            if(index<index2)
            {
                for(int i=index;i<index2;i++)
                {
                    moveRecipientDown(recipients[i].getNID());
                }
                return true;
            }
            else if(index>index2)
            {
                for (int i = index; i > index2; i--)
                {
                    moveRecipientUp(recipients[i].getNID());
                }
                return true;
            }
            return false;
        }

        internal bool moveRecipientUp(int p)
        {
            Recipient r = getRecipient((short)p);
            int index = recipients.IndexOf(r);
            if(index>0)
            {
                recipients.Remove(r);
                short nid=r.getNID();
                r.setNID(recipients[index-1].getNID());
                recipients.Insert(index - 1, r);
                recipients[index].setNID(nid);
                switch2Recipients(r, index);
                return true;
            }
            return false;
        }

        internal bool moveRecipientDown(int p)
        {
            Recipient r = getRecipient((short)p);
            int index = recipients.IndexOf(r);
            if (index < recipients.Count-1)
            {
                recipients.Remove(r);
                short nid = r.getNID();
                r.setNID(recipients[index].getNID());
                recipients.Insert(index+1, r);
                recipients[index].setNID(nid);
                switch2Recipients(r, index);
                return true;
            }
            return false;
        }

        void switch2Recipients(Recipient r, int index)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("UPDATE dbo.doc_mech SET msd=-1 WHERE shotef_klali=@id AND msd=@nid", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@nid", r.getNID());
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            comm = new SqlCommand("UPDATE dbo.doc_mech SET msd=@newnid WHERE shotef_klali=@id AND msd=@nid", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@nid", recipients[index].getNID());
            comm.Parameters.AddWithValue("@newnid", r.getNID());
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
            comm = new SqlCommand("UPDATE dbo.doc_mech SET msd=@newnid WHERE shotef_klali=@id AND msd=@nid", conn);
            comm.Parameters.AddWithValue("@id", id);
            comm.Parameters.AddWithValue("@nid", -1);
            comm.Parameters.AddWithValue("@newnid", recipients[index].getNID());
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
        }

        internal void updateRecipientRole(short nid, string newRole)
        {
            Recipient r = getRecipient(nid);
            if(r.getRole()!=newRole)
            {
                r.setRole(newRole);
                SqlConnection conn = new SqlConnection(Global.ConStr);
                SqlCommand comm = new SqlCommand("UPDATE dbo.doc_mech SET tiur_tafkid=@newRole WHERE shotef_klali=@id AND msd=@nid", conn);
                comm.Parameters.AddWithValue("@id", id);
                comm.Parameters.AddWithValue("@nid", r.getNID());
                comm.Parameters.AddWithValue("@newRole", newRole);
                conn.Open();
                comm.ExecuteNonQuery();
                conn.Close();
            }
        }

        internal int getCreatorCode()
        {
            MantakDBDataSetDocumentsTableAdapters.documentsTableAdapter dta = new MantakDBDataSetDocumentsTableAdapters.documentsTableAdapter();
            string c = dta.GetDataById(id).First().user_metaiek;
            int creator;
            if(int.TryParse(c, out creator))
            {
                if (PublicFuncsNvars.users.Any(x => x.userCode == creator))
                    return creator;
            }
            return -1;
        }

        internal int getSignedAtt()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT signedAttId FROM dbo.documents (nolock) WHERE shotef_mismach=@docId", conn);
            comm.Parameters.AddWithValue("@docId", id);
            conn.Open();
            int attId = (int)comm.ExecuteScalar();
            conn.Close();
            return attId;
        }
    }
}
