using System;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Collections.Generic;
using System.Reflection;


public class Global
{
    public static string Version = " - Ver 21.06 - 26/06/2025";
    //public static string IconFile;//= Application.StartupPath + "\\MantakDocuments.ico";
    public static Icon AppIcon=Icon.ExtractAssociatedIcon(Assembly.GetExecutingAssembly().Location);// = new Icon(IconFile);
    public static string AppFileName = System.AppDomain.CurrentDomain.FriendlyName.Replace(".exe", "");
    public static string IniFileName = Application.StartupPath + "\\" + "DocumentsModule.INI";//AppFileName.Split('.')[0] + ".INI";
    public static bool IniFileExists = File.Exists(IniFileName);
    
    //public static Dictionary<string, string> INIvalues=File.ReadLines(IniFileName).Where(line => (!String.IsNullOrWhiteSpace(line) && !line.StartsWith("#"))).Select(line =>line.Split(new char[] { '=' },2,0)).ToDictionary(parts => parts[0].Trim(),parts => parts.Length >1 ? parts [1].Trim() : null);
    public static Dictionary<string, string> INIvalues;
    public static string P_SQL_SRV;// = INIvalues["SQL_SRV"];
    public static string P_SQL_DB;// = INIvalues["SQL_DB"];
    public static string P_MAX_CLASS;// = INIvalues["MAX_CLASS"];
    public static string P_SQL_USR;// = INIvalues["SQL_URL"];
    public static string P_SQL_PSW;// = INIvalues["SQL_PSW"];
    public static string P_APP;
    public static string P_LCL;
    //public static string ConStr = "Persist Security Info=False;User ID=" + P_SQL_USR + ";Password=" + "; Initial Catalog =" + P_SQL_DB + ":Server=" + P_SQL_PSW;
    //public static string ConStr = "Data Source=" + P_SQL_SRV + ";Initial Catalog=" + P_SQL_DB + ";Integrated Security=True";
    public static string ConStr = "Server=185.145.252.75,24412;Database=MantakDB;User ID=MantakApp;Password=MantakApp";
    public static string Key = "BekolDarKehaDaeu";
    public static int SecondToCloseDocForm = 900000; // 90,000 ms = 15 minutes.
    public static string ReadOnly;
}