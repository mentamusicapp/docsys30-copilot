using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    class User
    {
        internal int userCode { get; set; }
        internal int magicUser { get; set; }
        internal string email { get; set; }
        internal string userLogin { get; set; }
        internal string job { get; set; }
        internal Branch branch { get; set; }
        internal Branch permissionsBranch { get; set; }
        internal int commanderCode { get; set; }
        internal RoleType roleType { get; set; }
        internal string lastName { get; set; }
        internal string firstName { get; set; }
        internal bool isActive { get; set; }
        public bool allowedToOpenFolders { get; set; }

        public User(int uc, int mu, string fn, string ln, string em, string lg, string j, char b, short pb, int cc, short rt, bool ia, bool atof)
        {
            userCode = uc;
            magicUser = mu;
            firstName = fn;
            lastName = ln;
            email = em;
            userLogin = lg;
            job = j;
            commanderCode = cc;
            roleType = (RoleType)rt;
            isActive = ia;
            allowedToOpenFolders = atof;

            try
            {
                branch = (Branch)b;
            }
            catch(Exception e)
            {
                PublicFuncsNvars.saveLogError("User", e.ToString(), e.Message);
                branch = Branch.other;
            }

            try
            {
                permissionsBranch = (Branch)char.Parse(pb.ToString());
            }
            catch (Exception e)
            {
                PublicFuncsNvars.saveLogError("User", e.ToString(), e.Message);
                permissionsBranch = branch;
            }
        }

        internal string getFullName()
        {
            return firstName + " " + lastName;
        }

        internal bool addAuthorization(int userCode, bool isForEdit)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT userCode FROM dbo.doc_AutoAuthorizations WHERE userCode=@userCode AND authorizedUserCode=@authorizedUserCode", conn);
            comm.Parameters.AddWithValue("@authorizedUserCode", userCode);
            comm.Parameters.AddWithValue("@userCode", this.userCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            if (!sdr.Read())
            {
                sdr.Close();
                comm.CommandText = "INSERT INTO dbo.doc_AutoAuthorizations(userCode, authorizedUserCode, isForEdit) VALUES (@userCode, @authorizedUserCode, @isForEdit)";
                comm.Parameters.AddWithValue("@isForEdit", isForEdit);
                comm.ExecuteNonQuery();
                conn.Close();
                return true;
            }
            else
            {
                conn.Close();
                return false;
            }
        }

        internal Dictionary<int, bool> getAutoAuthorizedUsers()
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("SELECT authorizedUserCode, isForEdit FROM dbo.doc_AutoAuthorizations WHERE userCode=@userCode", conn);
            comm.Parameters.AddWithValue("@userCode", userCode);
            conn.Open();
            SqlDataReader sdr = comm.ExecuteReader();
            Dictionary<int, bool> authorizedUsers = new Dictionary<int, bool>();
            while(sdr.Read())
            {
                authorizedUsers.Add(sdr.GetInt32(0), sdr.GetBoolean(1));
            }
            conn.Close();
            return authorizedUsers;
        }

        internal void removeAuthorization(int authorizedUserCode)
        {
            SqlConnection conn = new SqlConnection(Global.ConStr);
            SqlCommand comm = new SqlCommand("DELETE FROM dbo.doc_AutoAuthorizations WHERE userCode=@userCode AND authorizedUserCode=@authorizedUserCode", conn);
            comm.Parameters.AddWithValue("@userCode", userCode);
            comm.Parameters.AddWithValue("@authorizedUserCode", authorizedUserCode);
            conn.Open();
            comm.ExecuteNonQuery();
            conn.Close();
        }
    }
}
