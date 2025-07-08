using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    public class Recipient
    {
        int code;
        short numInDoc;
        string role;
        bool isForAct;
        bool sendMail;
        string emailAddress;

        public Recipient(int code, short nid, string name, bool ifa, bool sendMail, string email)
        {
            this.code = code;
            this.numInDoc = nid;
            this.role = name;
            this.isForAct = ifa;
            this.sendMail = sendMail;
            this.emailAddress = email;
        }

        internal int getId()
        {
            return code;
        }

        internal string getRole()
        {
            return role;
        }

        internal bool getIFA()
        {
            return isForAct;
        }

        internal string getEmail()
        {
            return emailAddress;
        }

        internal short getNID()
        {
            return numInDoc;
        }

        internal void setNID(short nid)
        {
            numInDoc = nid;
        }

        internal void setIFA(bool ifa)
        {
            this.isForAct = ifa;
        }

        internal bool getSendMail()
        {
            return sendMail;
        }

        internal void setRole(string newRole)
        {
            role = newRole;
        }
    }
}
