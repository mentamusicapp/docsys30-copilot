using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DocumentsModule
{
    public class RecipientList
    {
        internal short id { get; private set; }
        internal int owner { get; set; }
        internal RecipientListsLevel level { get; set; }
        internal Branch branch { get; set; }
        internal string name { get; set; }
        List<Recipient> recipients;

        public RecipientList(short id, int owner, int rll, char branch, string name, List<Recipient> recipients)
        {
            this.id = id;
            this.owner = owner;
            this.level = (RecipientListsLevel)rll;
            this.branch = (Branch)branch;
            this.name = name;
            this.recipients = recipients;
        }

        internal string getLevelString()
        {
            switch(level)
            {
                case RecipientListsLevel.personal:
                    return "אישית";
                case RecipientListsLevel.branch:
                    return "ענפית";
                case RecipientListsLevel.unit:
                    return "יחידתית";
            }
            return null;
        }

        internal string getOwnerName()
        {
            return PublicFuncsNvars.getUserNameByUserCode(owner);
        }

        internal List<Recipient> getRecipients()
        {
            return recipients;
        }

        internal void removeRecipient(short nid)
        {
            recipients.RemoveAll(x => x.getNID() == nid);
        }

        internal void addRecipient(Recipient recipient)
        {
            recipients.Add(recipient);
        }
    }
}
