using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ActiveDirectoryAccess;

namespace AdAccessTest1
{
    class Program
    {
        static void Main(string[] args)
        {
            AdAccess ad = new AdAccess("BSZ.local");
            List<UserPrincipal> de = ad.GetUsers(new string[] { "Const*" }, new int[] { (int)AdAccess.UserFilter.UserName });
            
            //1. Test Namen ausgabe
            Console.WriteLine(de.Count);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Surname);
            List<GroupPrincipal> gp = ad.GetGroups(new string[] { "BFI*" }, new int[] { (int)AdAccess.UserFilter.UserName });
            for (int i = 0; i < gp.Count; i++)
                Console.WriteLine(gp[i].Name);

            de = ad.GetGroupMembers("BFI11a", true);
            for(int i = 0;i<de.Count;i++)
                Console.WriteLine(de[i].SamAccountName);
            gp = ad.GetSubGroups("Benutzer", true);
               for (int i = 0; i < gp.Count; i++)
                Console.WriteLine(gp[i].Name);

            Console.WriteLine(ad.LoginCheck("romc","Herdsfsdfsd,"));
            Console.ReadKey();
        }
    }
}
