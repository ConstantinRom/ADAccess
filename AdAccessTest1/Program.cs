using System;
using System.Collections.Generic;
using System.DirectoryServices;
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
            Console.WriteLine("Test123");
        
            AdAccess ad = new AdAccess("BSZ.local");
            List<DirectoryEntry> de = ad.GetUsersDirectoryEntry(new string[] { "*rom*" }, new int[] { (int)AdAccess.UserFilter.SamAccountName });
           
            //1. Test Namen ausgabe
            Console.WriteLine(de.Count);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());



            //2.Test Gruppen ausgabe
            de = ad.GetGroupsDirectoryEntry(new string[] { "*rom*" },  new int[] { (int)AdAccess.GroupFilter.Name });                   
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());

            
            de = ad.GetMembersinGroupDirectoryEntry("Benutzer",true);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());

            de = ad.GetSubGroupsDirectoryEntry("Benutzer", true);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());

           
            Console.ReadKey();
        }
    }
}
