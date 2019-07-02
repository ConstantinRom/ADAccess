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
            Console.ReadKey();
            AdAccess ad = new AdAccess("BSZ.local");
            List<DirectoryEntry> de = ad.GetUsers(new string[] { "r*","con*" }, new int[] { (int)AdAccess.UserFilter.SamAccountName,(int)AdAccess.UserFilter.UserName });
           
            //1. Test Namen ausgabe
            Console.WriteLine(de.Count);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());



            //2.Test Gruppen ausgabe
            de = ad.GetGroups(new string[] { "BFI*" },  new int[] { (int)AdAccess.GroupFilter.Name });                   
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());


            de = ad.GetMembersinGroup("BFI11a");
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());
            Console.ReadKey();
        }
    }
}
