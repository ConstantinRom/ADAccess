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
            List<DirectoryEntry> de = ad.GetUsers("*Meier*", (int)AdAccess.UserFilter.UserName);
            Console.WriteLine(de.Count);
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());

            de = ad.GetGroups("BFI11a", (int)AdAccess.GroupFilter.Name);
            
            for (int i = 0; i < de.Count; i++)
                Console.WriteLine(de[i].Properties["SamAccountName"].Value.ToString());
            Console.ReadKey();
        }
    }
}
