using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation.Peers;

namespace ActiveDirectoryAccess
{


    public class AdAccess
    {


        //TIPP:%SystemRoot%\SYSTEM32\rundll32.exe dsquery,OpenQueryWindow 



        //Für die Suche nach Beschreibung (in der Schule steht hier die Schüler Id

        /// <summary>
        /// Filter optionen der Benutzer Suche
        /// </summary>         
        public enum UserFilter
        {
            SamAccountName,
            UserName,
            Description,
            EmailAddress,
            EmployeeId,
            GivenName,
            MiddleName,
            Surname,
            VoiceTelephoneNumber,
            AccountExpirationDate,
            DisplayName,
            PasswordNeverExpires,
            UserCannotChangePassword,
            Enabled
        }


        //GroupScope = Gruppen Typ (Global,Universal,Lokal)
        /// <summary>
        /// Filter optionen der Gruppen Suche
        /// </summary>       
        public enum GroupFilter
        {
            SamAccountName,
            Name,
            Description,
            GroupScope,
            IsSecurityGroup,
            UserPrincipalName,
            DisplayName,
        }







      
        private string _domain;



        public string Domain
        {
            get { return _domain; }

            set { _domain = value; }
        }

        public AdAccess(string domain)
        {
            _domain = domain;
        }



        /// <summary>
        /// SearchUsers() Methode zum Suchen von Benutzern in einer Domäne 
        /// Wichtig!!!:liefert nur ein Suchergebnis (PrincipalSearchResult)
        /// Zum umwandeln in ein DirectoryEntry muss ConvertPrincipalsToDirectoryEntries();
        /// benutzt werden oder wenn man gleich ein DirectoryEntry will muss man GetUsersDirectoryEntry() benutzen
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public PrincipalSearchResult<Principal> SearchUsers(object[] filter, int[] searchmode)
        {
            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);

            //Ein "User-Objekt" im Kontext anlegen
            UserPrincipal user = new UserPrincipal(domainContext);

            for (int i = 0; i < searchmode.Length; i++)
            {

                switch (searchmode[i])
                {


                    //Den Suchparameter angeben
                    case (int)UserFilter.UserName:
                        //Layout für Namen Suche = Vorname Nachname
                        user.Name = filter[i].ToString();
                        break;
                    case (int)UserFilter.SamAccountName:
                        user.SamAccountName = filter[i].ToString();

                        break;
                    case (int)UserFilter.Description:
                        user.Description = filter[i].ToString();

                        break;
                    case (int)UserFilter.EmailAddress:
                        user.EmailAddress = filter[i].ToString();

                        break;
                    case (int)UserFilter.EmployeeId:
                        user.EmployeeId = filter[i].ToString();

                        break;
                    case (int)UserFilter.GivenName:
                        user.GivenName = filter[i].ToString();

                        break;
                    case (int)UserFilter.MiddleName:
                        user.MiddleName = filter[i].ToString();

                        break;
                    case (int)UserFilter.Surname:
                        user.Surname = filter[i].ToString();

                        break;
                    case (int)UserFilter.VoiceTelephoneNumber:
                        user.VoiceTelephoneNumber = filter[i].ToString();

                        break;
                    case (int)UserFilter.AccountExpirationDate:
                        user.AccountExpirationDate = Convert.ToDateTime(filter[i]);
                        break;
                    case (int)UserFilter.DisplayName:
                        user.DisplayName = filter[i].ToString();

                        break;
                    case (int)UserFilter.PasswordNeverExpires:
                        user.PasswordNeverExpires = Convert.ToBoolean(filter[i]);
                        break;
                    case (int)UserFilter.UserCannotChangePassword:
                        user.UserCannotChangePassword = Convert.ToBoolean(filter[i]);
                        break;
                    case (int)UserFilter.Enabled:
                        user.Enabled = Convert.ToBoolean(filter[i]);
                        break;
                    default:
                        throw new Exception("Error: Filter is less than 0 or greater than 13");
                }

            }


            //Den Sucher anlegen und ihm die Suchkriterien
            //(unser User-Objekt) übergeben
            
            PrincipalSearcher pS = new PrincipalSearcher();
            pS.QueryFilter = user;

            //Die Suche durchführen
            PrincipalSearchResult<Principal> results = pS.FindAll();
            
            return results;

        }


        /// <summary>
        /// SearchGroups() Methode zum Suchen von Gruppen in einer Domäne 
        /// Wichtig!!!:liefert nur ein Suchergebnis (PrincipalSearchResult)
        /// Zum umwandeln in ein DirectoryEntry muss ConvertPrincipalsToDirectoryEntries();
        /// benutzt werden oder wenn man gleich ein DirectoryEntry will muss man GetUsersDirectoryEntry() benutzen
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public PrincipalSearchResult<Principal> SearchGroups(object[] filter, int[] searchmode)
        {
            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);

            //Ein "User-Objekt" im Kontext anlegen
            GroupPrincipal group = new GroupPrincipal(domainContext);

            for (int i = 0; i < searchmode.Length; i++)
            {
                switch (searchmode[i])
                {


                    //Den Suchparameter angeben
                    case (int)GroupFilter.UserPrincipalName:
                        //Layout für Namen Suche = Vorname Nachname
                        group.UserPrincipalName = filter[i].ToString();

                        break;
                    case (int)GroupFilter.Description:
                        group.Description = filter[i].ToString();

                        break;
                    case (int)GroupFilter.DisplayName:
                        group.DisplayName = filter[i].ToString();

                        break;
                    case (int)GroupFilter.GroupScope:
                        group.GroupScope = (GroupScope)Convert.ToInt32(filter[i]);

                        break;
                    case (int)GroupFilter.IsSecurityGroup:
                        group.IsSecurityGroup = Convert.ToBoolean(filter[i]);
                        break;
                    case (int)GroupFilter.SamAccountName:
                        group.SamAccountName = filter[i].ToString();
                        break;

                    case (int)GroupFilter.Name:
                        group.Name = filter[i].ToString();

                        break;

                    default:
                        throw new Exception("Error: Filter is less than 0 or greater than 6");
                }
            }


            //Den Sucher anlegen und ihm die Suchkriterien
            //(unser Gruppen-Objekt) übergeben
            PrincipalSearcher pS = new PrincipalSearcher();
            pS.QueryFilter = group;

            //Die Suche durchführen
            PrincipalSearchResult<Principal> results = pS.FindAll();
            return results;

        }


        [Obsolete("Nur benutzen falls DirectoryEntry gebraucht wir ansonten Get* Methoden benutzen")]
        /// <summary>
        /// Wandelt Principal in DirectoryEntry um
        /// </summary>
        /// <param name="results"></param>
        /// <returns></returns>
        public List<DirectoryEntry> ConvertPrincipalsToDirectoryEntries(PrincipalSearchResult<Principal> results)
        {
            return (
                //Umwandlung Principal->DirectoryEntry
                results.ToList().Cast<Principal>().Select(pc => (DirectoryEntry)pc.GetUnderlyingObject())).ToList();

        }


        [Obsolete("Nur benutzen falls DirectoryEntry gebraucht wir ansonten Get* Methoden benutzen")]
        /// <summary>
        /// Wandelt Principal in DirectoryEntry um
        /// </summary>
        /// <param name="results"></param>
        /// <returns></returns>
        public List<DirectoryEntry> ConvertPrincipalsToDirectoryEntries(List<UserPrincipal> results)
        {
            return (
                //Umwandlung Principal->DirectoryEntry
                results.ToList().Cast<Principal>().Select(pc => (DirectoryEntry)pc.GetUnderlyingObject())).ToList();

        }

        [Obsolete("Nur benutzen falls DirectoryEntry gebraucht wir ansonten Get* Methoden benutzen")]
        /// <summary>
        /// Wandelt Principal in DirectoryEntry um
        /// </summary>
        /// <param name="results"></param>
        /// <returns></returns>
        public List<DirectoryEntry> ConvertPrincipalsToDirectoryEntries(List<GroupPrincipal> results)
        {


            List<DirectoryEntry> de = new List<DirectoryEntry>();
                    //Umwandlung Principal->DirectoryEntry
            foreach(Principal pc in results)
            {
                try
                {
                    de.Add((DirectoryEntry)pc.GetUnderlyingObject());
                }

                catch(Exception ex)
                {
                    
                }
            }

            return de;
            
        }

        [Obsolete("Bitte GetUsers Benutzen")]
        /// <summary>
        /// GetMembersinGroupDirectoryEntry() Methode zum Suchen von Benutzern in Domäne
        /// und umwandlung in DirectoryEntry
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>

        public List<DirectoryEntry> GetUsersDirectoryEntry(object[] filter, int[] searchmode)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchUsers(filter, searchmode));

        }

        [Obsolete("Bitte GetGroups Benutzen")]
        /// <summary>
        /// GetGroupsDirectoryEntry() Methode zum Suchen von Gruppen in Domäne
        /// und umwandlung in DirectoryEntry
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<DirectoryEntry> GetGroupsDirectoryEntry(object[] filter, int[] seachmode)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchGroups(filter, seachmode));
        }


        [Obsolete("verltete methjode")]

        public void test()
        {
            MessageBox.Show("Hello");
        }


        /// <summary>
        /// SearchMembersinGroup() Methode zum Suchen von Benutzern in Gruppe in einer Domäne 
        /// Wichtig!!!:liefert nur ein Suchergebnis (PrincipalSearchResult)
        /// Zum umwandeln in ein DirectoryEntry muss ConvertPrincipalsToDirectoryEntries();
        /// benutzt werden oder wenn man gleich ein DirectoryEntry will muss man GetUsersDirectoryEntry() benutzen
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<UserPrincipal> SearchMembersinGroup(string groupname, bool searchsubgroupmembers, List<Principal> checkredundancy = null)
        {
            //Beim erstauruf der Methode ist die Liste null
            //kann auch dafür verwedet werden beim erstaufruf Nutzer und Gruppen auszuschileßen

            if (checkredundancy == null)
            {
                checkredundancy = new List<Principal>();
            }

            List<UserPrincipal> users = new List<UserPrincipal>();

            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);
      
            //Gruppe bestimmen
            GroupPrincipal group = GroupPrincipal.FindByIdentity(domainContext, groupname);


            if (group != null)
            {
                //Mitglieder der Gruppe herausholen
                PrincipalSearchResult<Principal> results = group.GetMembers();

               

                
                foreach (Principal p in results)
                {
                    //Nach Mitgileder in den Subgruppen Suchen
                    if (p is GroupPrincipal && searchsubgroupmembers)
                    {
                        bool found = false;

                        //Verhindern von endlosschleifen (z.B Gruppe1 mitglied Gruppe 3 mitglied Gruppe 2 mitglied Gruppe 1)
                        foreach (Principal p2 in checkredundancy)
                        {
                            if (p2 is GroupPrincipal)
                            {
                                if (((GroupPrincipal)p2).Name.Equals(((GroupPrincipal)p).Name))
                                {
                                    found = true;
                                    break;
                                    
                                }
                            }
                        }

                        if (!found)
                        {   //Rekursiver aufruf mit den SubGruppen
                            checkredundancy.Add((GroupPrincipal)p);
                            users.AddRange(SearchMembersinGroup(p.SamAccountName, true, checkredundancy));
                        }
                    }

                    //Falls user wenden die benutzer zur liste hinzugefügt
                    else if (p is UserPrincipal)
                    {
                        bool found = false;

                        foreach (Principal p2 in checkredundancy)
                        {
                            if (p2 is UserPrincipal)
                            {
                                if (((UserPrincipal)p2).Name.Equals(((UserPrincipal)p).Name))
                                {
                                    found = true;
                                    break;

                                }
                            }
                        }

                        if (!found)
                        {
                            checkredundancy.Add(p);

                            //Benutzer wird zur liste hindzugefügt
                            users.Add((UserPrincipal)p);
                        }
                    }
                }


            }


            return users;
        }




        /// <summary>
        /// SearchSubGroups() Methode zum Suchen von Sub-Gruppen in einer Domäne 
        /// Wichtig!!!:liefert nur ein Suchergebnis (PrincipalSearchResult)
        /// Zum umwandeln in ein DirectoryEntry muss ConvertPrincipalsToDirectoryEntries();
        /// benutzt werden oder wenn man gleich ein DirectoryEntry will muss man GetUsersDirectoryEntry() benutzen
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<GroupPrincipal> SearchSubGroups(string groupname, bool searchsubgroupmembers, List<Principal> checkredundancy = null)
        {
            //Beim erstauruf der Methode ist die Liste null
            //kann auch dafür verwedet werden beim erstaufruf Nutzer und Gruppen auszuschileßen
            if (checkredundancy == null)
            {
                checkredundancy = new List<Principal>();
            }



            List<GroupPrincipal> groups = new List<GroupPrincipal>();
            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);

            //Gruppe bestimmen
            GroupPrincipal group = GroupPrincipal.FindByIdentity(domainContext, groupname);

            if (group != null)
            {
                PrincipalSearchResult<Principal> results = group.GetMembers();


                //Nach Subgruppen Suchen
                foreach (Principal p in results)
                {

                    //Auf redundanz prüfen
                    if (p is GroupPrincipal)
                    {
                        bool found = false;

                        foreach (Principal p2 in checkredundancy)
                        {
                            if (p2 is GroupPrincipal)
                            {
                                if (((GroupPrincipal)p2).Name.Equals(((GroupPrincipal)p).Name))
                                {
                                    found = true;
                                    break;
                                }
                            }
                        }

                        if (!found)
                        {
                            checkredundancy.Add((GroupPrincipal)p);
                            groups.Add((GroupPrincipal)p);

                            //Falls auch Subgruppen in den Subgruppen gesucht werden sollen
                            if (searchsubgroupmembers)
                            {
                                groups.AddRange(SearchSubGroups(p.SamAccountName, true, checkredundancy));
                            }
                        }
                    }


                }

            }

            return groups;

        }

        [Obsolete("Bitte GetGroupMembers Benutzen")]
        /// <summary>
        /// GetMembersinGroupDirectoryEntry() Methode zum Suchen von Benutzern in Gruppen
        /// und umwandlung in DirectoryEntry
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>

        public List<DirectoryEntry> GetMembersinGroupDirectoryEntry(string groupname, bool searchsubgroupmembers)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchMembersinGroup(groupname, searchsubgroupmembers));

        }

        /// <summary>
        /// GetUsers() Methode zum Suchen von Usern in Domäne 
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<UserPrincipal> GetUsers(object[] filter, int[] searchmode)
        {
            List<UserPrincipal> up = new List<UserPrincipal>();
            PrincipalSearchResult<Principal> ps  = SearchUsers(filter,searchmode);
            foreach(Principal p in ps)
            {
                if(p is UserPrincipal)
                {
                    up.Add((UserPrincipal)p);
                }
            }

            return up;
        }

        /// <summary>
        /// GetGroups() Methode zum Suchen von Gruppen in Domäne 
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<GroupPrincipal> GetGroups(object[] filter, int[] searchmode)
        {
            List<GroupPrincipal> gp = new List<GroupPrincipal>();
            PrincipalSearchResult<Principal> ps = SearchGroups(filter, searchmode);
            foreach (Principal p in ps)
            {
                if (p is GroupPrincipal)
                {
                    gp.Add((GroupPrincipal)p);
                }
            }

            return gp;
        }


        /// <summary>
        /// GetUsers() Methode zum Suchen von Usern und (Usern von Sub  in Gruppe 
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>
        public List<UserPrincipal> GetGroupMembers(string groupname,bool searchsubgroupmembers,List<Principal> checkredundancy = null)
        {          
            return SearchMembersinGroup(groupname,searchsubgroupmembers,checkredundancy);           
        }


        public List<GroupPrincipal> GetSubGroups(string groupname, bool searchsubgroupmembers, List<Principal> checkredundancy = null)
        {
            return SearchSubGroups(groupname, searchsubgroupmembers,checkredundancy);
        }

        [Obsolete("Bitte GetSubGroups Benutzen")]

        /// <summary>
        /// GetMembersinGroupsDirectoryEntry() Methode zum Suchen von Sub-Gruppen in einer Gruppe 
        /// und umwandlung in DirectoryEntry
        /// </summary>
        /// <param name="filter"></param>
        /// <param name="searchmode"></param>
        /// <returns></returns>


        public List<DirectoryEntry> GetSubGroupsDirectoryEntry(string groupname, bool searchsubgroupmembers)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchSubGroups(groupname, searchsubgroupmembers));

        }


    }


}





