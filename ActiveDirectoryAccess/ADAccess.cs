using System;
using System.Collections.Generic;
using System.DirectoryServices;
using System.DirectoryServices.AccountManagement;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace ActiveDirectoryAccess
{


    public class AdAccess
    {





        //Für die Suche nach Anmeldename => romc
        public const int SamAccountName = 0;

        //Für die Suche nach Name => Constantin Rom
        public const int UserName = 1;

        //Für die Suche nach Beschreibung (in der Schule steht hier die Schüler Id
        public enum UserFilter
        { SamAccountName, UserName, Description, EmailAddress, EmployeeId, GivenName, MiddleName, Surname, VoiceTelephoneNumber, AccountExpirationDate, DisplayName, PasswordNeverExpires, UserCannotChangePassword, Enabled }


        //GroupScope = Gruppen Typ (Global,Universal,Lokal)
        public enum GroupFilter
        { SamAccountName, Name, Description, GroupScope, IsSecurityGroup, UserPrincipalName, DisplayName, }







        private string _uid;
        private string _pwd;
        private string _domain;

        public string Uid
        {
            get { return _uid; }

            set { _uid = value; }
        }

        public string Pwd
        {
            get { return _pwd; }

            set { _pwd = value; }
        }

        public string Domain
        {
            get { return _domain; }

            set { _domain = value; }
        }

        public AdAccess(string domain)
        {
            _domain = domain;
        }

        public PrincipalSearchResult<Principal> SearchUsers(object filter, int searchmode)
        {
            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);

            //Ein "User-Objekt" im Kontext anlegen
            UserPrincipal user = new UserPrincipal(domainContext);


            switch (searchmode)
            {
                //Den Suchparameter angeben
                case (int)UserFilter.UserName:
                    //Layout für Namen Suche = Vorname Nachname
                    user.Name = filter.ToString();
                    break;
                case (int)UserFilter.SamAccountName:
                    user.SamAccountName = filter.ToString();

                    break;
                case (int)UserFilter.Description:
                    user.Description = filter.ToString();

                    break;
                case (int)UserFilter.EmailAddress:
                    user.EmailAddress = filter.ToString();

                    break;
                case (int)UserFilter.EmployeeId:
                    user.EmployeeId = filter.ToString();

                    break;
                case (int)UserFilter.GivenName:
                    user.GivenName = filter.ToString();

                    break;
                case (int)UserFilter.MiddleName:
                    user.MiddleName = filter.ToString();

                    break;
                case (int)UserFilter.Surname:
                    user.Surname = filter.ToString();

                    break;
                case (int)UserFilter.VoiceTelephoneNumber:
                    user.VoiceTelephoneNumber = filter.ToString();

                    break;
                case (int)UserFilter.AccountExpirationDate:
                    user.AccountExpirationDate = Convert.ToDateTime(filter);
                    break;
                case (int)UserFilter.DisplayName:
                    user.DisplayName = filter.ToString();

                    break;
                case (int)UserFilter.PasswordNeverExpires:
                    user.PasswordNeverExpires = Convert.ToBoolean(filter);
                    break;
                case (int)UserFilter.UserCannotChangePassword:
                    user.UserCannotChangePassword = Convert.ToBoolean(filter);
                    break;
                case (int)UserFilter.Enabled:
                    user.Enabled = Convert.ToBoolean(filter);
                    break;
                default:
                    throw new Exception("Error: Filter is less than 0 or greater than 13");
            }


            //Den Sucher anlegen und ihm die Suchkriterien
            //(unser User-Objekt) übergeben
            PrincipalSearcher pS = new PrincipalSearcher();
            pS.QueryFilter = user;

            //Die Suche durchführen
            PrincipalSearchResult<Principal> results = pS.FindAll();
            return results;

        }

        public PrincipalSearchResult<Principal> SearchGroups(object filter, int searchmode)
        {
            // Einen Kontext zur entsprechenden Windows Domäne erstellen
            PrincipalContext domainContext = new PrincipalContext(ContextType.Domain,
                _domain);

            //Ein "User-Objekt" im Kontext anlegen
            GroupPrincipal group = new GroupPrincipal(domainContext);


            switch (searchmode)
            {


                //Den Suchparameter angeben
                case (int)GroupFilter.UserPrincipalName:
                    //Layout für Namen Suche = Vorname Nachname
                    group.UserPrincipalName = filter.ToString();

                    break;
                case (int)GroupFilter.Description:
                    group.Description = filter.ToString();

                    break;
                case (int)GroupFilter.DisplayName:
                    group.DisplayName = filter.ToString();

                    break;
                case (int)GroupFilter.GroupScope:
                    group.GroupScope = (GroupScope)Convert.ToInt32(filter);

                    break;
                case (int)GroupFilter.IsSecurityGroup:
                    group.IsSecurityGroup = Convert.ToBoolean(filter);
                    break;
                case (int)GroupFilter.SamAccountName:
                    group.SamAccountName = filter.ToString();
                    break;

                case (int)GroupFilter.Name:
                    group.Name = filter.ToString();

                    break;

                default:
                    throw new Exception("Error: Filter is less than 0 or greater than 6");
            }


            //Den Sucher anlegen und ihm die Suchkriterien
            //(unser User-Objekt) übergeben
            PrincipalSearcher pS = new PrincipalSearcher();
            pS.QueryFilter = group;

            //Die Suche durchführen
            PrincipalSearchResult<Principal> results = pS.FindAll();
            return results;

        }

        public List<DirectoryEntry> ConvertPrincipalsToDirectoryEntries(PrincipalSearchResult<Principal> results)
        {
            return (
                //Umwandlung Principal->DirectoryEntry
                results.ToList().Cast<Principal>().Select(pc => (DirectoryEntry)pc.GetUnderlyingObject())).ToList();

        }

        public List<DirectoryEntry> GetUsers(object filter, int searchmode)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchUsers(filter, searchmode));

        }

        public List<DirectoryEntry> GetGroups(object filter, int seachmode)
        {
            return ConvertPrincipalsToDirectoryEntries(SearchGroups(filter, seachmode));
        }

   
    
    } 


 }





