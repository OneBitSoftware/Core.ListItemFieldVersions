using System;
using System.Net;
using System.Security;
using System.Xml;
using Core.ListItemFieldVersions.ListsASMX;
using Microsoft.SharePoint.Client;

namespace Core.ListItemFieldVersions
{
    public class VersionsHandler
    {
        private const string ListsServiceUrl = "/_vti_bin/Lists.asmx";
        private Lists lists = null;
        public string TenantUrl{get;set;}
        public String User { get; set; }
        public String Password { get; set; }
        public string Domain { get; set; }
        public string MySiteHost { get; set; }
        private Lists _lists
        {
            get
            {
                if (lists == null)
                {
                    if (!String.IsNullOrEmpty(TenantUrl))
                    {
                        this.lists = new Lists();
                        lists.Url = TenantUrl + ListsServiceUrl;
                        lists.UseDefaultCredentials = false;
                        lists.CookieContainer = new CookieContainer();
                        lists.CookieContainer.Add(GetFedAuthCookie(CreateSharePointOnlineCredentials()));
                        return lists;
                    }
                    else if (this.User.Length > 0 && this.Password.Length > 0 && this.Domain.Length > 0 && this.MySiteHost.Length > 0)
                    {
                        this.lists = new Lists();
                        lists.Url = this.MySiteHost + ListsServiceUrl;
                        NetworkCredential credential = new NetworkCredential(this.User, this.Password, this.Domain);
                        CredentialCache credentialCache = new CredentialCache();
                        credentialCache.Add(new Uri(this.MySiteHost), "NTLM", credential);
                        lists.Credentials = credentialCache;
                        return lists;
                    }
                    else
                    {
                        throw new Exception("Please specify an authentication provider or specify domain credentials");
                    }
                }
                else
                {
                    return this.lists;
                }
            }
        }

        public XmlNode GetVersionCollection(string listId, string itemId, string fieldName)
        {
            return _lists.GetVersionCollection(listId, itemId, fieldName);
        }

        private SharePointOnlineCredentials CreateSharePointOnlineCredentials()
        {
            var spoPassword = new SecureString();
            foreach (char c in Password)
            {
                spoPassword.AppendChar(c);
            }
            return new SharePointOnlineCredentials(User, spoPassword);
        }

        private Cookie GetFedAuthCookie(SharePointOnlineCredentials credentials)
        {
            string authCookie = credentials.GetAuthenticationCookie(new Uri(this.TenantUrl));
            if (authCookie.Length > 0)
            {
                return new Cookie("SPOIDCRL", authCookie.TrimStart("SPOIDCRL=".ToCharArray()), String.Empty, new Uri(this.TenantUrl).Authority);
            }
            else
            {
                return null;
            }
        }
    }   
}
