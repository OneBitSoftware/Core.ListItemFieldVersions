using System;
using System.Configuration;
using System.Security;
using System.Xml;
using Microsoft.SharePoint.Client;

namespace Core.ListItemFieldVersions
{
    class Program
    {
        private static Web _web;
        private static ClientContext _context;

        static void Main(string[] args)
        {
            _context = new ClientContext(ConfigurationManager.AppSettings["TenantUrl"]);
            string userName = ConfigurationManager.AppSettings["AdminUser"];
            var passWord = new SecureString();
            foreach (char c in ConfigurationManager.AppSettings["Password"].ToCharArray())
            {
                passWord.AppendChar(c);
            }
            _context.Credentials = new SharePointOnlineCredentials(userName, passWord);
            _web = _context.Web;
            var list = _context.Web.Lists.GetByTitle("Retention Rules");
            var query = new CamlQuery();
            query.ViewXml =
              @"<View Scope='RecursiveAll'>  
                            <Query> 
                               <Where><Eq><FieldRef Name='FSObjType' /><Value Type='Integer'>0</Value></Eq></Where> 
                            </Query> 
                            <RowLimit>5000</RowLimit> 
                      </View>";
            var listItems = list.GetItems(query);
            _context.Load(_web);
            _context.Load(list);
            _context.Load(listItems);
            _context.ExecuteQuery();

            var versionsHandler = new VersionsHandler();
            versionsHandler.User = userName;
            versionsHandler.Password = ConfigurationManager.AppSettings["Password"];
            versionsHandler.TenantUrl = ConfigurationManager.AppSettings["TenantUrl"];
            
            if(listItems.Count == 0) throw new ArgumentException("No list items");

            var listId = list.Id.ToString();
            var itemId = listItems[0].Id.ToString();

            var versionNodes = versionsHandler.GetVersionCollection(listId, itemId, "Title");
            foreach (XmlNode node in versionNodes.ChildNodes)
            {
                if (node.Attributes != null)
                {
                    var title = node.Attributes["Title"].Value;
                    var modified = node.Attributes["Modified"].Value;

                    Console.WriteLine("Modified: {0}, Title: {1}", modified, title);
                }
            }
            Console.WriteLine("Hit any key to end.");
            Console.ReadKey();
        }
    }
}
