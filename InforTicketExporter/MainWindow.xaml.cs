using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Xml;
using System.Xml.XPath;

namespace InforTicketExporter
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();


        }

        private async void FetchResults_Button_Click(object sender, RoutedEventArgs e)
        {
            InforIntegration.Tasks.SecureCreds(UsernameTextBox.Text, PasswordTextBox.Password);

            string BaseURL = BaseURLTextBox.Text;
            string APIQuery = APIQueryTextBox.Text;

            List<XmlDocument> QueryResultDocuments = await GetTicketInfo(BaseURL, APIQuery);

            List<string> JsonTickets = ConvertXmlDocumentListToJsonTickets(QueryResultDocuments);
        }

        private async Task<List<XmlDocument>> GetTicketInfo(string BaseURL, string APIQuery)
        {
            List<XmlDocument> QueryResultDocuments = new List<XmlDocument>();
            bool ItsGoTime = true;
            int StartIndex = 1;

            while (ItsGoTime)
            {
                APIQuery = APIQuery + "&StartIndex=" + StartIndex.ToString();

                XmlDocument ResultDocument = await InforIntegration.Tasks.RunInforGet(BaseURL, APIQuery);
                
                QueryResultDocuments.Add(ResultDocument);
                
                // Build a Namespace Manager using the namespaces from the XML
                XmlNamespaceManager NamespaceManager = new XmlNamespaceManager(ResultDocument.NameTable);
                XPathNavigator RootNode = ResultDocument.CreateNavigator();
                RootNode.MoveToFollowing(XPathNodeType.Element);
                IDictionary<string, string> Namespaces = RootNode.GetNamespacesInScope(XmlNamespaceScope.All);
                foreach (KeyValuePair<string, string> kvp in Namespaces)
                {
                    NamespaceManager.AddNamespace(kvp.Key, kvp.Value);
                }

                // Get the total result count, start index and items per page
                int TotalResultCount = Convert.ToInt32(ResultDocument.SelectSingleNode("//opensearch:totalResults", NamespaceManager).InnerText);
                int LastStartIndex = Convert.ToInt32(ResultDocument.SelectSingleNode("//opensearch:startIndex", NamespaceManager).InnerText);
                int ItemsPerPage = Convert.ToInt32(ResultDocument.SelectSingleNode("//opensearch:itemsPerPage", NamespaceManager).InnerText);

                // Check if there is more to this tale
                if ((LastStartIndex - 1 + ItemsPerPage) < TotalResultCount)
                {
                    // Update the StartIndex, prepare to go again
                    StartIndex = LastStartIndex + ItemsPerPage;
                    ItsGoTime = true;
                }
                else
                {
                    ItsGoTime = false;
                }
            }

            return QueryResultDocuments;
        }

        private List<string> ConvertXmlDocumentListToJsonTickets(List<XmlDocument> QueryResultDocuments)
        {
            List<string> JsonTickets = new List<string>();

            foreach (XmlDocument Document in QueryResultDocuments)
            {
                // Build a Namespace Manager using the namespaces from the XML
                XmlNamespaceManager NamespaceManager = new XmlNamespaceManager(Document.NameTable);
                XPathNavigator RootNode = Document.CreateNavigator();
                RootNode.MoveToFollowing(XPathNodeType.Element);
                IDictionary<string, string> Namespaces = RootNode.GetNamespacesInScope(XmlNamespaceScope.All);
                foreach (KeyValuePair<string, string> kvp in Namespaces)
                {
                    NamespaceManager.AddNamespace(kvp.Key, kvp.Value);
                }

                XmlNodeList TicketsXml = Document.SelectNodes("//sdata:payload/slx:Ticket", NamespaceManager);

                foreach (XmlNode TicketXml in TicketsXml)
                {
                    string JsonString = JsonConvert.SerializeXmlNode(TicketXml);
                    JsonTickets.Add(JsonString);
                }                
            }

            return JsonTickets;
        }
    }
}
