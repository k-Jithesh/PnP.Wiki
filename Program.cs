using Microsoft.SharePoint.Client;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Tooling.Connector;
using Microsoft.Xrm.Sdk.Query;
using PowerApps.Samples;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PnP.Wiki
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            /*
            string siteUrl = "https://aucklandtransport.sharepoint.com/sites/KBBK";
            var authManager = new OfficeDevPnP.Core.AuthenticationManager();
            // This method calls a pop up window with the login page and it also prompts 
            // for the multi factor authentication code. 
            ClientContext ctx = authManager.GetWebLoginClientContext(siteUrl);
            // The obtained ClientContext object can be used to connect to the SharePoint site. 
            Web web = ctx.Web;
            ctx.Load(web, w => w.Title);
            ctx.ExecuteQuery();
            Console.WriteLine("You have connected to {0} site, with Multi Factor Authentication enabled!!", web.Title);
            Console.ReadLine();
            */

            List<string> kbIds = new List<string>() { "437", "677", "1899", "2423", "2585" };

            List<KBModel> KBList = new List<KBModel>();

            KBList = System.IO.File.ReadAllLines(@"C:\Jithesh\Customer\AT\SPO\Wiki\list.csv")
                                    .Skip(1)
                                    .Select(v => KBModel.FromCsv(v))
                                    .ToList<KBModel>();

            List<KBModel> migrationList = new List<KBModel>();

            foreach (KBModel item in KBList)
            {
                if (kbIds.Contains(item.Id.ToString()))
                {
                    migrationList.Add(item);
                }
            }

            CrmServiceClient service = null;

            try
            {
                service = SampleHelpers.Connect("Connect");

                if (service != null)
                {
                    // Service implements IOrganizationService interface 
                    if (service.IsReady)
                    {
                        #region Sample Code
                        //////////////////////////////////////////////
                        #region Demonstrate

                        // Instantiate an account object.

                        foreach (KBModel item in migrationList)
                        {
                            Entity newKb = new Entity("knowledgearticle");
                            newKb["title"] = item.Title.ToString();
                            newKb["content"] = GetHtml(item.Id.ToString());
                            newKb["keywords"] = item.Keywords.ToString();
                            newKb["description"] = item.FileLeafRef.ToString();

                            Guid kbId = service.Create(newKb);
                            Console.WriteLine(kbId);
                        }

                        #endregion Demonstrate
                        //////////////////////////////////////////////
                        #endregion Sample Code

                        Console.WriteLine("The sample completed successfully");
                        return;
                    }
                    else
                    {
                        const string UNABLE_TO_LOGIN_ERROR = "Unable to Login to Microsoft Dataverse";
                        if (service.LastCrmError.Equals(UNABLE_TO_LOGIN_ERROR))
                        {
                            Console.WriteLine("Check the connection string values in cds/App.config.");
                            throw new Exception(service.LastCrmError);
                        }
                        else
                        {
                            throw service.LastCrmException;
                        }
                    }
                }
            }

            catch (Exception ex)
            {
                SampleHelpers.HandleException(ex);
            }


            finally
            {
                if (service != null)
                    service.Dispose();

                Console.WriteLine("Press <Enter> to exit.");
                Console.ReadLine();
            }



        }

        private static string GetHtml(string docId)
        {
            StringBuilder contents = new StringBuilder();

            contents.Append(System.IO.File.ReadAllText("C:\\Jithesh\\Customer\\AT\\SPO\\Wiki\\Files\\" + docId + ".html"));

            contents.Replace("<br>", "<br />");

            XElement wikiField = XElement.Parse(contents.ToString());
            XElement div = wikiField.Descendants("div").Where(e =>
                (string)e.Attribute("class") == "ms-rte-layoutszone-inner")
                .FirstOrDefault();
            return div.ToString();
        }
    }

    class KBModel
    {
        public int Id { get; set; }
        public string Title { get; set; }
        public string Categ { get; set; }
        public string FileLeafRef { get; set; }
        public string Keywords { get; set; }

        public static KBModel FromCsv(string csvLine)
        {
            string[] values = csvLine.Split(',');
            KBModel kb = new KBModel();
            int i;
            int.TryParse(values[0].Replace("\"",""), out i);
            kb.Id = i;
            kb.Title = values[1].Replace("\"", "");
            kb.Categ = values[2].Replace("\"", "");
            kb.FileLeafRef = values[3].Replace("\"", "");
            kb.Keywords = values[4].Replace("\"", "");

            return kb;
        }
    }

}
