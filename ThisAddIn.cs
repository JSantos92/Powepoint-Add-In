using Google.Apis.Auth.OAuth2;
using Google.Apis.Script.v1;
using Google.Apis.Script.v1.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace MyRibbonAddIn
{
    public partial class ThisAddIn
    {

        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/script-dotnet-quickstart.json
        
        static string[] Scopes = { "https://www.googleapis.com/auth/forms" };
        static string ApplicationName = "Google Apps Script Execution API .NET Quickstart";
        // Newtonsoft.Json.Linq.JObject responseSet;
        public int currentSlide = 0;


        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.PresentationNewSlide += new PowerPoint.EApplication_PresentationNewSlideEventHandler(Application_PresentationNewSlide);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new MyRibbon(this);
        }

        void Application_PresentationNewSlide(PowerPoint.Slide Sld)
        {
            currentSlide += 1;
        }

        public Newtonsoft.Json.Linq.JObject getFormResponses(string formURL)
        {
            UserCredential credential;

            using (var stream =
                new FileStream(Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)+"\\client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/script-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Google Apps Script Execution API service.
            string scriptId = "MIZ8ME5AgqFMS6rD8ZLQTQxU3ND9pJt_J";
            var service = new ScriptService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            // Create an execution request object.
            ExecutionRequest request = new ExecutionRequest();
            request.Function = "getFormAnswers";

            string[] requestParam = new string[] { formURL };

            request.Parameters = requestParam;

            ScriptsResource.RunRequest runReq =
                    service.Scripts.Run(request, scriptId);

            try
            {
                // Make the API request.
                Operation op = runReq.Execute();

                if (op.Error != null)
                {
                    // The API executed, but the script returned an error.

                    // Extract the first (and only) set of error details
                    // as a IDictionary. The values of this dictionary are
                    // the script's 'errorMessage' and 'errorType', and an
                    // array of stack trace elements. Casting the array as
                    // a JSON JArray allows the trace elements to be accessed
                    // directly.
                    IDictionary<string, object> error = op.Error.Details[0];
                    Console.WriteLine(
                            "Script error message: {0}", error["errorMessage"]);
                    if (error["scriptStackTraceElements"] != null)
                    {
                        // There may not be a stacktrace if the script didn't
                        // start executing.
                        Console.WriteLine("Script error stacktrace:");
                        Newtonsoft.Json.Linq.JArray st =
                            (Newtonsoft.Json.Linq.JArray)error["scriptStackTraceElements"];
                        foreach (var trace in st)
                        {
                            Console.WriteLine(
                                    "\t{0}: {1}",
                                    trace["function"],
                                    trace["lineNumber"]);
                        }
                    }
                }
                else
                {
                    // The result provided by the API needs to be cast into
                    // the correct type, based upon what types the Apps
                    // Script function returns. Here, the function returns
                    // an Apps Script Object with String keys and values.
                    // It is most convenient to cast the return value as a JSON
                    // JObject (folderSet).

                    Newtonsoft.Json.Linq.JObject response = (Newtonsoft.Json.Linq.JObject)op.Response["result"];

                    if (response.Count == 0)
                    {
                        Console.WriteLine("No responses returned!");

                    }
                    else
                    {
                        Console.WriteLine("Responses:");
                        Console.WriteLine(response);
                        /*foreach (var folder in folderSet)
                        {
                            Console.WriteLine(
                                "\t{0} ({1})", folder.Value, folder.Key);
                        }*/
                    }

                    return response;
                }
            }
            catch (Google.GoogleApiException entry)
            {
                // The API encountered a problem before the script
                // started executing.
                Console.WriteLine("Error calling API:\n{0}", entry);
            }

            return null;
        }



        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
