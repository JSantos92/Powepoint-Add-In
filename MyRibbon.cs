using Microsoft.VisualBasic;
using QRCoder;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new MyRibbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a bu tton.Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace MyRibbonAddIn
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        private ThisAddIn addin;
        string globalUrl;
        List<string> formQuestions = new List<string>();
        public MyRibbon(ThisAddIn addin)
        {
            this.addin = addin;
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MyRibbonAddIn.MyRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void createFormPage(Office.IRibbonControl control)
        {
            string formURL = Interaction.InputBox("Insert Google Form URL:", "Google Form URL", "https://docs.google.com/forms/d/11Vnlhtcw_kvjAB5QZFrYXiRUiu5Imlix-H8XOOEp9Vs/edit", 800, 500);

            globalUrl = formURL;

            // Generating TinyUrl Link and adding it to current slide

            Uri address = new Uri("http://tinyurl.com/api-create.php?url=" + formURL);
            System.Net.WebClient client = new System.Net.WebClient();
            string tinyUrl = client.DownloadString(address);
            Console.WriteLine(tinyUrl);

            PowerPoint.Shape textBox = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 350, 50, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter(tinyUrl.ToString());

            // Generating QrCode and adding it to current slide

            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(tinyUrl, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            qrCodeImage.Save("./qrcode.bmp");

            PowerPoint.Shape ppPicture = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddPicture("./qrcode.bmp", Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 285, 150, 350, 350);

        }

        public void generateResponsesSlides(Office.IRibbonControl control)
        {
            Newtonsoft.Json.Linq.JObject responses = this.addin.getFormResponses(globalUrl);

            List<string> questions = parseQuestions(responses);

            List<string> answers = parseAnswers(responses);

            PowerPoint.CustomLayout layout = this.addin.Application.ActivePresentation.SlideMaster.CustomLayouts[1];

            for (int i = 0; i < questions.Count(); i++)
            {
                this.addin.Application.ActivePresentation.Slides.AddSlide(this.addin.currentSlide+1, layout);

                this.addin.Application.ActiveWindow.Presentation.Slides[this.addin.currentSlide + 1].Select();

                PowerPoint.Shape textBox = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 305, 50, 400, 400);
                textBox.TextFrame.TextRange.Font.Size = 48;
                textBox.TextFrame.TextRange.InsertAfter(questions[i]);

                // TODO remove placeholders
            }



        }

        public List<string> parseQuestions(Newtonsoft.Json.Linq.JObject responses)
        {
            List<string> listQuestions = new List<string>();

            foreach (KeyValuePair<string, Newtonsoft.Json.Linq.JToken> response in responses)
            {
                for (int j = 0; j < response.Value["Question"].Count(); j++) {
                    listQuestions.Add((string)response.Value["Question"][j]);
                }
            }

            return listQuestions;
        }

        public List<string> parseAnswers(Newtonsoft.Json.Linq.JObject responses)
        {
            List<string> listAnswers = new List<string>();

            foreach (KeyValuePair<string, Newtonsoft.Json.Linq.JToken> response in responses)
            {
                for (int j = 0; j < response.Value["Answer"].Count(); j++)
                {
                    listAnswers.Add((string)response.Value["Answer"][j]);
                }
            }

            return listAnswers;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
