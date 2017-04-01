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
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
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
            string formURL = Interaction.InputBox("Insert Google Form URL:", "Google Form URL", "https://docs.google.com/forms/d/11Vnlhtcw_kvjAB5QZFrYXiRUiu5Imlix-H8XOOEp9Vs/edit", 100, 200);
            Newtonsoft.Json.Linq.JObject responses = this.addin.getFormResponses(formURL);

            /*PowerPoint.Shape textBox = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddTextbox(
                 Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
             textBox.TextFrame.TextRange.InsertAfter(responses.ToString());*/

            // Generating TinyUrl Link and adding it to current slide

            Uri address = new Uri("http://tinyurl.com/api-create.php?url=" + formURL);
            System.Net.WebClient client = new System.Net.WebClient();
            string tinyUrl = client.DownloadString(address);
            Console.WriteLine(tinyUrl);

            PowerPoint.Shape textBox = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0, 500, 50);
            textBox.TextFrame.TextRange.InsertAfter(tinyUrl.ToString());

            // Generating QrCode and adding it to current slide

            QRCodeGenerator qrGenerator = new QRCodeGenerator();
            QRCodeData qrCodeData = qrGenerator.CreateQrCode(tinyUrl, QRCodeGenerator.ECCLevel.Q);
            QRCode qrCode = new QRCode(qrCodeData);
            Bitmap qrCodeImage = qrCode.GetGraphic(20);

            qrCodeImage.Save("./qrcode.bmp");

            Microsoft.Office.Interop.PowerPoint.Shape ppPicture = this.addin.Application.ActiveWindow.View.Slide.Shapes.AddPicture("./qrcode.bmp", Office.MsoTriState.msoTrue, Office.MsoTriState.msoTrue, 50, 50, 200, 200);

            //PowerPoint.Shape qrCodeImg = this.addin

            //PowerPoint.Shape qrCodePower = this.addin.Application.ActiveWindow.View.Slide.AddPicture("./qrcode.bmp",Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse, 50,50,200,200);
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
