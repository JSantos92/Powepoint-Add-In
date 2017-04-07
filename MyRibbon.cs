using Microsoft.VisualBasic;
using Newtonsoft.Json.Linq;
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
using Excel = Microsoft.Office.Interop.Excel;

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

            List<Question> questions = parseResponses(responses);

            PowerPoint.CustomLayout layout = this.addin.Application.ActivePresentation.SlideMaster.CustomLayouts[4];

            for (int i = 0; i < questions.Count(); i++)
            {
                //Adicionar novo slide e mudar slide de foco

                this.addin.Application.ActivePresentation.Slides.AddSlide(this.addin.currentSlide + 1, layout);
                this.addin.Application.ActiveWindow.Presentation.Slides[this.addin.currentSlide + 1].Select();

                //Output do Titulo da Pergunta 
                this.addin.Application.ActiveWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text = questions[i].Title;

/*              //Output das opçoes de resposta
                this.addin.Application.ActiveWindow.View.Slide.Shapes.Placeholders().Delete(2);
*/
                //Output da opção de resposta 
                this.addin.Application.ActiveWindow.View.Slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = " ";

                //Output do Gráfico 
                this.addin.Application.ActiveWindow.View.Slide.Shapes.AddChart(Office.XlChartType.xl3DColumn, 200F, 70F, 600F, 500F);

                //Access the added chart
                PowerPoint.Chart ppChart = this.addin.Application.ActiveWindow.View.Slide.Shapes[3].Chart;

                //Access the chart data
                PowerPoint.ChartData chartData = ppChart.ChartData;

                //Create instance to Excel workbook to work with chart data
                Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;

                //Accessing the data worksheet for chart
                Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];

                //Setting the range of chart
                Excel.Range tRange = dataSheet.Cells.get_Range("A1", "B5");

                //Applying the set range on chart data table
                Excel.ListObject tbl1 = dataSheet.ListObjects["Tabela1"];
                tbl1.Resize(tRange);

                //Setting values for categories and respective series data TODO : for loop para as choices

                ((Excel.Range)(dataSheet.Cells.get_Range("A2"))).FormulaR1C1 = questions[i].Choices[0].option;
                ((Excel.Range)(dataSheet.Cells.get_Range("A3"))).FormulaR1C1 = questions[i].Choices[1].option;
                ((Excel.Range)(dataSheet.Cells.get_Range("A4"))).FormulaR1C1 = questions[i].Choices[2].option;
                ((Excel.Range)(dataSheet.Cells.get_Range("A5"))).FormulaR1C1 = questions[i].Choices[3].option;
                ((Excel.Range)(dataSheet.Cells.get_Range("B2"))).FormulaR1C1 = questions[i].Choices[1].count.ToString();
                ((Excel.Range)(dataSheet.Cells.get_Range("B3"))).FormulaR1C1 = questions[i].Choices[1].count.ToString();
                ((Excel.Range)(dataSheet.Cells.get_Range("B4"))).FormulaR1C1 = questions[i].Choices[2].count.ToString(); 
                ((Excel.Range)(dataSheet.Cells.get_Range("B5"))).FormulaR1C1 = questions[i].Choices[3].count.ToString(); 

                //Setting chart title
               // ppChart.ChartTitle.Font.Italic = true;
                ppChart.ChartTitle.Text = " ";
                /*ppChart.ChartTitle.Font.Size = 18;
                ppChart.ChartTitle.Font.Color = Color.Black.ToArgb();
                ppChart.ChartTitle.Format.Line.Visible = Office.MsoTriState.msoTrue;
                ppChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();*/

                //Accessing Chart value axis
                PowerPoint.Axis valaxis = ppChart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary);

                //Setting values axis units
                valaxis.MajorUnit = 100.0F;
                valaxis.MinorUnit = 50.0F;
                valaxis.MinimumScale = 0.0F;
                valaxis.MaximumScale = 500.0F;

                //Accessing Chart Depth axis
                PowerPoint.Axis Depthaxis = ppChart.Axes(PowerPoint.XlAxisType.xlSeriesAxis, PowerPoint.XlAxisGroup.xlPrimary);
                Depthaxis.Delete();

                //Setting chart rotation
                ppChart.Rotation = 0; //Y-Value
                ppChart.Elevation = 0; //X-Value
                ppChart.RightAngleAxes = false; 

    
                
                // Cálcular numero de votos por opção



                //Output do gráfico das respostas




            }
        



        }

        public List<Question> parseResponses(Newtonsoft.Json.Linq.JObject responses)
        {
            List<Question> listQuestions = new List<Question>();

            foreach (KeyValuePair<string, Newtonsoft.Json.Linq.JToken> response in responses)
            {
                Question questao = new Question();
                questao.Title = (string)response.Value["Title"];

                var responseArray = (JArray)response.Value["Answers"];
                foreach(var r in responseArray)
                {
                    questao.Answers.Add(r.ToObject<string>());
                }
                //questao.Answers = (JArray)response.Value["Answers"].ToArray<string>();
                questao.Type = (string)response.Value["Type"];

                var choices = (JObject)response.Value["Choices"];
                foreach (JProperty prop in choices.Properties())
                {
                    Choice c = new Choice(prop.Name, prop.Value.ToObject<int>());
                    questao.Choices.Add(c);
                }

                listQuestions.Add(questao);
            }

            return listQuestions;
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
