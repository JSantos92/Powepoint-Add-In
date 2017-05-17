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
using System.Threading;

namespace MyRibbonAddIn
{
    [ComVisible(true)]
    public class MyRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        int timerCounter = 0;
        Timer timer = null;
        private ThisAddIn addin;
        string globalUrl;
        List<Question> formQuestions = new List<Question>();
        Excel.Workbook dataWorkbook;
        List<Excel.Workbook>  dataWorkbooks = new List<Excel.Workbook>();
        List<PowerPoint.Chart>  ppCharts = new List<PowerPoint.Chart>();
        object misValue = System.Reflection.Missing.Value;
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
            string formURL = Interaction.InputBox("Insert Google Form URL:", "Google Form URL", "Introduzir URL do Form", 800, 500);

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
            JObject responses = this.addin.getFormResponses(globalUrl);

            List<Question> questions = parseResponses(responses);

            PowerPoint.CustomLayout layout = this.addin.Application.ActivePresentation.SlideMaster.CustomLayouts[6];

            for (int i = 0; i < questions.Count(); i++)
            {
                //Adicionar novo slide e mudar slide de foco

                this.addin.Application.ActivePresentation.Slides.AddSlide(this.addin.currentSlide+1, layout);
                this.addin.Application.ActiveWindow.Presentation.Slides[this.addin.currentSlide+1].Select();
                this.addin.currentSlide++;

                //Output do Titulo da Pergunta 
                this.addin.Application.ActiveWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text = questions[i].Title;

                /*              //Output das opçoes de resposta
                                this.addin.Application.ActiveWindow.View.Slide.Shapes.Placeholders(.Delete(2);
                */
                //Output da opção de resposta 
                //this.addin.Application.ActiveWindow.View.Slide.Shapes.Placeholders(2).TextFrame.TextRange.Text = " ";

                //Output do Gráfico 

                float width = 500F;
                float left = 230F;

                if (questions[i].Choices.Count > 6)
                {
                    width = 750F;
                    left = 130F;
                }

                else if (questions[i].Choices.Count > 12)
                {
                    width = 1100;
                    left = 80F;
                }

                this.addin.Application.ActiveWindow.View.Slide.Shapes.AddChart(Office.XlChartType.xlColumnClustered, left, 120F, width, 400F);

                //Access the added chart
                PowerPoint.Chart ppChart = this.addin.Application.ActiveWindow.View.Slide.Shapes[2].Chart;               

                //Access the chart data
                PowerPoint.ChartData chartData = ppChart.ChartData;

                //Create instance to Excel workbook to work with chart data
                dataWorkbook = (Excel.Workbook)chartData.Workbook;

                dataWorkbooks.Add(dataWorkbook);

                dataWorkbook.Windows[1].WindowState = Excel.XlWindowState.xlMinimized;

                //Accessing the data worksheet for chart
                Excel.Worksheet dataSheet = ((Excel.Worksheet)dataWorkbook.Worksheets[1]);
                               
                //Setting the range of chart

                string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H" };
                string lowerRange = "1";
                int upRange = questions[i].Choices.Count() + 1;
                string upperRange = upRange.ToString();

                lowerRange = "A" + lowerRange;

                if (questions[i].Type == "GRID")
                {
                    upperRange = "F" + upRange/2;
                    int aux = 0;
                    string categoria1 = questions[i].Choices[0].row;

                    for (int y=0; y < questions[i].Choices.Count; y++)
                    {
                        if (questions[i].Choices[y].row == categoria1)
                            aux++;

                    }

                    upperRange = rangeids[aux] + (questions[i].Choices.Count/2-1).ToString();
                    
                }

                else
                    upperRange = "B" + upperRange;

                Excel.Range tRange = dataSheet.Cells.get_Range(lowerRange, upperRange);

                //Applying the set range on chart data table
                Excel.ListObject tbl1 = dataSheet.ListObjects["Tabela1"];
                tbl1.Resize(tRange);

                //Setting values for categories and respective series data

                string option = "A";
                string count = "B";
        
                List<string> rows = new List<string>();

                if (questions[i].Type == "GRID")
                {
                    int iteraux = 2;
                    
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        
                        string optionaux = option + iteraux;

                        if(!rows.Contains(questions[i].Choices[j].row))
                        {
                            dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = questions[i].Choices[j].row;
                            rows.Add(questions[i].Choices[j].row);
                            iteraux++;
                        }

                        
                    }
                    
                    string[] columnids = { "A", "B", "C", "D", "E", "F", "G", "H" };
                    int column_index = 0;
                    bool series_written = false;

                    for (int v = 0; v < rows.Count; v++)
                    {
                        string categoria = rows[v];
                        column_index = 0;
                        for (int u = 0; u < questions[i].Choices.Count; u++)
                        {
                            if(questions[i].Choices[u].row == categoria)
                            {
                                string celula = columnids[column_index + 1] + (v+2).ToString();
                                dataSheet.Cells.get_Range(celula).FormulaR1C1 = questions[i].Choices[u].count.ToString();
                                if (!series_written)
                                    dataSheet.Cells.get_Range(columnids[column_index+1] + (v+1).ToString()).FormulaR1C1 = questions[i].Choices[u].option;
                                column_index++;

                            }
                        }
                        series_written = true;
                    }

                }

                else
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = questions[i].Choices[j].option;

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();

                    }

                    dataSheet.Cells.get_Range("A1").FormulaR1C1 = "";
                }

              
                //((Excel.Range)(dataSheet.Cells.get_Range("B1"))).FormulaR1C1 = null;

                //Setting chart title

                if (questions[i].Type != "GRID")
                    ppChart.ChartTitle.Delete();


                // Insert graphic in Excel

                Excel.Range chartRange;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)dataSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = xlCharts.Add(50, 150, 300, 250);
                Excel.Chart chartPage = myChart.Chart;

                chartRange = dataSheet.get_Range(lowerRange, upperRange);
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                Excel.Axis excelaxis = chartPage.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                int maxScale = 0;
                for (int h = 0; h < questions[i].Choices.Count; h++)
                {
                    if (questions[i].Choices[h].count > maxScale)
                        maxScale = questions[i].Choices[h].count;
                }

                excelaxis.MajorUnit = (int)(maxScale + 10.0) / 5;
                excelaxis.MinorUnit = (int)(maxScale + 10.0) / 10;
                excelaxis.MinimumScale = 0;
                excelaxis.MaximumScale = maxScale + 10.0;

                //Accessing Chart value axis
                PowerPoint.Axis valaxis = ppChart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary);

                //Setting values axis units

                valaxis.MajorUnit = excelaxis.MajorUnit;
                valaxis.MinorUnit = excelaxis.MinorUnit;
                valaxis.MinimumScale = excelaxis.MinimumScale;
                valaxis.MaximumScale = excelaxis.MaximumScale;

                //Accessing Chart Depth axis
                //PowerPoint.Axis Depthaxis = ppChart.Axes(PowerPoint.XlAxisType.xlSeriesAxis, PowerPoint.XlAxisGroup.xlPrimary);
                //Depthaxis.Delete();

                //Setting chart rotation
                //ppChart.Rotation = 0; //Y-Value
                //ppChart.Elevation = 0; //X-Value
                //ppChart.RightAngleAxes = false; 


               


                // Live Reload
                /*
                                 timer = new Timer((e) =>
                                 {
                                     refreshButton(control);
                                     timerCounter++;
                                 }, null, 0, Convert.ToInt32(TimeSpan.FromSeconds(5).TotalMilliseconds));

                    */

                //Copy chart to ppcharts

                ppCharts.Add(ppChart);


            } 
            

        }


        public void enableForm(Office.IRibbonControl control)
        {

            this.addin.setAcceptingResponses(globalUrl,true);
        }


        public void disableForm(Office.IRibbonControl control)
        {
            this.addin.setAcceptingResponses(globalUrl,false);
        }

        public void refreshButton(Office.IRibbonControl control)
        {

            if (timerCounter > 20)
                timer.Dispose();

            JObject responses = this.addin.getFormResponses(globalUrl);

            List<Question> questions = parseResponses(responses);

            int i = 0;

            foreach (PowerPoint.Chart ppChart in ppCharts)
               {

                //Access the chart data
                PowerPoint.ChartData chartData = ppChart.ChartData;

                //Create instance to Excel workbook to work with chart data
                Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;

                //Accessing the data worksheet for chart
                Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];

                //Setting the range of chart

                string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H" };
                string lowerRange = "1";
                int upRange = questions[i].Choices.Count() + 1;
                string upperRange = upRange.ToString();

                lowerRange = "A" + lowerRange;

                if (questions[i].Type == "GRID")
                {
                    upperRange = "F" + upRange / 2;
                    int aux = 0;
                    string categoria1 = questions[i].Choices[0].row;

                    for (int y = 0; y < questions[i].Choices.Count; y++)
                    {
                        if (questions[i].Choices[y].row == categoria1)
                            aux++;

                    }

                    upperRange = rangeids[aux] + (questions[i].Choices.Count / 2 - 1).ToString();

                }

                else
                    upperRange = "B" + upperRange;

                Excel.Range tRange = dataSheet.Cells.get_Range(lowerRange, upperRange);

                //Applying the set range on chart data table
                Excel.ListObject tbl1 = dataSheet.ListObjects["Tabela1"];
                tbl1.Resize(tRange);

                //Setting values for categories and respective series data

                string option = "A";
                string count = "B";

                List<string> rows = new List<string>();

                if (questions[i].Type == "GRID")
                {
                    int iteraux = 2;

                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {

                        string optionaux = option + iteraux;

                        if (!rows.Contains(questions[i].Choices[j].row))
                        {
                            dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = questions[i].Choices[j].row;
                            rows.Add(questions[i].Choices[j].row);
                            iteraux++;
                        }


                    }

                    string[] columnids = { "A", "B", "C", "D", "E", "F", "G", "H" };
                    int column_index = 0;
                    bool series_written = false;

                    for (int v = 0; v < rows.Count; v++)
                    {
                        string categoria = rows[v];
                        column_index = 0;
                        for (int u = 0; u < questions[i].Choices.Count; u++)
                        {
                            if (questions[i].Choices[u].row == categoria)
                            {
                                string celula = columnids[column_index + 1] + (v + 2).ToString();
                                dataSheet.Cells.get_Range(celula).FormulaR1C1 = questions[i].Choices[u].count.ToString();
                                if (!series_written)
                                    dataSheet.Cells.get_Range(columnids[column_index + 1] + (v + 1).ToString()).FormulaR1C1 = questions[i].Choices[u].option;
                                column_index++;

                            }
                        }
                        series_written = true;
                    }

                }

                else
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = questions[i].Choices[j].option;

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();

                    }

                    dataSheet.Cells.get_Range("A1").FormulaR1C1 = "";
                }



                //((Excel.Range)(dataSheet.Cells.get_Range("B1"))).FormulaR1C1 = null;

                
                //Accessing Chart value axis
                PowerPoint.Axis valaxis = ppChart.Axes(PowerPoint.XlAxisType.xlValue, PowerPoint.XlAxisGroup.xlPrimary);

                //Setting values axis units
                int maxScale = 0;
                for (int h = 0; h < questions[i].Choices.Count; h++)
                {
                    if (questions[i].Choices[h].count > maxScale)
                        maxScale = questions[i].Choices[h].count;
                }
                valaxis.MajorUnit = (int)(maxScale + 10.0) / 5;
                valaxis.MinorUnit = (int)(maxScale + 10.0) / 10;
                valaxis.MinimumScale = 0;
                valaxis.MaximumScale = maxScale + 10.0; ;

                //Accessing Chart Depth axis
                //PowerPoint.Axis Depthaxis = ppChart.Axes(PowerPoint.XlAxisType.xlSeriesAxis, PowerPoint.XlAxisGroup.xlPrimary);
                //Depthaxis.Delete();

                //Setting chart rotation
                //ppChart.Rotation = 0; //Y-Value
                //ppChart.Elevation = 0; //X-Value
                //ppChart.RightAngleAxes = false; 

                i++;
            }

        }

        public void newTry(Office.IRibbonControl control)
        {
            //Limpar as respostas existentes no Inquérito

            //this.addin.deleteResponses(globalUrl);




        }

        public List<Question> parseResponses(JObject responses)
        {
            List<Question> listQuestions = new List<Question>();

            foreach (KeyValuePair<string, JToken> response in responses)
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

                if (questao.Type == "GRID")
                {
                    foreach (JProperty row in choices.Properties())
                    {
                        var name = row.Value;

                        foreach (JProperty ans in row.Value)
                        {
                            if (ans.Name != "null") {
                            Choice c = new Choice(row.Name, ans.Name, ans.Value.ToObject<int>());
                            questao.Choices.Add(c);
                        }
                        }

                    }
                }

                else
                {
                    foreach (JProperty prop in choices.Properties())
                    {
                        Choice c = new Choice("", prop.Name, prop.Value.ToObject<int>());
                        questao.Choices.Add(c);
                    }
                }

                listQuestions.Add(questao);
            }

            formQuestions = listQuestions;
            return listQuestions;
        }

        public List<Excel.Workbook> getWorkbooks()
        {
            return dataWorkbooks;
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
