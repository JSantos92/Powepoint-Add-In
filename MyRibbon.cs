﻿using Microsoft.VisualBasic;
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
        string gridUpperRangeGlobal = "";
        string lowerRangeGlobal = "A";
        string upperRangeGlobal = "B";
        System.Threading.Timer timer;
        int newTryCount = 0;
        bool newTryBool = false;
        private Office.IRibbonUI ribbon;
        private ThisAddIn addin;
        string globalUrl;
        List<Question> formQuestions = new List<Question>();
        Excel.Workbook dataWorkbook = null;
        List<Excel.Workbook>  dataWorkbooks = new List<Excel.Workbook>();
        List<Excel.ChartObject>  xlsCharts = new List<Excel.ChartObject>();
        List<Excel.ChartObject> xlsChartsTemp = new List<Excel.ChartObject>();
        object misValue = System.Reflection.Missing.Value;
        PowerPoint.ShapeRange shapeRange = null;
        PowerPoint.Slide pptSlide = null;
        List<PowerPoint.Slide> pptSlides = new List<PowerPoint.Slide>();
        List<Excel.Worksheet> dataWorkSheets = new List<Excel.Worksheet>();
        object paramMissing = Type.Missing;

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
            string fileName = this.addin.getTitle(globalUrl);
            string paramWorkbookPath = @".\" + fileName + ".xlsx";
            string paramPowerpointPath = @".\" + fileName + ".pptx";

            List<Question> questions = parseResponses(responses);

            PowerPoint.CustomLayout layout = this.addin.Application.ActivePresentation.SlideMaster.CustomLayouts[6];

            for (int i = 0; i < questions.Count(); i++)
            {
                //Adicionar novo slide e mudar slide de foco

                pptSlide = this.addin.Application.ActivePresentation.Slides.AddSlide(this.addin.currentSlide+1, layout);

                pptSlides.Add(pptSlide);
                this.addin.Application.ActiveWindow.Presentation.Slides[this.addin.currentSlide+1].Select();
                this.addin.currentSlide++;

                //Output do Titulo da Pergunta 
                this.addin.Application.ActiveWindow.View.Slide.Shapes.Title.TextFrame.TextRange.Text = questions[i].Title;

                Excel.Worksheet dataSheet;

                if (i == 0)
                {

                    //Create instance to Excel workbook to work with chart data
                    Excel.Application excelApp = new Excel.Application();
                    excelApp.Visible = true;

                    if (!File.Exists(paramWorkbookPath))
                        {
                            dataWorkbook = excelApp.Workbooks.Add();

                            dataWorkbooks.Add(dataWorkbook);

                            dataWorkbook.Windows[1].WindowState = Excel.XlWindowState.xlMinimized;

                            //Accessing the data worksheet for chart
                            dataSheet = ((Excel.Worksheet)dataWorkbook.Worksheets[i + 1]);

                            dataWorkSheets.Add(dataSheet);

                        }

                    else
                        {
                            dataWorkbook = excelApp.Workbooks.Open(paramWorkbookPath);

                            dataWorkbooks.Add(dataWorkbook);

                            dataWorkbook.Windows[1].WindowState = Excel.XlWindowState.xlMinimized;

                            //Accessing the data worksheet for chart
                            dataSheet = ((Excel.Worksheet)dataWorkbook.Worksheets[i + 1]);

                            dataWorkSheets.Add(dataSheet);
                         
                        }

                }

                else
                {
                    dataSheet = dataWorkbook.Worksheets.Add();
                    dataWorkSheets.Add(dataSheet);

                }

                               
                //Setting the range of chart

                string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                string lowerRange = "1";
                int upRange = questions[i].Choices.Count() + 1;
                string upperRange = upRange.ToString();

                lowerRange = "A" + lowerRange;

                //TODO TIRAR ESTE "F" HARDCODED

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
                    gridUpperRangeGlobal = upperRange;

                    dataSheet.Cells.get_Range("A50").FormulaR1C1 = upperRange;

                }

                else
                {
                    upperRange = "B" + upperRange;
                    dataSheet.Cells.get_Range("A50").FormulaR1C1 = upperRange;
                }


                Excel.Range tRange = dataSheet.Cells.get_Range(lowerRange, upperRange);

                Excel.ListObject tbl1;

                string tableName = "Tabela";

                //Applying the set range on chart data table

                tableName = tableName + i;
                dataSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange,Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = tableName;
                tbl1 = dataSheet.ListObjects[tableName];
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


                    dataSheet.Cells.get_Range("A1").FormulaR1C1 = "Try1";

                }

                else if (questions[i].Type == "SCALE")
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = "- " + questions[i].Choices[j].option.ToString() + " -";

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();

                    }


                    dataSheet.Cells.get_Range("A1").FormulaR1C1 = "Categoria";
                    dataSheet.Cells.get_Range("B1").FormulaR1C1 = "Try1";

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


                    dataSheet.Cells.get_Range("A1").FormulaR1C1 = "Categoria";
                    dataSheet.Cells.get_Range("B1").FormulaR1C1 = "Try1";

                }



                // Insert graphic in Excel

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

                //Save value of last 

                Excel.Range chartRange;
                Excel.ChartObjects xlCharts = (Excel.ChartObjects)dataSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = xlCharts.Add(50, 150, width, 250);
                Excel.Chart chartPage = myChart.Chart;

                object paramMissing = Type.Missing;

                // Declare variables for the Chart.ChartWizard method.
                object paramChartFormat = 1;
                object paramCategoryLabels = 0;
                object paramSeriesLabels = 0;
                bool paramHasLegend = true;
         


                // Create a new chart of the data.
                myChart.Chart.ChartWizard(tRange, Excel.XlChartType.xlColumnClustered, paramChartFormat, Excel.XlRowCol.xlRows,
                    paramCategoryLabels, paramSeriesLabels, paramHasLegend, paramMissing, paramMissing, paramMissing, paramMissing);

                chartRange = dataSheet.get_Range(lowerRange, upperRange);
                chartPage.SetSourceData(chartRange, misValue);
                chartPage.ChartType = Excel.XlChartType.xlColumnClustered;

                if (questions[i].Type != "GRID")
                    chartPage.ChartTitle.Delete();


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

                myChart.Copy();

                shapeRange = pptSlide.Shapes.Paste();

                // Position the chart on the slide.



                shapeRange.Left = left;
                shapeRange.Top = 160F;
                shapeRange.Height = 350F;
                shapeRange.Width = width;



                //Copy chart to xlsCharts

                xlsCharts.Add(myChart);

            }

            // Save excel file

            if(newTryBool == true)
            {
                //dataWorkbook.Save();
            }

            else
            {
                //dataWorkbook.SaveAs(paramWorkbookPath, paramMissing, paramMissing, paramMissing, paramMissing,
            //paramMissing, Excel.XlSaveAsAccessMode.xlNoChange, paramMissing, paramMissing, paramMissing, paramMissing, paramMissing);
            }


            pptSlides[0].Select();


            // Save the presentation.
            //this.addin.Application.ActivePresentation.SaveAs(paramPowerpointPath, PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Office.MsoTriState.msoTrue);
        }


        public void enableForm(Office.IRibbonControl control)
        {

            this.addin.setAcceptingResponses(globalUrl,true);
        }


        public void disableForm(Office.IRibbonControl control)
        {
            this.addin.setAcceptingResponses(globalUrl,false);
        }

        public void startUpdate(Office.IRibbonControl control)
        {
            startLiveUpdate();
        }

        public void stopUpdate(Office.IRibbonControl control)
        {
            stopLiveUpdate();
        }

        public void startLiveUpdate()
        {
            timer = new Timer(
            e => liveUpdate(),
            null,
            TimeSpan.Zero,
            TimeSpan.FromMinutes(0.15));
        }

        public void stopLiveUpdate()
        {
            timer.Change(-1,-1);
        }

        public void refreshButton(Office.IRibbonControl control)
        {

            refresh();

        }

        public void liveUpdate()
        {

            string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };

            JObject responses = this.addin.getFormResponses(globalUrl);

            List<Question> questions = parseResponses(responses);

            int i = 0;

            foreach (Excel.Worksheet dataSheet in dataWorkSheets)
            {
                int lastCellUsed = 0;

                //Descobrir o initial range

                for (int x = 0; x < 70; x++)
                {
                    if (dataSheet.Cells[2, x + 1].Value != null)
                    {
                        lastCellUsed = x;
                    }
                }

                //Setting values for categories and respective series data

                string option = lowerRangeGlobal;
                string count = upperRangeGlobal;
                List<string> rows = new List<string>();

                if (questions[i].Type == "GRID")
                {
                    if (newTryCount != 0)
                        option = lowerRangeGlobal + 2;
                    else
                        option = lowerRangeGlobal;

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

                    string[] columnids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                    int column_index = 0;
                    bool series_written = false;

                    for (int v = 0; v < rows.Count; v++)
                    {
                        string categoria = rows[v];
                        if (newTryCount != 0)
                            column_index = Array.IndexOf(columnids, lowerRangeGlobal)+2*newTryCount;
                        else
                            column_index = Array.IndexOf(columnids,lowerRangeGlobal);

                        for (int u = 0; u < questions[i].Choices.Count; u++)
                        {
                            if (questions[i].Choices[u].row == categoria)
                            {
                                string celula = columnids[column_index + 1] + (v + 2).ToString();
                                dataSheet.Cells.get_Range(celula).FormulaR1C1 = questions[i].Choices[u].count.ToString();
                                if (!series_written)
                                {
                                    try
                                    {
                                        dataSheet.Cells.get_Range(columnids[column_index + 1] + (v + 1).ToString()).FormulaR1C1 = questions[i].Choices[u].option;
                                    }
                                    catch (COMException e)
                                    {
                                        Thread.Sleep(500);
                                        u--;
                                    }
                                    

                                }
                                    
                                column_index++;

                            }
                        }
                        series_written = true;
                    }

                }


                else if (questions[i].Type == "SCALE")
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        try
                        {
                            dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = "- " + questions[i].Choices[j].option.ToString() + " -";
                        }
                        catch (COMException e)
                        {
                            Thread.Sleep(500);
                            j--;
                        }
                        

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        try
                        {
                            dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();
                        }
                        catch (COMException e)
                        {
                            Thread.Sleep(500);
                            k--;
                        }
                        

                    }

                }


                else
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        try
                        {
                            dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = questions[i].Choices[j].option;
                        } catch(COMException e)
                        {
                            Thread.Sleep(500);
                            j--;
                        }

                    }


                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        try
                        {
                            dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count;
                        }
                        catch (COMException e)
                        {
                            Thread.Sleep(500);
                            k--;
                        }
                        

                    }

                }

                i++;
            }
            
        }

        public void refresh()
        {
            JObject responses = this.addin.getFormResponses(globalUrl);

            List<Question> questions = parseResponses(responses);

            int i = 0;

            foreach (Excel.ChartObject xlsChart in xlsCharts)
            {

                Excel.Worksheet dataSheet = dataWorkSheets[i];

                //Setting the range of chart

                string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                string lowerRange = "1";
                int upRange = questions[i].Choices.Count() + 1;
                string upperRange = upperRangeGlobal;

                int lastCellUsed = 0;

                lowerRange = rangeids[lastCellUsed + 2] + lowerRange;

                //Descobrir o initial range

                for (int x = 0; x < 70; x++)
                {
                    if (dataSheet.Cells[2, x + 1].Value != null)
                    {
                        lastCellUsed = x;
                    }
                }

                if (questions[i].Type == "GRID")
                {

                    int aux = 0;
                    string categoria1 = questions[i].Choices[0].row;

                    for (int y = 0; y < questions[i].Choices.Count; y++)
                    {
                        if (questions[i].Choices[y].row == categoria1)
                            aux++;

                    }

                    upperRange = rangeids[aux + lastCellUsed + 2] + (questions[i].Choices.Count / 2 - 1).ToString();

                }

                else
                    upperRange = rangeids[lastCellUsed + 3] + upRange;

                Excel.Range tRange = dataSheet.Cells.get_Range(lowerRange, upperRange);

                //Setting values for categories and respective series data

                string option = lowerRangeGlobal;
                string count = upperRangeGlobal;

                List<string> rows = new List<string>();

                if (questions[i].Type == "GRID")
                {

                    if (newTryCount != 0)
                        option = lowerRangeGlobal + 2;
                    else
                        option = lowerRangeGlobal;

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

                    string[] columnids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                    int column_index = 0;
                    bool series_written = false;

                    for (int v = 0; v < rows.Count; v++)
                    {
                        string categoria = rows[v];
                        column_index = Array.IndexOf(columnids, lowerRangeGlobal) + 2 * newTryCount;
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

                else if (questions[i].Type == "SCALE")
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = "- " + questions[i].Choices[j].option.ToString() + " -";

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();

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

                }

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

                Excel.Axis excelaxis = xlsChart.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);

                if (xlsChart.Chart.ChartType == Excel.XlChartType.xlColumnClustered)
                {

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
                }

                PowerPoint.Slide currentSlide = this.addin.Application.ActiveWindow.View.Slide;

                pptSlides[i].Select();

                xlsChart.Copy();

                shapeRange = pptSlides[i].Shapes.Paste();


                PowerPoint.Shape previousGraph = pptSlides[i].Shapes[2];

                previousGraph.Delete();

                // Position the chart on the slide.

                shapeRange.Left = left;
                shapeRange.Top = 160F;
                shapeRange.Height = 350F;
                shapeRange.Width = width;

                i++;

                currentSlide.Select();
            }
        }

        public void newTry(Office.IRibbonControl control)
        {
            //Limpar as respostas existentes no Inquérito

            //this.addin.deleteResponses(globalUrl);

            JObject responses = this.addin.getFormResponses(globalUrl);

            List<Question> questions = parseResponses(responses);

            int i = 0;

            foreach (Excel.ChartObject xlsChart in xlsCharts)
            {
                xlsChart.Delete();

                Excel.Worksheet dataSheet = dataWorkSheets[i];

                //Setting the range of chart

                string[] rangeids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                string lowerRange = "1";
                int upRange = questions[i].Choices.Count() + 1;
                string upperRange = upRange.ToString();
                int lastCellUsed = 0;

                //Descobrir o initial range

                for (int x=0; x < 70; x++)
                {
                    if(dataSheet.Cells[2, x+1].Value != null)
                    {
                        lastCellUsed = x;
                    }
                }

                


                lowerRange = rangeids[lastCellUsed+2] + lowerRange;
                lowerRangeGlobal = lowerRange[0].ToString();

                if (questions[i].Type == "GRID")
                {
                    int aux = 0;
                    string categoria1 = questions[i].Choices[0].row;

                    for (int y = 0; y < questions[i].Choices.Count; y++)
                    {
                        if (questions[i].Choices[y].row == categoria1)
                            aux++;

                    }

                    upperRange = rangeids[aux+lastCellUsed+2] + (questions[i].Choices.Count / 2 - 1).ToString();
                    gridUpperRangeGlobal = upperRange[0].ToString();

                }

                else
                {
                    upperRange = rangeids[lastCellUsed + 3] + upperRange;
                    upperRangeGlobal = upperRange[0].ToString();
                }
                    
                    

                Excel.Range tRange = dataSheet.Cells.get_Range(lowerRange, upperRange);

                Excel.ListObject tbl1;

                string tableName = "Tabela";

                //Applying the set range on chart data table

                DateTime date1 = DateTime.Now;

                dataSheet.Cells.get_Range("A50").FormulaR1C1 = tRange;

                tableName = date1.ToString();

                dataSheet.Cells.get_Range("A51").FormulaR1C1 = tableName;

                dataSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tRange, Type.Missing, Excel.XlYesNoGuess.xlYes, Type.Missing).Name = tableName;
                tbl1 = dataSheet.ListObjects[tableName];
                tbl1.Resize(tRange);

                //Setting values for categories and respective series data

                string option = rangeids[lastCellUsed + 2];
                string count = rangeids[lastCellUsed + 3];

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

                    string[] columnids = { "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
                    int column_index = 0;
                    bool series_written = false;

                    for (int v = 0; v < rows.Count; v++)
                    {
                        string categoria = rows[v];
                        column_index = lastCellUsed +2;
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

                    dataSheet.Cells.get_Range(lowerRange).FormulaR1C1 = "Try " + (newTryCount + 2).ToString();


                }

                else if (questions[i].Type == "SCALE")
                {
                    for (var j = 0; j < questions[i].Choices.Count; j++)
                    {
                        int index1 = j + 2;
                        string optionaux = option + index1;
                        dataSheet.Cells.get_Range(optionaux).FormulaR1C1 = "- " + questions[i].Choices[j].option.ToString() + " -";

                    }

                    for (var k = 0; k < questions[i].Choices.Count; k++)
                    {
                        int index2 = k + 2;
                        string countaux = count + index2;
                        dataSheet.Cells.get_Range(countaux).FormulaR1C1 = questions[i].Choices[k].count.ToString();

                    }


                    dataSheet.Cells.get_Range(lowerRange).FormulaR1C1 = "Categoria";
                    string auxUpperRange = upperRange[0] + "1";
                    dataSheet.Cells.get_Range(auxUpperRange).FormulaR1C1 = "Try " + (newTryCount + 2).ToString();

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


                    dataSheet.Cells.get_Range(lowerRange).FormulaR1C1 = "Categoria";
                    string auxUpperRange = upperRange[0] + "1";
                    dataSheet.Cells.get_Range(auxUpperRange).FormulaR1C1 = "Try " + (newTryCount + 2).ToString();

                }


                // Insert graphic in Excel

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

                //Save value of last 

                Excel.Range chartRange;
                Excel.ChartObjects xlCharts = (Excel.ChartObjects)dataSheet.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = xlCharts.Add(50, 150, width, 250);
                Excel.Chart chartPage = myChart.Chart;

                object paramMissing = Type.Missing;

                // Declare variables for the Chart.ChartWizard method.
                object paramChartFormat = 1;
                object paramCategoryLabels = 0;
                object paramSeriesLabels = 0;
                bool paramHasLegend = true;

                

                // Create a new chart of the data.
                myChart.Chart.ChartWizard(tRange, Excel.XlChartType.xlColumnClustered, paramChartFormat, Excel.XlRowCol.xlRows,
                    paramCategoryLabels, paramSeriesLabels, paramHasLegend, paramMissing, paramMissing, paramMissing, paramMissing);
                
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

                if (questions[i].Type != "GRID")
                    chartPage.ChartTitle.Delete();

                excelaxis.MajorUnit = (int)(maxScale + 10.0) / 5;
                excelaxis.MinorUnit = (int)(maxScale + 10.0) / 10;
                excelaxis.MinimumScale = 0;
                excelaxis.MaximumScale = maxScale + 10.0;


                PowerPoint.Slide currentSlide = this.addin.Application.ActiveWindow.View.Slide;

                pptSlides[i].Select();

                try
                {
                    myChart.Copy();
                }

                catch (COMException e)
                {
                    Thread.Sleep(500);
                    myChart.Copy();
                }

                shapeRange = pptSlides[i].Shapes.Paste();


                PowerPoint.Shape previousGraph = pptSlides[i].Shapes[2];

                previousGraph.Delete();

                shapeRange.Left = left;
                shapeRange.Top = 160F;
                shapeRange.Height = 350F;
                shapeRange.Width = width;

                xlsChartsTemp.Add(myChart);

                i++;
            }

            xlsCharts = xlsChartsTemp;
            xlsChartsTemp = new List<Excel.ChartObject>();
            newTryBool = true;
            newTryCount++;
           // dataWorkbook.Save();

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
