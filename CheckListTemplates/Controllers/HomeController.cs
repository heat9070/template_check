using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace CheckListTemplates.Controllers
{
    public class HomeController : Controller
    {
        void ChangeChartPart(ChartPart chartPart1)
        {
            ChartSpace chartSpace1 = chartPart1.ChartSpace;

            DocumentFormat.OpenXml.Drawing.Charts.Chart chart1 = chartSpace1.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();

            PlotArea plotArea1 = chart1.GetFirstChild<PlotArea>();

            BarChart barChart1 = plotArea1.GetFirstChild<BarChart>();

            //BarChartSeries barChartSeries1 = barChart1.Elements<BarChartSeries>().ElementAtOrDefault(2);
            for (int i = 0; i < barChart1.Elements<BarChartSeries>().Count(); i++)
            {
                BarChartSeries barChartSeries = barChart1.Elements<BarChartSeries>().ElementAtOrDefault(i);
                if (barChartSeries != null)
                {

                    Values values1 = barChartSeries.GetFirstChild<Values>();

                    NumberReference numberReference1 = values1.GetFirstChild<NumberReference>();

                    NumberingCache numberingCache1 = numberReference1.GetFirstChild<NumberingCache>();

                    NumericPoint numericPoint1 = numberingCache1.GetFirstChild<NumericPoint>();
                    NumericPoint numericPoint2 = numberingCache1.Elements<NumericPoint>().ElementAt(1);
                    NumericPoint numericPoint3 = numberingCache1.Elements<NumericPoint>().ElementAt(2);

                    NumericValue numericValue1 = numericPoint1.GetFirstChild<NumericValue>();
                    numericValue1.Text = ".10";


                    NumericValue numericValue2 = numericPoint2.GetFirstChild<NumericValue>();
                    numericValue2.Text = ".10";


                    NumericValue numericValue3 = numericPoint3.GetFirstChild<NumericValue>();
                    numericValue3.Text = ".80";


                }
            }

            chartSpace1.Save();
        }

        public ActionResult Index()
        {
            PresentationDocument oPDoc = PresentationDocument.Open(Server.MapPath("~/App_Data/TDM.pptx"), true);
            PresentationPart oPPart = oPDoc.PresentationPart;
            SlidePart sectionSlidePart = (SlidePart)oPPart.GetPartById("rId3");
            SlidePart sectionSlidePart2 = (SlidePart)oPPart.GetPartById("rId4");
            SlidePart sectionSlidePart3 = (SlidePart)oPPart.GetPartById("rId8");
            var tbl = sectionSlidePart3.Slide.Descendants<Table>().First();
            //var tr1 = tbl.Descendants<TableRow>().ElementAtOrDefault(2);
            //var cl1 = tr1.Descendants<TableCell>().FirstOrDefault();
            //DocumentFormat.OpenXml.Drawing.TextBody tb = cl1.Elements<DocumentFormat.OpenXml.Drawing.TextBody>().First();
            //Paragraph p = tb.Elements<Paragraph>().ElementAtOrDefault(0);
            //Run r = p.Elements<Run>().First();
            //DocumentFormat.OpenXml.Drawing.Text t = r.Elements<DocumentFormat.OpenXml.Drawing.Text>().First();
            //t.Text = "gendy101";

            //var tr = new TableRow();
            //TableCell tc1 = CreateTextCell("hi1");
            //tr.Append(tc1);
            //tc1 = CreateTextCell("hi2");
            //tr.Append(tc1);
            //tc1 = CreateTextCell("hi3");
            //tr.Append(tc1);
            //tc1 = CreateTextCell("hi114");
            //tr.Append(tc1);
            //tc1 = CreateTextCell("hi5");
            //tr.Append(tc1);
            //tc1 = CreateTextCell("hi6");
            //tr.Append(tc1);
            //tbl.Append(tr);

            foreach (var chartPart1 in sectionSlidePart2.ChartParts)
            {
                try
                {
                    ChangeChartPart(chartPart1);

                }
                catch (Exception ex)
                {

                }

            }
            //SlideLayoutPart slideLayoutPart = (SlideLayoutPart)sectionSlidePart.GetPartById("rId1");
            //DocumentFormat.OpenXml.Drawing.TextBody textBody1 = sectionSlidePart.SlideParts.FirstOrDefault().Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().First();
            //SlidePart slide1 = GetFirstSlide(oPDoc);
            //DocumentFormat.OpenXml.Drawing.Shape textBody1 = slide1.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Shape>().First();
            //DocumentFormat.OpenXml.Drawing.TextBody textBody1 = sectionSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().First();
            foreach (var item in sectionSlidePart.Slide.Descendants())
            {
                if (item.GetType() == typeof(CommonSlideData))
                {
                    CommonSlideData sldData1 = (CommonSlideData)item;
                    //var shapes1 = sldData1.Descendants<DocumentFormat.OpenXml.Drawing.Shape>();
                    //var shapes2 = sldData1.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().ToList();
                    //int count1 = shapes1.Count();

                    ShapeTree tree = sldData1.ShapeTree;
                    foreach (DocumentFormat.OpenXml.Presentation.Shape shape in tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                    {

                        // Run through all the paragraphs in the document
                        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                        {
                            foreach (Run run in paragraph.Elements<Run>())
                            {
                                if (run.Text.InnerText.Contains("Mohamed Elgendy"))
                                {
                                    run.Text = new DocumentFormat.OpenXml.Drawing.Text("Mohamed Elgendy2");

                                }
                            }

                        }
                    }
                    sectionSlidePart.Slide.Save();
                }

            }
            sectionSlidePart2.Slide.Save();
            sectionSlidePart3.Slide.Save();
            //oPPart.Presentation.Save();
            oPDoc.PresentationPart.Presentation.Save();
            foreach (var slideMasterPart in oPDoc.PresentationPart.SlideMasterParts)
            {
                slideMasterPart.SlideMaster.Save();
            }

            //System.IO.FileStream f = new FileStream(@"E:\test111.pptx", FileMode.Create);
            //oPPart.GetStream().CopyTo(f);
            //f.Close();


            //ImagePart imagePart = (ImagePart)sectionSlidePart.GetPartById("rId3");
            //if (imagePart != null)
            //{
            //    using (FileStream fileStream = new FileStream(Server.MapPath("~/App_Data/cup.png"), FileMode.Open))
            //    {
            //        imagePart.FeedData(fileStream);
            //        fileStream.Close();
            //    }
            //}
            //ImagePart imagePart2 = (ImagePart)sectionSlidePart.GetPartById("rId4");
            //sectionSlidePart.Slide.Save();



            //XmlDocument doc = new XmlDocument();
            //doc.Load(slideLayoutPart.GetStream());
            //doc.Save("e:\\101.xml");

            //System.Drawing.Image img = System.Drawing.Image.FromStream(imagePart2.GetStream());
            //img.Save(@"E:\temp202.jpg");








            return View();
        }
        private static TableCell CreateTextCell(string text)
        {
            var textCol = new string[2];
            if (!string.IsNullOrEmpty(text))
            {
                if (text.Length > 25)
                {
                    textCol[0] = text.Substring(0, 25);
                    textCol[1] = text.Substring(26);
                }
                else
                {
                    textCol[0] = text;
                }
            }
            else
            {
                textCol[0] = string.Empty;
            }


            TableCell tableCell3 = new TableCell();
            //DocumentFormat.OpenXml.Drawing.TextBody textBody1 = tableCell3.GetFirstChild<DocumentFormat.OpenXml.Drawing.TextBody>();

            DocumentFormat.OpenXml.Drawing.TextBody textBody3 = new DocumentFormat.OpenXml.Drawing.TextBody();
            //BodyProperties bodyProperties3 = new BodyProperties();

            //ListStyle listStyle3 = new ListStyle();

            //textBody3.Append(bodyProperties3);
            //textBody3.Append(listStyle3);


            var nonNull = textCol.Where(t => !string.IsNullOrEmpty(t)).ToList();

            foreach (var textVal in nonNull)
            {
                //if (!string.IsNullOrEmpty(textVal))
                //{
                Paragraph paragraph3 = new Paragraph();
                Run run2 = new Run();
                RunProperties runProperties2 = new RunProperties() { Language = "en-US", Underline = TextUnderlineValues.DashHeavy };
                DocumentFormat.OpenXml.Drawing.Text text2 = new DocumentFormat.OpenXml.Drawing.Text();
                text2.Text = textVal;
                run2.Append(runProperties2);
                run2.Append(text2);
                paragraph3.Append(run2);
                textBody3.Append(paragraph3);
                //}
            }

            TableCellProperties tableCellProperties3 = new TableCellProperties();
            tableCell3.Append(textBody3);
            tableCell3.Append(tableCellProperties3);



            //var tc = new A.TableCell(
            //                    new A.TextBody(
            //                        new A.BodyProperties(),
            //                    new A.Paragraph(
            //                        new A.Run(
            //                            new A.Text(text)))),
            //                    new A.TableCellProperties());

            //return tc;
            return tableCell3;
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}