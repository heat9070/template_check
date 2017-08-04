using CheckListTemplates.Models.DTO;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Web;
using System.Web.Mvc;

namespace CheckListTemplates.Controllers
{
    public class HomeController : Controller
    {
        public TDMGuestServices TDMGuestServices { get; set; }

        void setChartData(ChartPart chartPart1, ChartDataHolder[] ChartDataSlide)
        {
            ChartSpace chartSpace = chartPart1.ChartSpace;

            DocumentFormat.OpenXml.Drawing.Charts.Chart chart1 = chartSpace.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>();
            PlotArea plotArea1 = chart1.GetFirstChild<PlotArea>();
            BarChart barChart1 = plotArea1.GetFirstChild<BarChart>();

            //BarChartSeries barChartSeries1 = barChart1.Elements<BarChartSeries>().ElementAtOrDefault(2);
            for (int i = 0; i < barChart1.Elements<BarChartSeries>().Count(); i++)
            {
                BarChartSeries barChartSeries = barChart1.Elements<BarChartSeries>().ElementAtOrDefault(i);
                ChartDataSlide dataModel = ChartDataSlide.ElementAtOrDefault(i).chartData;

                SeriesText seriesText = barChartSeries.Elements<SeriesText>().ElementAtOrDefault(0);
                if (seriesText != null)
                {
                    var stringReference = seriesText.Descendants<StringReference>().FirstOrDefault();
                    var stringCache = stringReference.Descendants<StringCache>().FirstOrDefault();
                    var stringPoint = stringCache.Descendants<StringPoint>().FirstOrDefault();
                    var barLabel = stringPoint.GetFirstChild<NumericValue>();
                    barLabel.Text = ChartDataSlide.ElementAtOrDefault(i).seriesText;
                }

                if (barChartSeries != null)
                {
                    
                    Values values1 = barChartSeries.GetFirstChild<Values>();
                    NumberReference numberReference1 = values1.GetFirstChild<NumberReference>();
                    NumberingCache numberingCache1 = numberReference1.GetFirstChild<NumberingCache>();

                    NumericPoint numericPoint1 = numberingCache1.GetFirstChild<NumericPoint>();
                    NumericPoint numericPoint2 = numberingCache1.Elements<NumericPoint>().ElementAt(1);
                    NumericPoint numericPoint3 = numberingCache1.Elements<NumericPoint>().ElementAt(2);

                    NumericValue numericValue1 = numericPoint1.GetFirstChild<NumericValue>();
                    //numericValue1.Text = ".50";
                    if (numericValue1 != null)
                        numericValue1.Text = dataModel.Interaction.ToString();


                    NumericValue numericValue2 = numericPoint2.GetFirstChild<NumericValue>();
                    //numericValue2.Text = ".10";
                    if (numericValue2 != null)
                        numericValue2.Text = dataModel.Knowlegde.ToString();


                    NumericValue numericValue3 = numericPoint3.GetFirstChild<NumericValue>();
                    //numericValue3.Text = ".40";
                    if (numericValue3 != null)
                        numericValue3.Text = dataModel.Image.ToString();


                }
            }

            chartSpace.Save();
        }

        public ActionResult Index()
        {
            TDMGuestServices = loadFileData<TDMGuestServices>(Server.MapPath("~/App_Data/jsonData.txt"));

            PresentationDocument oPDoc = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            PresentationPart oPPart = oPDoc.PresentationPart;
            SlidePart sectionSlidePart = (SlidePart)oPPart.GetPartById("rId3");
            SlidePart sectionSlidePart2 = (SlidePart)oPPart.GetPartById("rId4");
            //SlidePart sectionSlidePart3 = (SlidePart)oPPart.GetPartById("rId8");
            SlidePart sectionSlidePart4 = (SlidePart)oPPart.GetPartById("rId7");

            var tbl = sectionSlidePart4.Slide.Descendants<Table>().First();
            for (int i = 1; i < TDMGuestServices.TblDataSlide6.Length + 1; i++)
            {
                var tr1 = tbl.Descendants<TableRow>().ElementAtOrDefault(i);
                setCellData(tr1, 0, TDMGuestServices.TblDataSlide6[i - 1].Date);

                setCellData(tr1, 1, TDMGuestServices.TblDataSlide6[i - 1].Location);

                setCellData(tr1, 2, TDMGuestServices.TblDataSlide6[i - 1].Time);

                setCellData(tr1, 3, TDMGuestServices.TblDataSlide6[i - 1].Category);

                setCellData(tr1, 4, TDMGuestServices.TblDataSlide6[i - 1].Positive);

                setCellData(tr1, 5, TDMGuestServices.TblDataSlide6[i - 1].Negative);
            }


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

            setChartData(sectionSlidePart2.ChartParts.ElementAt(0), TDMGuestServices.Chart1DataSlide3);
            setChartData(sectionSlidePart2.ChartParts.ElementAt(1), TDMGuestServices.Chart2DataSlide3);

            //foreach (var chartPart1 in sectionSlidePart2.ChartParts)
            //{
            //    try
            //    {
            //        ChangeChartPart(chartPart1);

            //    }
            //    catch (Exception ex)
            //    {

            //    }

            //}


            //SlideLayoutPart slideLayoutPart = (SlideLayoutPart)sectionSlidePart.GetPartById("rId1");
            //DocumentFormat.OpenXml.Drawing.TextBody textBody1 = sectionSlidePart.SlideParts.FirstOrDefault().Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().First();
            //SlidePart slide1 = GetFirstSlide(oPDoc);
            //DocumentFormat.OpenXml.Drawing.Shape textBody1 = slide1.Slide.Descendants<DocumentFormat.OpenXml.Drawing.Shape>().First();
            //DocumentFormat.OpenXml.Drawing.TextBody textBody1 = sectionSlidePart.Slide.Descendants<DocumentFormat.OpenXml.Drawing.TextBody>().First();


            setStringValues(sectionSlidePart);
            setStringValues(sectionSlidePart2);
            setStringValues(sectionSlidePart4);

            sectionSlidePart2.Slide.Save();
            //sectionSlidePart3.Slide.Save();
            sectionSlidePart4.Slide.Save();

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

        private static void setCellData(TableRow tr1, int index, string value)
        {
            var cell = tr1.Descendants<TableCell>().ElementAtOrDefault(index);
            DocumentFormat.OpenXml.Drawing.TextBody tb = cell.Elements<DocumentFormat.OpenXml.Drawing.TextBody>().First();
            Paragraph p = tb.Elements<Paragraph>().ElementAtOrDefault(0);
            Run r = p.Elements<Run>().FirstOrDefault();
            //Run r = new Run();
            //RunProperties runProperties2 = new RunProperties() { Language = "en-US" };
            //DocumentFormat.OpenXml.Drawing.Text t = new DocumentFormat.OpenXml.Drawing.Text();
            //t.Text = value;
            //r.Append(runProperties2);
            //r.Append(t);
            //p.Append(r);
            if (r != null)
            {
                DocumentFormat.OpenXml.Drawing.Text t = r.Elements<DocumentFormat.OpenXml.Drawing.Text>().First();
                t.Text = value;
            }
            else
            {

            }
            //return t;
        }

        private void setStringValues(SlidePart sectionSlidePart)
        {
            foreach (var item in sectionSlidePart.Slide.Descendants())
            {
                if (item.GetType() == typeof(DocumentFormat.OpenXml.Presentation.Shape))
                {
                    foreach (Paragraph paragraph in item.Descendants().OfType<Paragraph>())
                    {
                        setRunData(paragraph);
                    }
                }

                if (item.GetType() == typeof(ShapeTree))
                {
                    foreach (DocumentFormat.OpenXml.Presentation.Shape shape in item.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                    {
                        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                        {
                            setRunData(paragraph);
                        }
                    }
                }
                if (item.GetType() == typeof(CommonSlideData))
                {
                    CommonSlideData sldData1 = (CommonSlideData)item;

                    ShapeTree tree = sldData1.ShapeTree;
                    foreach (DocumentFormat.OpenXml.Presentation.Shape shape in tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                    {
                        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                        {
                            setRunData(paragraph);
                        }
                    }

                    foreach (DocumentFormat.OpenXml.Presentation.GroupShape groupShape in tree.Elements<DocumentFormat.OpenXml.Presentation.GroupShape>())
                    {
                        foreach (DocumentFormat.OpenXml.Presentation.Shape shapeinner in groupShape.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                        {
                            foreach (Paragraph paragraph in shapeinner.Descendants().OfType<Paragraph>())
                            {
                                setRunData(paragraph);
                            }
                        }

                        foreach (Paragraph paragraph in groupShape.Descendants().OfType<Paragraph>())
                        {
                            setRunData(paragraph);
                        }
                    }

                    foreach (DocumentFormat.OpenXml.Presentation.GraphicFrame shape in tree.Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
                    {
                        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                        {
                            setRunData(paragraph);
                        }
                    }

                    foreach (DocumentFormat.OpenXml.Presentation.ConnectionShape shape in tree.Elements<DocumentFormat.OpenXml.Presentation.ConnectionShape>())
                    {
                        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                        {
                            setRunData(paragraph);
                        }
                    }

                    sectionSlidePart.Slide.Save();
                }

            }
        }
        
        private void setRunData(Paragraph paragraph)
        {
            foreach (Run run in paragraph.Elements<Run>())
            {
                if ((GetPropertyName(() => TDMGuestServices.ScoreSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.ScoreSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.VisitsSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.VisitsSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.ScoreYTDlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.ScoreYTDlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.VisitsYTDSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.VisitsYTDSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title1Slide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title1Slide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title2Slide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title2Slide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title1Slide2).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title1Slide2);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title2Slide2).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title2Slide2);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Chart1Title).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Chart1Title);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Chart2Title).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Chart2Title);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Month1MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Month1MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Month2MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Month2MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Month3MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Month3MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Month4MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Month4MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Month4MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Month4MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.MonthValue4MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.MonthValue4MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.ScoreValueMiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.ScoreValueMiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.MonthValue1MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.MonthValue1MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.MonthValue2MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.MonthValue2MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.MonthValue3MiddleSlide3).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.MonthValue3MiddleSlide3);
                }
                else if ((GetPropertyName(() => TDMGuestServices.ProductKnowledgeSlide2).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.ProductKnowledgeSlide2);
                }
                else if ((GetPropertyName(() => TDMGuestServices.InterActionWithGuestsSlide2).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.InterActionWithGuestsSlide2);
                }
                else if ((GetPropertyName(() => TDMGuestServices.ProfessionalImageSlide2).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.ProfessionalImageSlide2);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title1Slide6).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title1Slide6);
                }
                else if ((GetPropertyName(() => TDMGuestServices.Title2Slide6).Contains(run.Text.InnerText)))
                {
                    run.Text = new DocumentFormat.OpenXml.Drawing.Text(TDMGuestServices.Title2Slide6);
                }
            }
        }

        #region utils
        public static string GetPropertyName<T>(Expression<Func<T>> propertyLambda)
        {
            MemberExpression me = propertyLambda.Body as MemberExpression;
            if (me == null)
            {
                throw new ArgumentException("You must pass a lambda of the form: '() => Class.Property' or '() => object.Property'");
            }

            string result = string.Empty;
            do
            {
                result = me.Member.Name + "." + result;
                me = me.Expression as MemberExpression;
            } while (me != null);

            result = result.Remove(result.Length - 1); // remove the trailing "."
            return result;
        }

        private T loadFileData<T>(string path)
        {
            FileStream fstream = new FileStream(path, FileMode.Open);
            StreamReader reader = new StreamReader(fstream);
            string json = reader.ReadToEnd();
            T model = JsonConvert.DeserializeObject<T>(json);
            return model;
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
        #endregion

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