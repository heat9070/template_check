using CheckListTemplates.Models.DTO;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using DocumentFormat.OpenXml.Presentation;
using System.Text.RegularExpressions;

namespace CheckListTemplates.Controllers
{
    public class PPTController : Controller
    {
        static PresentationDocument presentationDocument;
        PresentationPart presentationPart;

        //public PPTController()
        //{
        //    presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
        //    presentationPart = presentationDocument.PresentationPart;
        //}

        public ActionResult Index()
        {
            //presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            //presentationPart = presentationDocument.PresentationPart;
            if (presentationDocument == null)
                presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var lst = new List<Models.DTO.SlideDTO>();
            foreach (DocumentFormat.OpenXml.Presentation.SlideId item in presentationPart.Presentation.SlideIdList)
            {
                lst.Add(new Models.DTO.SlideDTO { Id = item.RelationshipId, Name = item.RelationshipId });
            }
            //var slides = presentationPart.Presentation.SlideIdList.Select(o => new Slide { Id = presentationPart.GetPartById(o.RelationshipId}).ToList();
            return View(lst);
        }

        public ActionResult TextBoxes(string Id)
        {
            if (presentationDocument == null)
                presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var slide = (SlidePart)presentationPart.GetPartById(Id);
            var lst = new List<RunDTO>();
            setStringValues(slide, ref lst);

            return View(lst);
        }

        public ActionResult Charts(string Id)
        {
            if (presentationDocument == null)
                presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var slide = (SlidePart)presentationPart.GetPartById(Id);

            var lst = new List<ChartDTO>();
            foreach (var chart in slide.ChartParts)
            {
                foreach (var item in chart.ChartSpace.Descendants())
                {
                    if (item.GetType() == typeof(DocumentFormat.OpenXml.Drawing.Charts.Title))
                    {
                        lst.Add(new ChartDTO { Name = ((DocumentFormat.OpenXml.Drawing.Charts.Title)item).ChartText.InnerText,
                        Id = ((DocumentFormat.OpenXml.Drawing.Charts.Title)item).ChartText.InnerText
                        });
                    }

                }
            }

            //var lst = new List<RunTB>();

            return View(lst);
        }

        public ActionResult Tables(string Id)
        {
            if (presentationDocument == null)
                presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var slide = (SlidePart)presentationPart.GetPartById(Id);

            var tables = slide.Slide.Descendants<Table>().ToList();
            var lst = new List<TableDTO>();
            foreach (var table in tables)
            {
                lst.Add(new TableDTO { Id = table.InnerText, Name = table.InnerText });
            }

            return View(lst);
        }

        private void setStringValues(SlidePart sectionSlidePart, ref List<RunDTO> lst)
        {
            foreach (var item in sectionSlidePart.Slide.Descendants())
            {
                if (item.GetType() == typeof(DocumentFormat.OpenXml.Presentation.Shape))
                {
                    foreach (Paragraph paragraph in item.Descendants().OfType<Paragraph>())
                    {
                        setRunData(paragraph, ref lst);
                    }
                }

                //if (item.GetType() == typeof(ShapeTree))
                //{
                //    foreach (DocumentFormat.OpenXml.Presentation.Shape shape in item.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                //    {
                //        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                //        {
                //            setRunData(paragraph, ref lst);
                //        }
                //    }
                //}
                //if (item.GetType() == typeof(CommonSlideData))
                //{
                //    CommonSlideData sldData1 = (CommonSlideData)item;

                //    ShapeTree tree = sldData1.ShapeTree;
                //    foreach (DocumentFormat.OpenXml.Presentation.Shape shape in tree.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                //    {
                //        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                //        {
                //            setRunData(paragraph, ref lst);
                //        }
                //    }

                //    foreach (DocumentFormat.OpenXml.Presentation.GroupShape groupShape in tree.Elements<DocumentFormat.OpenXml.Presentation.GroupShape>())
                //    {
                //        foreach (DocumentFormat.OpenXml.Presentation.Shape shapeinner in groupShape.Elements<DocumentFormat.OpenXml.Presentation.Shape>())
                //        {
                //            foreach (Paragraph paragraph in shapeinner.Descendants().OfType<Paragraph>())
                //            {
                //                setRunData(paragraph, ref lst);
                //            }
                //        }

                //        foreach (Paragraph paragraph in groupShape.Descendants().OfType<Paragraph>())
                //        {
                //            setRunData(paragraph, ref lst);
                //        }
                //    }

                //    foreach (DocumentFormat.OpenXml.Presentation.GraphicFrame shape in tree.Elements<DocumentFormat.OpenXml.Presentation.GraphicFrame>())
                //    {
                //        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                //        {
                //            setRunData(paragraph, ref lst);
                //        }
                //    }

                //    foreach (DocumentFormat.OpenXml.Presentation.ConnectionShape shape in tree.Elements<DocumentFormat.OpenXml.Presentation.ConnectionShape>())
                //    {
                //        foreach (Paragraph paragraph in shape.Descendants().OfType<Paragraph>())
                //        {
                //            setRunData(paragraph, ref lst);
                //        }
                //    }
                //}

            }
        }

        private void setRunData(Paragraph paragraph, ref List<RunDTO> lst)
        {
            foreach (Run run in paragraph.Elements<Run>())
            {
                try
                {
                    var pattern = @"\[(.*?)\]";
                    var matches = Regex.Matches(run.Text.Text, pattern);
                    foreach (var match in matches)
                    {
                        lst.Add(new Models.DTO.RunDTO { Text = run.Text.Text });
                    }

                }
                catch (Exception ex)
                {

                }

            }
        }


    }
}