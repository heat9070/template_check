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
        PresentationDocument presentationDocument;
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
            var lst = new List<Models.DTO.Slide>();
            foreach (DocumentFormat.OpenXml.Presentation.SlideId item in presentationPart.Presentation.SlideIdList)
            {
                lst.Add(new Models.DTO.Slide { Id = item.RelationshipId, Name = item.RelationshipId });
            }
            //var slides = presentationPart.Presentation.SlideIdList.Select(o => new Slide { Id = presentationPart.GetPartById(o.RelationshipId}).ToList();
            return View(lst);
        }

        public ActionResult Edit(string Id)
        {
            presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var slide = (SlidePart)presentationPart.GetPartById(Id);
            var lst = new List<RunTB>();
            setStringValues(slide, ref lst);

            return View(lst);
        }

        public ActionResult Edit2(string Id)
        {
            presentationDocument = PresentationDocument.Open(Server.MapPath("~/App_Data/TDMGuestServices.pptx"), true);
            presentationPart = presentationDocument.PresentationPart;
            var slide = (SlidePart)presentationPart.GetPartById(Id);

            var lst = new List<ChartHolder>();
            foreach (var chart in slide.ChartParts)
            {
                lst.Add(new ChartHolder { Name = "" });
            }

            //var lst = new List<RunTB>();

            return View(lst);
        }

        private void setStringValues(SlidePart sectionSlidePart, ref List<RunTB> lst)
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

        private void setRunData(Paragraph paragraph, ref List<RunTB> lst)
        {
            foreach (Run run in paragraph.Elements<Run>())
            {
                try
                {
                    var pattern = @"\[(.*?)\]";
                    var matches = Regex.Matches(run.Text.Text, pattern);
                    foreach (var match in matches)
                    {
                        lst.Add(new Models.DTO.RunTB { Text = run.Text.Text });
                    }

                }
                catch (Exception ex)
                {

                }

            }
        }


    }
}