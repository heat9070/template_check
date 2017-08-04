using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CheckListTemplates.Models.DTO
{
    public class TDMGuestServices
    {
        public string ProductKnowledgeSlide2 { get; set; }
        public string InterActionWithGuestsSlide2 { get; set; }
        public string ProfessionalImageSlide2 { get; set; }
        public string VisitsSlide3 { get; set; }
        public string VisitsYTDSlide3 { get; set; }
        public string ScoreSlide3 { get; set; }
        public string ScoreYTDlide3 { get; set; }
        public string Title1Slide3 { get; set; }
        public string Title2Slide3 { get; set; }
        public string Title1Slide2 { get; set; }
        public string Title2Slide2 { get; set; }
        public string Title1Slide6 { get; set; }
        public string Title2Slide6 { get; set; }
        public string Chart1Title { get; set; }
        public string Chart2Title { get; set; }
        public string Month1MiddleSlide3 { get; set; }
        public string MonthValue1MiddleSlide3 { get; set; }
        public string Month2MiddleSlide3 { get; set; }
        public string MonthValue2MiddleSlide3 { get; set; }
        public string Month3MiddleSlide3 { get; set; }
        public string MonthValue3MiddleSlide3 { get; set; }
        public string Month4MiddleSlide3 { get; set; }
        public string MonthValue4MiddleSlide3 { get; set; }
        public string ScoreValueMiddleSlide3 { get; set; }
        public ChartDataHolder[] Chart2DataSlide3 { get; set; }
        public ChartDataHolder[] Chart1DataSlide3 { get; set; }
        public TblData[] TblDataSlide6 { get; set; }
    }

    public class ChartDataHolder
    {
        public ChartDataSlide chartData { get; set; }
        public string seriesText { get; set; }
    }

    public class ChartDataSlide
    {
        public double Interaction { get; set; }
        public double Knowlegde { get; set; }
        public double Image { get; set; }
    }

    public class TblData
    {
        public string Date { get; set; }
        public string Location { get; set; }
        public string Time { get; set; }
        public string Category { get; set; }
        public string Positive { get; set; }
        public string Negative { get; set; }
    }
}