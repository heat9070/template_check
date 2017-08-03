using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace CheckListTemplates.Models.DTO
{
    public class TDMGuestServices
    {
        public int ProductKnowledgeSlide2 { get; set; }
        public int InterActionWithGuestsSlide2 { get; set; }
        public int ProfessionalImageSlide2 { get; set; }
        public int VisitsSlide3 { get; set; }
        public int VisitsYTDSlide3 { get; set; }
        public int ScoreSlide3 { get; set; }
        public int ScoreYTDlide3 { get; set; }
        public ChartDataHolder[] Chart2DataSlide2 { get; set; }
        public ChartDataHolder[] Chart1DataSlide2 { get; set; }
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
}