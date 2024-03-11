using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Reader1.Models.Configuration
{
    public class Configuration
    {
        public Configuration() { }

        //a.Сокращенное название образовательной организации по Уставу;
        public string OrganisationName { get; set; }

        //b.Номер Класса;
        public string ClassNumber { get; set; }

        //c.Буква Класса;
        public string ClassLetter { get; set; }

        //d.Ф.И.О.КР;
        public string ClassroomTeacherName { get; set; }

        //e.E - mail КР;
        public string ClassroomTeacherEmail { get; set;}

        //f.Контактный номер телефона КР;
        public string ClassroomTeacherPhone { get; set;}

        //g.Ф.И.О.педагога - психолога школы;
        public string PsychologistName { get; set; }

        //h.E - mail педагога - психолога школы;
        public string PsychologistEmail { get; set; }

        //i.Контактный номер телефона педагога-психолога;
        public string PsychologistPhone { get; set; }

        //j.Учебный год старта отчетности;
        public string AcademicYearStartReporting { get; set; }

        //k.Номер четверти старта отчетности;
        public string ReportingStartQuarterNumber { get;set; }

        //l.Отчетный учебный год;
        public string ReportingAcademicYear { get;set; }

        //m.Номер отчетной четверти;
        public string ReportingQuarterNumber { get;set; }
        
    }
}
