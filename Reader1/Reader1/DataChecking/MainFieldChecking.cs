using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Org.BouncyCastle.Math.EC.ECCurve;
using Reader1.DataChecking;
using Reader1.Forms;
using Config = Reader1.Models.Configuration.Configuration;

namespace Reader1.DataChecking
{
    public struct MethodStructure
    {
        public Func<string, bool> Method;
        public string[] PropertyNames;
    }

    public class MainFieldChecking
    {
        public static bool IsFilled { get; set; }

        private static string[] configNames = { "Сокращенное название образовательной организации по Уставу", "Номер Класса",
            "Буква Класса", "Ф.И.О. КР", "E-mail КР", "Контактный номер телефона КР", "Ф.И.О. педагога-психолога школы",
            "E-mail педагога-психолога школы", "Контактный номер телефона педагога-психолога", "Учебный год старта отчетности",
            "Номер четверти старта отчетности", "Отчетный учебный год", "Номер отчетной четверти"};
        // todo ресурс файлы

        public static void CheckAllFields(Config config)
        {
            MethodStructure[] methods = new MethodStructure[]
{
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckOrganizationName, PropertyNames = new string[] { "OrganisationName" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckClassNumber, PropertyNames = new string[] { "ClassNumber" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckClassLetter, PropertyNames = new string[] { "ClassLetter" } },
            new MethodStructure { Method =  FieldsCorrectnessChecking.CheckName, PropertyNames = new string[] { "ClassroomTeacherName", "PsychologistName" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckEmail, PropertyNames = new string[] { "ClassroomTeacherEmail", "PsychologistEmail" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckPhone, PropertyNames = new string[] { "ClassroomTeacherPhone", "PsychologistPhone" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.QuarterNumber, PropertyNames = new string[] { "ReportingStartQuarterNumber", "ReportingQuarterNumber" } },
            new MethodStructure { Method = FieldsCorrectnessChecking.CheckYear, PropertyNames = new string[] { "AcademicYearStartReporting", "ReportingAcademicYear" } }
};
            int counter = 0;
            foreach (var property in typeof(Config).GetProperties())
            {
                if (property.CanRead)
                {
                    string propertyName = property.Name;

                    var methodStructure = methods.FirstOrDefault(m => m.PropertyNames.Contains(propertyName));
                    if (methodStructure.Method != null)
                    {

                        string propertyValue = (string)property.GetValue(config);

                        bool isCorrect = false;
                        while (!isCorrect)
                        {
                            isCorrect = methodStructure.Method(propertyValue);
                            if (!isCorrect)
                            {
                                CheckingInformation checkingInformation = new CheckingInformation(configNames[counter], propertyValue);

                                checkingInformation.ShowDialog();

                                propertyName = checkingInformation.field;
                                propertyValue = propertyName;
                            }
                        }
                    }
                }
                counter++;
            }
            IsFilled = true;
        }
    }
}
