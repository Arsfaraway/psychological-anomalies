using System.Data;
using Npgsql;
using Reader1.Models.Configuration;
using System;

namespace Reader1.Database
{
    public class DatabaseContext
    {
        private readonly string _connectionString;

        public DatabaseContext(string connectionString)
        {
            _connectionString = connectionString;
        }

        public bool IsTableEmpty()
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(_connectionString))
            {
                connection.Open();

                using (NpgsqlCommand command = new NpgsqlCommand())
                {
                    command.Connection = connection;
                    command.CommandText = "SELECT COUNT(*) FROM configurations";

                    int count = Convert.ToInt32(command.ExecuteScalar());

                    return count == 0;
                }
            }
        }
        public void DeleteAllConfigurations()
        {
            using (NpgsqlConnection connection = new NpgsqlConnection(_connectionString))
            {
                connection.Open();

                using (NpgsqlCommand command = new NpgsqlCommand())
                {
                    command.Connection = connection;
                    command.CommandText = "DELETE FROM configurations";
                    command.ExecuteNonQuery();
                }
            }
        }

        public void SaveConfiguration(Configuration config)
        {
            if (!IsTableEmpty())
            {
                DeleteAllConfigurations();
            }

            using (NpgsqlConnection connection = new NpgsqlConnection(_connectionString))
            {
                connection.Open();

                using (NpgsqlCommand command = new NpgsqlCommand())
                {
                    command.Connection = connection;
                    command.CommandText = @"
                        INSERT INTO configurations (
                            organisation_name, 
                            class_number, 
                            class_letter,
                            classroom_teacher_name,
                            classroom_teacher_email,
                            classroom_teacher_phone,
                            psychologist_name,
                            psychologist_email,
                            psychologist_phone,
                            academic_year_start_reporting,
                            reporting_start_quarter_number,
                            reporting_academic_year,
                            reporting_quarter_number
                        ) VALUES (
                            @OrganisationName,
                            @ClassNumber,
                            @ClassLetter,
                            @ClassroomTeacherName,
                            @ClassroomTeacherEmail,
                            @ClassroomTeacherPhone,
                            @PsychologistName,
                            @PsychologistEmail,
                            @PsychologistPhone,
                            @AcademicYearStartReporting,
                            @ReportingStartQuarterNumber,
                            @ReportingAcademicYear,
                            @ReportingQuarterNumber
                        )";

                    command.Parameters.AddWithValue("@OrganisationName", config.OrganisationName);
                    command.Parameters.AddWithValue("@ClassNumber", config.ClassNumber);
                    command.Parameters.AddWithValue("@ClassLetter", config.ClassLetter);
                    command.Parameters.AddWithValue("@ClassroomTeacherName", config.ClassroomTeacherName);
                    command.Parameters.AddWithValue("@ClassroomTeacherEmail", config.ClassroomTeacherEmail);
                    command.Parameters.AddWithValue("@ClassroomTeacherPhone", config.ClassroomTeacherPhone);
                    command.Parameters.AddWithValue("@PsychologistName", config.PsychologistName);
                    command.Parameters.AddWithValue("@PsychologistEmail", config.PsychologistEmail);
                    command.Parameters.AddWithValue("@PsychologistPhone", config.PsychologistPhone);
                    command.Parameters.AddWithValue("@AcademicYearStartReporting", config.AcademicYearStartReporting);
                    command.Parameters.AddWithValue("@ReportingStartQuarterNumber", config.ReportingStartQuarterNumber);
                    command.Parameters.AddWithValue("@ReportingAcademicYear", config.ReportingAcademicYear);
                    command.Parameters.AddWithValue("@ReportingQuarterNumber", config.ReportingQuarterNumber);

                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
