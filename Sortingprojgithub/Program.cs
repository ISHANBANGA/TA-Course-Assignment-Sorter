using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;

namespace SortingExcelSheets
{
    class Program
    {
        static void Main(string[] args)
        {
            // Define the file paths (These paths should be updated to the actual file locations)
            string CourseSummaryFilePath = @""; // Path to the file with course data (to be matched with individuals in the other file)
            string TASummaryFilePath = @""; // Path to the file with TA data (contains the names of individuals to be matched)
            string outputFilePath = @""; // Path to save the output file (sorted and merged summary)

            // Read and extract TA assignment details from the course summary file
            var taDetails = ReadCourseSummary(CourseSummaryFilePath);

            // Read the summary file containing TA details
            var summaryData = ReadTASummary(TASummaryFilePath);

            // Merge the TA details into the summary data
            var updatedSummary = MergeTADetails(summaryData, taDetails);

            // Write the updated summary to a new Excel file
            WriteUpdatedSummary(outputFilePath, updatedSummary);
        }

        // Method to read course summary details from the course file
        static List<CourseSummary> ReadCourseSummary(string filePath)
        {
            var taDetails = new List<CourseSummary>();

            // Open the Excel workbook and read the data
            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Assuming the data is in the first sheet

                // Loop through the rows of the worksheet, skipping the header row
                foreach (var row in worksheet.RowsUsed().Skip(2))
                {
                    string course = row.Cell(2).GetValue<string>(); // Course name
                    string type = row.Cell(5).GetValue<string>(); // Type of assignment (e.g., tutorial, lab)

                    // Loop through multiple columns to extract student details
                    for (int i = 0; i < 11; i++)
                    {
                        int bannerIdIndex = 8 + i * 3;
                        int unitsIndex = bannerIdIndex + 2;

                        string studentId = row.Cell(bannerIdIndex).GetValue<string>();
                        if (!string.IsNullOrEmpty(studentId))
                        {
                            var courseSummary = new CourseSummary
                            {
                                StudentID = studentId,
                                Course = course,
                                Type = type,
                                Units = TryGetNullableDecimal(row.Cell(unitsIndex))
                            };

                            taDetails.Add(courseSummary);
                        }
                    }
                }
            }

            return taDetails;
        }

        // Method to safely convert a cell value to a nullable decimal (if present)
        static decimal? TryGetNullableDecimal(IXLCell cell)
        {
            try
            {
                return cell.GetValue<decimal>();
            }
            catch
            {
                return null; // Return null if the conversion fails
            }
        }

        // Method to read TA summary data from the TA summary file
        static List<TASummary> ReadTASummary(string filePath)
        {
            var summaryData = new List<TASummary>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheet(1); // Assuming the data is in the first sheet

                // Loop through the rows of the worksheet, skipping the header row
                foreach (var row in worksheet.RowsUsed().Skip(1))
                {
                    var summary = new TASummary
                    {
                        Pool = row.Cell(1).GetValue<string>(),
                        StudentID = row.Cell(3).GetValue<string>(), // Banner ID is in column 3 (C)
                        FirstName = row.Cell(4).GetValue<string>(),
                        LastName = row.Cell(5).GetValue<string>(),
                        Email = row.Cell(6).GetValue<string>(),
                        Program = row.Cell(11).GetValue<string>(),
                        Rate = row.Cell(15).GetValue<string>()
                    };

                    summaryData.Add(summary);
                }
            }

            return summaryData;
        }

        // Method to merge the TA details into the summary data
        static List<TASummary> MergeTADetails(List<TASummary> summaryData, List<CourseSummary> courseSummary)
        {
            var updatedSummary = new List<TASummary>();

            // Loop through the summary data and match it with the course data
            foreach (var summary in summaryData)
            {
                var matchingDetails = courseSummary.Where(ta => ta.StudentID == summary.StudentID).ToList();

                if (matchingDetails.Count > 0)
                {
                    // Add all matching entries
                    for (int i = 0; i < matchingDetails.Count; i++)
                    {
                        var newSummary = new TASummary
                        {
                            Pool = summary.Pool,
                            StudentID = summary.StudentID,
                            FirstName = summary.FirstName,
                            LastName = summary.LastName,
                            Email = summary.Email,
                            Program = summary.Program,
                            Rate = summary.Rate,
                            Course = matchingDetails[i].Course,
                            Type = matchingDetails[i].Type,
                            Units = matchingDetails[i].Units
                        };

                        updatedSummary.Add(newSummary);
                    }
                }
                else
                {
                    // If no match is found, just add the original summary
                    updatedSummary.Add(summary);
                }
            }

            return updatedSummary;
        }

        // Method to write the updated summary to a new Excel file
        static void WriteUpdatedSummary(string filePath, List<TASummary> updatedSummary)
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Updated Summary");

                // Write header row
                worksheet.Cell(1, 1).Value = "Pool";
                worksheet.Cell(1, 2).Value = "First Name";
                worksheet.Cell(1, 3).Value = "Last Name";
                worksheet.Cell(1, 4).Value = "Student ID";
                worksheet.Cell(1, 5).Value = "Email";
                worksheet.Cell(1, 6).Value = "Program";
                worksheet.Cell(1, 7).Value = "Rate";
                worksheet.Cell(1, 8).Value = "Course";
                worksheet.Cell(1, 9).Value = "Type";
                worksheet.Cell(1, 10).Value = "Units";

                // Apply background color to header row
                var headerRange = worksheet.Range("A1:J1");
                headerRange.Style.Fill.BackgroundColor = XLColor.LightBlue;

                // Write data rows
                for (int i = 0; i < updatedSummary.Count; i++)
                {
                    var summary = updatedSummary[i];
                    worksheet.Cell(i + 2, 1).Value = summary.Pool;
                    worksheet.Cell(i + 2, 2).Value = summary.FirstName;
                    worksheet.Cell(i + 2, 3).Value = summary.LastName;
                    worksheet.Cell(i + 2, 4).Value = summary.StudentID;
                    worksheet.Cell(i + 2, 5).Value = summary.Email;
                    worksheet.Cell(i + 2, 6).Value = summary.Program;
                    worksheet.Cell(i + 2, 7).Value = summary.Rate;
                    worksheet.Cell(i + 2, 8).Value = summary.Course;
                    worksheet.Cell(i + 2, 9).Value = summary.Type;
                    worksheet.Cell(i + 2, 10).Value = summary.Units?.ToString() ?? "";

                    // Apply alternating row colors for better readability
                    var rowRange = worksheet.Range($"A{i + 2}:J{i + 2}");
                    if (i % 2 == 0)
                    {
                        rowRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    }
                }

                // Save the workbook
                workbook.SaveAs(filePath);
            }
        }
    }

    // Class to represent TA details from the course summary
    class CourseSummary
    {
        public string StudentID { get; set; }
        public string Course { get; set; }
        public string Type { get; set; }
        public decimal? Units { get; set; }
    }

    // Class to represent the summary data for each TA
    class TASummary
    {
        public string Pool { get; set; }
        public string StudentID { get; set; }
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string Email { get; set; }
        public string Program { get; set; }
        public string Rate { get; set; }
        public string Course { get; set; }
        public string Type { get; set; }
        public decimal? Units { get; set; }
    }
}
