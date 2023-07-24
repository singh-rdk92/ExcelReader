using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using ExcelDataReader;
using System.Text;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml;
using System.Drawing;

public class IndexModel : PageModel
{
    public class UploadResponse
    {
        public int SuccessCount { get; set; }
        public int FailureCount { get; set; }
        public byte[] ExcelFile { get; set; }
    }

    public IActionResult OnPostUpload(List<IFormFile> files)
    {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance); // Register custom encoding provider

        List<List<string>> failureResponseList = new List<List<string>>();
        List<List<string>> dataRows = new List<List<string>>();

        foreach (var file in files)
        {
            if (file.Length > 0)
            {
                // Read the Excel file using ExcelDataReader
                using (var stream = new MemoryStream())
                {
                    file.CopyTo(stream);
                    stream.Position = 0; // Reset the stream position to 0

                    using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration()
                    {
                        FallbackEncoding = Encoding.GetEncoding(1252) // Specify the encoding explicitly
                    }))
                    {
                        // Assuming the first two rows contain headers and you want to skip them
                        int startRow = 0; // Start processing from the first row
                        bool hasBlankValue = false;

                        do
                        {
                            // Skip the first two rows (headers)
                            for (int i = 0; i < 2; i++)
                            {
                                reader.Read();
                            }

                            // Process data rows
                            while (reader.Read()) // Read a row
                            {
                                var rowData = new List<string>();
                                hasBlankValue = false;
                                var remarks = new List<string>();

                                for (int i = 0; i < reader.FieldCount; i++) // Skip the first column
                                {
                                    var cellValue = reader.GetValue(i)?.ToString() ?? "";
                                    rowData.Add(cellValue);

                                    if (i == 1 || i == 2 || i == 3 || i == 4 || i == 5 || i == 6 || i == 7 || i == 9 || i == 10 || i == 11) // Check required columns (indexes 1, 2, 3, and 4)
                                    {
                                        if (string.IsNullOrEmpty(cellValue))
                                        {

                                            hasBlankValue = true;
                                        }
                                    }
                                }

                                if (hasBlankValue)
                                {
                                    failureResponseList.Add(rowData);
                                }
                                else
                                {
                                    dataRows.Add(rowData);
                                }
                            }
                        } while (reader.NextResult()); // Move to the next sheet if any
                    }
                }
            }
        }

        // Now, you have the dataRows containing data from rows (excluding the first column and the first two rows)
        // and the failureResponseList containing rows with any blank cell values

        // Count the number of successful and failed records
        int successCount = dataRows.Count;
        int failureCount = failureResponseList.Count;

        var response = new UploadResponse
        {
            SuccessCount = successCount,
            FailureCount = failureCount
        };

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        using (var excelPackage = new ExcelPackage())
        {
            var worksheet = excelPackage.Workbook.Worksheets.Add("Failure Records");

            // Add the fixed row with various texts
           
            worksheet.Cells["A1"].Value = "Sr.No";
            worksheet.Cells["B1"].Value = "Company";
            worksheet.Cells["C1"].Value = "city";
            worksheet.Cells["D1"].Value = "country";
            worksheet.Cells["E1"].Value = "isLocationIsUTn";
            worksheet.Cells["F1"].Value = "isProjectInUT";
            worksheet.Cells["G1"].Value = "CollectiveAction";
            worksheet.Cells["H1"].Value = "PartnerName";
            worksheet.Cells["I1"].Value = "ContactName";
            worksheet.Cells["J1"].Value = "ProjectInfo";
            worksheet.Cells["K1"].Value = "investment";
            worksheet.Cells["L1"].Value = "focusOfProject";
            worksheet.Cells["M1"].Value = "privateProject";
            worksheet.Cells["N1"].Value = "domainName";
            worksheet.Cells["O1"].Value = "Comments";
            worksheet.Cells["P1"].Value = "Remarks"; // Add the Remarks header
            

            worksheet.Cells["A2"].Value = "Sr.No";
            worksheet.Cells["B2"].Value = "Company Name";
            worksheet.Cells["C2"].Value = "Chandigarh";
            worksheet.Cells["D2"].Value = "India";
            worksheet.Cells["E2"].Value = "Please select";
            worksheet.Cells["F2"].Value = "Please select";
            worksheet.Cells["G2"].Value = "Please select";
            worksheet.Cells["H2"].Value = "If Action is Y";
            worksheet.Cells["I2"].Value = "Optional";
            worksheet.Cells["J2"].Value = "please enter Name or Title of Project";
            worksheet.Cells["K2"].Value = "Please select";
            worksheet.Cells["L2"].Value = "can have multiple comma seprated values";
            worksheet.Cells["M2"].Value = "Optional";
            worksheet.Cells["N2"].Value = "Optional";
            worksheet.Cells["O2"].Value = "Optional";


            //Adding Back Ground Color as blue and Font Color As White.
            using (var range = worksheet.Cells["A1:P1"])
            {
                var headerFill = range.Style.Fill;
                headerFill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                Color colFromHex = ColorTranslator.FromHtml("#395697");
                headerFill.BackgroundColor.SetColor(colFromHex);

                var headerFont = range.Style.Font;
                headerFont.Color.SetColor(Color.White);
            }

            // Add the failure records to the worksheet
            int row = 3; // Start from the Third row to add failure records
            foreach (var failureRow in failureResponseList)
            {
                int column = 1;
                foreach (var cellValue in failureRow)
                {
                    worksheet.Cells[row, column].Value = cellValue;
                    column++;
                }

                // Add the Remarks data to the last column
                var remarks = new List<string>();
                for (int i = 1; i < failureRow.Count; i++) // Skip the first column (Sr.No)
                {
                    if (i == 1 || i == 2 || i == 3 || i == 4 || i == 5 || i == 6 || i == 7 || i == 9 || i == 10 || i == 11) // Check required columns (indexes 1, 2, 3, and 4)
                    {
                        if (string.IsNullOrEmpty(failureRow[i]))
                        {
                            // Add the header name to the remarks list when there's a blank value
                            var headerName = worksheet.Cells[1, i + 1].Value?.ToString() ?? $"Header{i + 1}";

                            remarks.Add(headerName);
                        }
                    }
                }

                // Combine all the header names with commas and add " (Columns are mandatory)" at the end
                string remarksText = string.Join(", ", remarks);
                if (!string.IsNullOrEmpty(remarksText))
                {
                    remarksText += " (Columns are mandatory)";
                }

                // Set the Remarks data in the last column for the current row
                worksheet.Cells[row, failureRow.Count + 1].Value = remarksText;
                row++;
            }

            worksheet.Cells[worksheet.Dimension.Address].AutoFitColumns();

            // Save the Excel file to a MemoryStream
            using (var memoryStream = new MemoryStream())
            {
                excelPackage.SaveAs(memoryStream);
                response.ExcelFile = memoryStream.ToArray();
            }
        }

        // Return the JSON and Excel file in a custom response object
        return new JsonResult(response);
    }
}
