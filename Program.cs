using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

internal class Program
{
    private static void Main(string[] args)
    {
        var listOfStudents = new List<string>()
        {
            "Камила Анарбаева",
            "Севда Мавлюдова",
            "Давид Гимадиев",
            "Джамшид ктототам",
            "текст текс",
            "текст текст",
            "раз два",
        };

        var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "ListOfUsers.xlsx");

        try
        {
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("My Sheet");

                sheet.Cells["A1"].Value = "Nomer";
                sheet.Cells["B1"].Value = "Ism-familiya";

                for (int i = 0; i < listOfStudents.Count; i++)
                {
                    sheet.Cells[$"A{i + 2}"].Value = i + 1;
                    sheet.Cells[$"B{i + 2}"].Value = listOfStudents[i].ToString();
                }

                package.SaveAs(new FileInfo(path));
            }

            Console.WriteLine($"Файл {path} успешно создан.");
        }
        catch (LicenseException)
        {
            // Handle the LicenseException by setting the LicenseContext
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // Retry the code after setting the LicenseContext
            using (var package = new ExcelPackage())
            {
                var sheet = package.Workbook.Worksheets.Add("My Sheet");

                sheet.Cells["A1"].Value = "Nomer";
                sheet.Cells["B1"].Value = "Ism-familiya";

                for (int i = 0; i < listOfStudents.Count; i++)
                {
                    sheet.Cells[$"A{i + 2}"].Value = i + 1;
                    sheet.Cells[$"B{i + 2}"].Value = listOfStudents[i].ToString();
                }

                package.SaveAs(new FileInfo(path));
            }

            Console.WriteLine($"Файл {path} успешно создан после настройки LicenseContext.");
        }
    }
}