using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Text;
using System.Collections.Generic;

namespace finances
{

    class Company
    {
        public string ID { get; set; }
        public string SIRET { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Naf { get; set; }
        public string ShareholderType { get; set; }
        public string EmployeeCount { get; set; }

        public override bool Equals(object obj)
        {
            var o = obj as Company;
            if (o == null)
            {
                return false;
            }
            return this.ID == o.ID;
        }

        public override int GetHashCode()
        {
            return int.Parse(this.ID);
        }
    }

    class Program
    {
        static readonly string[] exclude = new string[] { "000 - note", "0001 - données macro" };

        static List<Company> companies = new List<Company>();

        static void Main(string[] args)
        {
            var baseDirectory = args.Length > 1 ? args[1] : "/home/rportalez/Documents/machines/data";
            ReadMainFile(new FileInfo(Path.Combine(baseDirectory, "000 - etablissements.xlsx")));
            foreach (var dir in Directory.GetDirectories(baseDirectory))
            {
                DirectoryInfo dirInfo = new DirectoryInfo(dir);
                var ID = dirInfo.Name.Substring(0, 3);
                var company = companies.Find(x => x.ID == ID);
                if (!exclude.Contains(dirInfo.Name) && company != null)
                {
                    Console.WriteLine("handling compagny {0}", company.Name);
                    var tables = Directory.GetFiles(dir, "*.xlsx");
                    if (!tables.Any())
                    {
                        Console.WriteLine("no table found");
                        continue;
                    }

                    foreach (var file in tables)
                    {
                        var fileInfo = new FileInfo(file);
                        var name = fileInfo.Name;
                        if (int.TryParse(name.Substring(name.Length - 9, 4), out int year))
                        {
                            ReadFinanceFile(fileInfo, year);
                        }
                        else
                        {
                            Console.WriteLine($"invalid xlsx file name : {file}");
                        }
                    }
                }
            }
        }

        private static void ReadFinanceFile(FileInfo fileInfo, int year)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(fileInfo))
            {
                var sheet = xlPackage.Workbook.Worksheets.First();
                sheet.Calculate();
                var rowCount = sheet.Dimension.End.Row;
                var columnCount = sheet.Dimension.End.Column;

                var sb = new StringBuilder(); //this is your data
                for (int rowNum = 1; rowNum <= rowCount; rowNum++) //select starting row here
                {
                    var row = sheet.Cells[rowNum, 1, rowNum, columnCount];
                    foreach (var cell in row.Where(c => c.Value != null))
                    {
                        //Console.WriteLine(cell.Value.ToString());
                    }
                }
            }
        }

        private static void ReadMainFile(FileInfo fileInfo)
        {

            using (ExcelPackage xlPackage = new ExcelPackage(fileInfo))
            {
                var workbook = xlPackage.Workbook;
                workbook.Calculate();
                var sheet = workbook.Worksheets.First();
                var rowCount = sheet.Dimension.End.Row;
                var columnCount = sheet.Dimension.End.Column;
                for (int rowNum = 2; rowNum <= rowCount; rowNum++) //select starting row here
                {
                    var row = sheet.Cells[rowNum, 1, rowNum, columnCount];
                    var status = row.GetCellValue<String>(0, 23);
                    if (String.IsNullOrEmpty(status))
                    {
                        var company = new Company();
                        company.ID = row.GetCellValue<String>(0, 0);
                        company.Name = row.GetCellValue<String>(0, 1);
                        company.SIRET = row.GetCellValue<String>(0, 3);
                        company.EmployeeCount = row.GetCellValue<String>(0, 4);
                        company.Naf = row.GetCellValue<String>(0, 5);
                        company.Address = String.Format("{0} {1} {2} {3} {4}",
                            row.GetCellValue<String>(0, 8),
                            row.GetCellValue<String>(0, 10),
                            row.GetCellValue<String>(0, 11),
                            row.GetCellValue<String>(0, 12),
                            row.GetCellValue<String>(0, 13)
                        );
                        company.ShareholderType = row.GetCellValue<String>(0, 24);
                        companies.Add(company);
                    }
                }
            }
        }
    }
}
