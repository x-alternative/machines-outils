using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using System.Text;
using System.Collections.Generic;
using Newtonsoft.Json;

namespace finances
{
    using FinancesStats = Dictionary<string, Dictionary<Company, Dictionary<int, string>>>;
    using RawFinancesStats = Dictionary<string, Dictionary<String, Dictionary<int, string>>>;

    class Company
    {
        public int ID { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
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
            return this.ID;
        }

        public override string ToString()
        {
            return String.Format($"{ID} - {Name}");
        }
    }

    class Program
    {
        static readonly int[] allYears = new int[6] { 2021, 2020, 2019, 2018, 2017, 2016 };

        static readonly HashSet<String> metricsOfInterest = new HashSet<string> {
            "Résultat net",
            "Chiffre d'affaires",
            "Valeur ajoutée",
            "Marge brute",
            "Excédent brut d'exploitation (EBITDA)",
            "Résultat d'exploitation (EBIT)",
            "Délai de paiement clients, en jours",
            "Délai de paiement fournisseurs, en jours",
            "Trésorerie",
            "Dettes financières",
            "Dette financière nette (DFN)",
        };

        static readonly string[] exclude = new string[] { "000 - note", "0001 - données macro" };

        static Dictionary<int, Company> companies = new Dictionary<int, Company>();

        const string statsJsonPath = "../Data/stats.json";
        const string reportPath = "../Data/report.xlsx";

        static void Main(string[] args)
        {
            var baseDirectory = args.Length > 1 ? args[1] : "/home/rportalez/Documents/machines/data";
            ReadMainFile(new FileInfo(Path.Combine(baseDirectory, "000 - etablissements.xlsx")));
            FinancesStats values;
            if (false) //!File.Exists(statsJsonPath))
            {
                values = ReadAndParseRawData(baseDirectory, out int companiesCount, out int financialDataCount);
                File.WriteAllText(statsJsonPath, JsonConvert.SerializeObject(values));
            }
            else
            {
                values = ReadLocalStats();
            }

            CreateExcelReport(values);
            ComputeTaxesInReport(values);
            ComputeCapitalCost(values);
        }

        private static FinancesStats ReadLocalStats()
        {
            FinancesStats values;
            // custom and ugly deserialization
            var rawvalues = JsonConvert.DeserializeObject<RawFinancesStats>(File.ReadAllText(statsJsonPath));
            values = rawvalues.ToDictionary(kvp => kvp.Key,
            kvp =>
            {
                return kvp.Value.ToDictionary(kk =>
                {
                    var key = kk.Key;
                    var id = int.Parse(key.Substring(0, key.IndexOf('-') - 1));
                    return companies[id];
                }, kk => kk.Value);
            }).OrderByDescending(kvp => kvp.Value.Count()).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            return values;
        }

        private static FinancesStats ReadAndParseRawData(string baseDirectory, out int companiesCount, out int financialDataCount)
        {
            var values = new FinancesStats();
            companiesCount = companies.Count();
            financialDataCount = 0;
            foreach (var dir in Directory.GetDirectories(baseDirectory))
            {
                DirectoryInfo dirInfo = new DirectoryInfo(dir);
                var companyId = int.Parse(dirInfo.Name.Substring(0, 3));

                var company = companies.FirstOrDefault(c => c.Value.ID == companyId);
                if (!exclude.Contains(dirInfo.Name) && company.Value != null)
                {
                    var tables = Directory.GetFiles(dir, "*.xlsx");
                    if (!tables.Any())
                    {
                        continue;
                    }

                    financialDataCount += 1;

                    foreach (var file in tables)
                    {
                        var fileInfo = new FileInfo(file);
                        var name = fileInfo.Name;
                        if (int.TryParse(name.Substring(name.Length - 9, 4), out int year))
                        {
                            ReadFinanceFile(fileInfo, year, company.Value, values);
                        }
                        else
                        {
                            Console.WriteLine($"invalid xlsx file name : {file}");
                        }
                    }
                }
            }

            return values.OrderByDescending(kvp => kvp.Value.Count()).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
        }

        private static void ComputeCapitalCost(FinancesStats values)
        {

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(reportPath)))
            {
                var workbook = xlPackage.Workbook;
                var cvaeSheet = workbook.Worksheets.Add("coût de la dette - résultat net");
                var capitalSheet = workbook.Worksheets["coût du capital"];
                cvaeSheet.Cells[1, 1].Value = "ID";
                cvaeSheet.Cells[1, 2].Value = "Effectifs";
                int columnNum = 3;
                foreach (var year in allYears)
                {
                    cvaeSheet.Cells[1, columnNum++].Value = year;
                }

                int rowCount = capitalSheet.Dimension.Rows;
                for (int rowNum = 2; rowNum <= rowCount; ++rowNum)
                {
                    columnNum = 1;
                    var companyId = int.Parse(capitalSheet.Cells[rowNum, 1].GetValue<String>());
                    var employeeCount = companies[companyId].EmployeeCount;
                    cvaeSheet.Cells[rowNum, columnNum++].Value = companyId;
                    cvaeSheet.Cells[rowNum, columnNum++].Value = employeeCount;

                    for (int yearId = 0; yearId < allYears.Length; ++yearId)
                    {
                        var capitalCost = capitalSheet.Cells[rowNum, yearId + 2].GetValue<double>();
                        int year = allYears[yearId];
                        if (values["résultat net"].Any(c => c.Key.ID == companyId))
                        {
                            var history = values["résultat net"].First(c => c.Key.ID == companyId).Value;
                            double netResult = 0;
                            if (history.TryGetValue(year, out string netResultStr))
                            {
                                double.TryParse(netResultStr, out netResult);
                            }

                            if (netResult != 0 && capitalCost != 0)
                            {
                                cvaeSheet.Cells[rowNum, columnNum].Value = Math.Abs(capitalCost / netResult);
                            }

                            columnNum++;
                        }
                    }
                }
                xlPackage.Save();
            }
        }

        private static void ComputeTaxesInReport(FinancesStats values)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(reportPath)))
            {
                var workbook = xlPackage.Workbook;
                var cvaeSheet = workbook.Worksheets.Add("cvae");
                var cvaeRatioSheet = workbook.Worksheets.Add("ratio cvae résultat net");
                var caSheet = workbook.Worksheets["chiffre d'affaires"];
                var vaSheet = workbook.Worksheets["valeur ajoutée"];
                cvaeSheet.Cells[1, 1].Value = "ID";
                cvaeSheet.Cells[1, 2].Value = "Effectifs";
                cvaeRatioSheet.Cells[1, 1].Value = "ID";
                cvaeRatioSheet.Cells[1, 2].Value = "Effectifs";
                int columnNum = 3;
                foreach (var year in allYears)
                {
                    cvaeSheet.Cells[1, columnNum].Value = year;
                    cvaeRatioSheet.Cells[1, columnNum++].Value = year;
                }

                int rowCount = caSheet.Dimension.Rows;
                for (int rowNum = 2; rowNum <= rowCount; ++rowNum)
                {
                    columnNum = 1;
                    var companyId = int.Parse(vaSheet.Cells[rowNum, 1].GetValue<String>());
                    var effectifs = companies[companyId].EmployeeCount;
                    cvaeSheet.Cells[rowNum, columnNum].Value = companyId;
                    cvaeRatioSheet.Cells[rowNum, columnNum++].Value = companyId;
                    cvaeSheet.Cells[rowNum, columnNum].Value = effectifs;
                    cvaeRatioSheet.Cells[rowNum, columnNum++].Value = effectifs;
                    for (int yearId = 0; yearId < allYears.Length; ++yearId)
                    {
                        var ca = caSheet.Cells[rowNum, yearId + 3].GetValue<double>();
                        var va = vaSheet.Cells[rowNum, yearId + 3].GetValue<double>();
                        int year = allYears[yearId];
                        var history = values["résultat net"].First(c => c.Key.ID == companyId).Value;
                        double netResult = 0;
                        if (history.TryGetValue(year, out string netResultStr))
                        {
                            double.TryParse(netResultStr, out netResult);
                        }

                        var cvae = Cvae.ComputeTax(ca, va);
                        // skip empty cells
                        if (ca == 0 || va == 0)
                        {
                            cvae = 0;
                        }
                        if (cvae != 0)
                        {
                            cvaeSheet.Cells[rowNum, columnNum].Value = cvae;
                        }
                        if (netResult != 0 && cvae != 0)
                        {
                            cvaeRatioSheet.Cells[rowNum, columnNum].Value = Math.Abs(cvae / netResult);
                        }

                        columnNum++;
                    }
                }

                xlPackage.Save();
            }
        }

        private static void CreateExcelReport(FinancesStats values)
        {
            if (File.Exists(reportPath))
            {
                File.Delete(reportPath);
            }
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(reportPath)))
            {
                var workbook = xlPackage.Workbook;
                foreach (var (metricName, metric) in values)
                {
                    var sheet = workbook.Worksheets.Add(metricName);
                    var cells = sheet.Cells;
                    int rowNum = 1;
                    int columnNum = 1;
                    cells[rowNum, columnNum++].Value = "ID";
                    cells[rowNum, columnNum++].Value = "Effectifs";
                    foreach (var year in allYears)
                    {
                        cells[rowNum, columnNum++].Value = year;
                    }
                    rowNum += 1;
                    foreach (var (company, companyMetric) in metric.OrderBy(kk => kk.Key.ID))
                    {
                        columnNum = 1;
                        cells[rowNum, columnNum++].Value = company.ID;
                        cells[rowNum, columnNum++].Value = company.EmployeeCount;
                        foreach (var year in allYears)
                        {
                            string sval = String.Empty;
                            companyMetric.TryGetValue(year, out sval);
                            double dval = 0.0;
                            double.TryParse(sval, out dval);
                            cells[rowNum, columnNum++].Value = dval;
                        }

                        rowNum += 1;
                    }
                }

                xlPackage.Save();
            }
        }

        private static void ReadFinanceFile(FileInfo fileInfo, int year, Company company, FinancesStats values)
        {
            using (ExcelPackage xlPackage = new ExcelPackage(fileInfo))
            {
                var sheet = xlPackage.Workbook.Worksheets.First();
                sheet.Calculate();

                var accountSheet = xlPackage.Workbook.Worksheets["Compte de résultat"];
                if (accountSheet != null)
                {
                    var debtCost = accountSheet.Cells["M41"].GetValue<String>();
                    if (!String.IsNullOrWhiteSpace(debtCost) && debtCost != "n/c" && debtCost != "n/a")
                    {
                        var key_lower = "coût du capital";
                        if (!values.ContainsKey(key_lower))
                        {
                            values.Add(key_lower, new Dictionary<Company, Dictionary<int, string>>());
                        }
                        var key_values = values[key_lower];
                        if (!key_values.ContainsKey(company))
                        {
                            key_values.Add(company, new Dictionary<int, string>());
                        }
                        var company_history = key_values[company];
                        if (!company_history.ContainsKey(year))
                        {
                            company_history.Add(year, debtCost);
                        }
                    }
                }

                var rowCount = sheet.Dimension.End.Row;
                var columnCount = sheet.Dimension.End.Column;
                var sb = new StringBuilder(); //this is your data
                for (int rowNum = 3; rowNum <= rowCount; rowNum++) //select starting row here
                {
                    var row = sheet.Cells[rowNum, 1, rowNum, columnCount];
                    var val = row.GetCellValue<String>(0, 3);
                    var key = row.GetCellValue<String>(0, 0);
                    if (metricsOfInterest.Contains(key) && !String.IsNullOrWhiteSpace(val) && val != "n/c" && val != "n/a" && !String.IsNullOrEmpty(key))
                    {
                        var key_lower = key.ToLowerInvariant();

                        if (!values.ContainsKey(key_lower))
                        {
                            values.Add(key_lower, new Dictionary<Company, Dictionary<int, string>>());
                        }
                        var key_values = values[key_lower];
                        if (!key_values.ContainsKey(company))
                        {
                            key_values.Add(company, new Dictionary<int, string>());
                        }
                        var company_history = key_values[company];
                        if (!company_history.ContainsKey(year))
                        {
                            company_history.Add(year, val);
                        }
                    }
                }
            }
        }

        private static void key_stats(FinancesStats values)
        {
            foreach (var kvp in values.OrderByDescending(v => v.Value.Count()))
            {
                Console.WriteLine($"key {kvp.Key} has {kvp.Value.Count()} companies entries");
            }

        }

        private static void ReadMainFile(FileInfo fileInfo)
        {

            if (!fileInfo.Exists)
            {
                Console.WriteLine("main file does not exist");
                Environment.Exit(6); // abort
            }
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
                    if (String.IsNullOrWhiteSpace(status))
                    {
                        var company = new Company();
                        company.ID = row.GetCellValue<int>(0, 0);
                        company.Name = row.GetCellValue<String>(0, 1);
                        company.EmployeeCount = row.GetCellValue<String>(0, 4);
                        company.Address = String.Format("{0}, {1}",
                            row.GetCellValue<String>(0, 14),
                            row.GetCellValue<String>(0, 12)
                        );
                        company.ShareholderType = row.GetCellValue<String>(0, 24);
                        companies.Add(company.ID, company);
                    }
                }
            }
        }
    }
}
