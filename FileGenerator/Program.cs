using System;
using System.Collections.Specialized;
using System.Linq;

using System.Data;
using System.IO;
using System.Configuration;

using OfficeOpenXml;
using Excel;

namespace FileGenerator
{
    public class Program
    {
        public static void Main(string[] args)
        {
            NameValueCollection nvc = ConfigurationManager.AppSettings;
            string sampleFilePath = nvc.Get("SampleFilePath");
            string outputFilePath = nvc.Get("OutputFilePath");
            int sizeOfOutputFileInMb = 0;
            Int32.TryParse(nvc.Get("SizeOfOutputFileInMb"), out sizeOfOutputFileInMb);
            bool isInputValid = true;
            if (!File.Exists(Path.GetFullPath(sampleFilePath)))
            {
                Console.WriteLine("Please create sample file and provide proper file path in the App.config.");
                isInputValid = false;
            }
            if (isInputValid && sizeOfOutputFileInMb < 1)
            {
                Console.WriteLine("Please provide proper size of output file to be generated in the App.config.");
                isInputValid = false;
            }

            DirectoryInfo dir = new DirectoryInfo(Path.GetDirectoryName(outputFilePath));
            if (!dir.Exists)
            {
                Console.WriteLine("Please create folder for output file to be generated.");
                isInputValid = false;
            }

            if (isInputValid)
            {
                GenerateFile(sampleFilePath, outputFilePath, sizeOfOutputFileInMb);
            }

            Console.ReadLine();
        }

        private static void GenerateFile(string sampleFilePath, string outputFilePath, int sizeOfOutputFileInMb)
        {
            string fistSheetName = string.Empty;
            int batchSize = 0;
            DataTable datatable = null, batchSizeDataTable = null;
            FileInfo destFileInfo = new FileInfo(outputFilePath);

            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(sampleFilePath)))
            {
                if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
                // Read one sheet of excel in batch
                excelReader.SheetName = fistSheetName = excelReader.GetSheetNames().FirstOrDefault();
                excelReader.IsFirstRowAsColumnNames = true;
                if (excelReader.ReadBatch())
                {
                    datatable = excelReader.GetCurrentBatch();
                }
            }
            batchSize = 10000;
            batchSizeDataTable = datatable.Clone();
            if (datatable.Rows.Count < batchSize)
            {
                while (batchSizeDataTable.Rows.Count < batchSize)
                {
                    datatable.AsEnumerable().CopyToDataTable(batchSizeDataTable, LoadOption.Upsert);
                }
                batchSizeDataTable = batchSizeDataTable.AsEnumerable().Take(batchSize).CopyToDataTable();
            }

            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                var workSheet = excelPackage.Workbook.Worksheets.Add(fistSheetName);
                workSheet.Cells[1, 1].LoadFromDataTable(batchSizeDataTable, true);
                excelPackage.SaveAs(destFileInfo);
            }
            int sizeOfBatchInKb = Convert.ToInt32(destFileInfo.Length / 1024);
            int loopCounter = (sizeOfOutputFileInMb * 1024 / sizeOfBatchInKb) + 1;

            if (batchSize * loopCounter > 1000000)
            {
                Console.WriteLine("The generated file would exceed excel row limit of ~1 million rows. Please reduce the size of excel file to be generated or increase the data in rows of sample excel file.");
                return;
            }
            Int32 updateRow = batchSize + 2;
            Console.WriteLine("Iteration: 0");
            Console.WriteLine($"Rows: {updateRow - 2}");
            if (loopCounter > 1)
            {
                using (ExcelPackage excelPackage = new ExcelPackage(destFileInfo))
                {
                    for (int i = 1; i < loopCounter; i++)
                    {
                        Console.WriteLine($"Iteration: {i}");
                        updateRow = i * batchSize + 2;
                        Console.WriteLine($"Rows: {batchSize + updateRow - 2}");
                        // Adding cells data
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[fistSheetName];
                        worksheet.Cells[updateRow, 1].LoadFromDataTable(batchSizeDataTable, false);
                    }

                    DateTime dt = DateTime.Now;
                    Console.WriteLine("File save on disk started...");
                    excelPackage.SaveAs(destFileInfo);
                    Console.WriteLine($"File save on disk completed in {(DateTime.Now - dt).TotalMinutes} minutes.");
                }
                updateRow += batchSize;
            }

            Console.WriteLine("");
            Console.WriteLine("******************");
            Console.WriteLine($"File of size ~{sizeOfOutputFileInMb} Mb having {updateRow - 2} data rows created successfully.");
            Console.WriteLine("******************");
        }
    }
}
