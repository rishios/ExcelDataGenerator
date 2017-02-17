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
            int sizeOfOutputFileInMb;
            bool isFirstRowAsColumnNames;
            Int32.TryParse(nvc.Get("SizeOfOutputFileInMb"), out sizeOfOutputFileInMb);
            bool.TryParse(nvc.Get("IsFirstRowAsColumnNames"), out isFirstRowAsColumnNames);
            bool isInputValid = true;
            if (!File.Exists(Path.GetFullPath(sampleFilePath)))
            {
                Console.WriteLine("Please create sample file and provide proper file path in the FileGenerator.exe.config.");
                isInputValid = false;
            }
            if (isInputValid && sizeOfOutputFileInMb < 1)
            {
                Console.WriteLine("Please provide proper size of output file to be generated in the FileGenerator.exe.config.");
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
                GenerateFile(isFirstRowAsColumnNames, sampleFilePath, outputFilePath, sizeOfOutputFileInMb);
            }

            Console.ReadLine();
        }

        private static void GenerateFile(bool isFirstRowAsColumnNames, string sampleFilePath, string outputFilePath, int sizeOfOutputFileInMb)
        {
            string firstSheetName = string.Empty;
            int batchSize, desiredRows, sizeOfBatchInKb;
            DataTable datatable = null, batchSizeDataTable = null;
            FileInfo destFileInfo = new FileInfo(outputFilePath);
            Console.WriteLine("Process started...");
            using (IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(File.OpenRead(sampleFilePath)))
            {
                if (!excelReader.IsValid) { throw new Exception(excelReader.ExceptionMessage); }
                // Read one sheet of excel in batch
                excelReader.SheetName = firstSheetName = excelReader.GetSheetNames().FirstOrDefault();
                excelReader.IsFirstRowAsColumnNames = isFirstRowAsColumnNames;
                if (excelReader.ReadBatch())
                {
                    datatable = excelReader.GetCurrentBatch();
                }
            }

            batchSize = 1000;
            batchSizeDataTable = CreateBatchSizeTable(batchSize, datatable);
            SaveTableToExcel(batchSizeDataTable, firstSheetName, isFirstRowAsColumnNames, true, ref destFileInfo);
            sizeOfBatchInKb = Convert.ToInt32(destFileInfo.Length / 1024);
            desiredRows = Convert.ToInt32(batchSize * (((sizeOfOutputFileInMb * 1024) / (decimal)sizeOfBatchInKb)));
            // making optimal batch size
            if (sizeOfBatchInKb < 1024 && desiredRows > batchSize)
            {
                int optimalBatchSizeInMb = sizeOfOutputFileInMb >= 5 ? 5 : 1; // // 5 mb or 1 mb batch
                batchSize = Convert.ToInt32(batchSize * ((optimalBatchSizeInMb * 1024) / (decimal)sizeOfBatchInKb));
                batchSizeDataTable = CreateBatchSizeTable(batchSize, batchSizeDataTable);
                SaveTableToExcel(batchSizeDataTable, firstSheetName, isFirstRowAsColumnNames, false, ref destFileInfo);
                sizeOfBatchInKb = Convert.ToInt32(destFileInfo.Length / 1024);
                desiredRows = Convert.ToInt32(batchSize * (((sizeOfOutputFileInMb * 1024) / (decimal)sizeOfBatchInKb)));
            }
            if (desiredRows > 1000000)
            {
                Console.WriteLine("The generated file would exceed excel row limit of ~1 million rows. Please try one or more of the following: a) Increase number of columns in sample excel file. b) Increase data in rows of sample excel file. c) Reduce the size of output file to be generated. ");
                return;
            }

            int loopCounter = (sizeOfOutputFileInMb * 1024 / sizeOfBatchInKb) + 1;
            int lastBatchRows = (desiredRows % batchSize);
            bool isLastBatchResizeRequired = batchSize - lastBatchRows > 0;
            int updateRow, updateRowHeaderAdjuster = isFirstRowAsColumnNames ? 2 : 1;
            long totalRows = batchSize;
            if (loopCounter > 1)
            {
                Console.WriteLine("Iteration: 0");
                Console.WriteLine($"Rows: {totalRows}");
                using (ExcelPackage excelPackage = new ExcelPackage(destFileInfo))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[firstSheetName];
                    for (int i = 1; i < loopCounter; i++)
                    {
                        if (i < loopCounter - 1) { totalRows += batchSize; }
                        else if (isLastBatchResizeRequired) // Adjusting rows of last batch
                        {
                            totalRows += lastBatchRows;
                            batchSizeDataTable = batchSizeDataTable.AsEnumerable().Take(lastBatchRows).CopyToDataTable();
                        }
                        Console.WriteLine($"Iteration: {i}");
                        Console.WriteLine($"Rows: {totalRows}");

                        // Adding cells data
                        updateRow = i * batchSize + updateRowHeaderAdjuster;
                        //ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[firstSheetName];
                        worksheet.Cells[updateRow, 1].LoadFromDataTable(batchSizeDataTable, false);
                    }

                    DateTime dt = DateTime.Now;
                    Console.WriteLine("File save on disk started...");
                    excelPackage.SaveAs(destFileInfo);
                    Console.WriteLine($"File save on disk completed in {(DateTime.Now - dt).TotalMinutes} minutes.");
                }
            }
            else
            {
                if (isLastBatchResizeRequired)
                {
                    totalRows = totalRows - batchSize + lastBatchRows;
                    batchSizeDataTable = batchSizeDataTable.AsEnumerable().Take(lastBatchRows).CopyToDataTable();
                    SaveTableToExcel(batchSizeDataTable, firstSheetName, isFirstRowAsColumnNames, true, ref destFileInfo);
                }
                Console.WriteLine("Iteration: 0");
                Console.WriteLine($"Rows: {totalRows}");
            }

            Console.WriteLine("");
            Console.WriteLine("******************");
            Console.WriteLine($"File of size ~{sizeOfOutputFileInMb} Mb having {totalRows} data rows created successfully.");
            Console.WriteLine("******************");
        }

        private static DataTable CreateBatchSizeTable(int batchSize, DataTable datatable)
        {
            DataTable batchSizeDataTable = null;
            if (datatable.Rows.Count < batchSize)
            {
                batchSizeDataTable = datatable.Clone();
                while (batchSizeDataTable.Rows.Count < batchSize)
                {
                    datatable.AsEnumerable().CopyToDataTable(batchSizeDataTable, LoadOption.Upsert);
                }
            }
            else
            {
                batchSizeDataTable = datatable;
            }
            return GetBatchSizeTable(batchSize, batchSizeDataTable);
        }

        private static DataTable GetBatchSizeTable(int batchSize, DataTable datatable)
        {
            return datatable.AsEnumerable().Take(batchSize).CopyToDataTable();
        }

        private static void SaveTableToExcel(DataTable datatable, string firstSheetName, bool isFirstRowAsColumnNames, bool createFile, ref FileInfo fileInfo)
        {
            ExcelWorksheet workSheet = null;
            using (ExcelPackage excelPackage = createFile ? new ExcelPackage() : new ExcelPackage(fileInfo))
            {
                workSheet = createFile ? excelPackage.Workbook.Worksheets.Add(firstSheetName) : excelPackage.Workbook.Worksheets[firstSheetName];
                workSheet.Cells[1, 1].LoadFromDataTable(datatable, isFirstRowAsColumnNames);
                excelPackage.SaveAs(fileInfo);
                excelPackage.File.Refresh();
            }
        }

    }
}
