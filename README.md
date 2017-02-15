ExcelDataGenerator
==================

Lightweight and fast tool written in C# to generate Microsoft excel file (.xlsx) based on a sample file data. It does not requires excel or access installation.
It reads the first 1000 rows of user provided sample excel file and keeps repeating the same rows in output excel to generate the file of desired size.
Useful for generating excel files of different sizes.

## Finding the binaries
The compiled binaries are available in the release. Download all files in binaries folder, use
FileGenerator.exe and FileGenerator.exe.config to generate excel files.

## How to use
1. Provide SampleFilePath, OutputFilePath and SizeOfOutputFileInMb in FileGenerator.exe.config file.
2. Make sure sample file has 10 or more rows of data in the first sheet.
3. Run the executable (FileGenerator.exe).
4. The file generation progress and completion message will be displayed in command prompt.

## RAM Considerations:
1. An excel file generally requires RAM of 3-5 times of the file size to open manually.
2. File generator requires RAM of 8-10 times of the desired excel file size to generate output file. For example, to generate an excel file of 1 GB the system should have at least 10 GB of available RAM.
   We are planning to reduce RAM consumption by writing excel in chunks using stream for which we would need to modify EPPlus source code or use a better library.

## Note:
1. It generates data of only first sheet.
2. If the system is not having sufficient RAM to generate output file then system out of memory exception will be raised.
3. It is using customized ExcelDataReader (https://github.com/rishios/ExcelDataGenerator) and EPPlus (https://github.com/pruiz/EPPlus). Thanks to their makers.