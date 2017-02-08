File Generator:
The tool generates excel file (.xlsx) of desired size based on the sample excel file provided by user.
It reads the first 1000 rows of sample excel and keeps repeating the same rows in output excel to generate the file of
desired size.

Usage:
1. Provide SampleFilePath, OutputFilePath and SizeOfOutputFileInMb in FileGenerator.exe.config file.
2. Make sure sample file has 10 or more rows of data.
3. Run the executable.
4. The file generation progress and completion message will be displayed in command prompt.

RAM Considerations:
1. An excel file requires RAM of 4-5 times of the file size to open manually.
2. File generator requires RAM of 8-10 times of the desired excel file size to generate output file.

Note:
1. It assumes that the first row of first sheet of sample excel is a header row.
2. The size of output file will be slightly higher than the desired size.
3. It generates data of only first sheet.