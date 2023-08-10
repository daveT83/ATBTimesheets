using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace BabyBoiProgramWithExtrraBoi.Infrastructure
{
    internal class ConvertToExcel
    {
        public string InputFile { get; set; }
        public string OutputFile { get; set; }
        public ReadData ReadData { get; set; }

        public ConvertToExcel(string inputFile, string outputFile)
        {
            InputFile = inputFile;
            OutputFile = outputFile;
            ReadData = new ReadData(inputFile);
        }

        /// <summary>
        /// Reads the file
        /// </summary>
        public void LoadFile()
        {
            ReadData.ReadFile();
            ReadData.ProcessData();
        }

        /// <summary>
        /// Writes to Excel
        /// </summary>
        public void ExportToExcel()
        {
            using (XLWorkbook workbook = new XLWorkbook())
            {
                var worksheet = workbook.AddWorksheet("Sheet1");
                bool isHeaderEncodingSelector = true;
                //worksheet.Range(1, 5, 1, 25).Row(2).Merge();
                worksheet.AddPicture(Path.Combine(GetResourcesFolder(new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory)).FullName, "Logo.png"));//.MoveTo(workbook.Cell("A1"));
                bool isRecords = false;
                bool isEndOfReport = false;

                for (int i = 0; i < ReadData.Data.Count; i++)
                {
                    int recordSpacing = 0;
                    int endOfReportSpacing = 6;
                    for (int k = 0; k < ReadData.Data[i].Count; k++)
                    {
                        if (!isRecords && ReadData.Data[i][0].Equals("Date"))
                        {
                            isRecords = true;
                        }
                        else if (isRecords && String.IsNullOrEmpty(ReadData.Data[i][0]))
                        {
                            isRecords = false;
                        }
                        else if (!isEndOfReport && ReadData.Data[i].Count >= 17 && ReadData.Data[i][16].Equals("End of Report"))
                        {
                            isEndOfReport = true;
                        }

                        if (isRecords)      //records
                        {
                            if (!String.IsNullOrEmpty(ReadData.Data[i][k]))
                            {
                                if (k == 0 || k == 1)
                                {
                                    recordSpacing = 1;
                                    worksheet.Range(1, 1, 1, 2).Row(i + 1).Merge();
                                }
                                else if (k >= 2 && k <= 5)
                                {
                                    recordSpacing = 3;
                                    worksheet.Range(1, 3, 1, 6).Row(i + 1).Merge();
                                }
                                else if (k >= 6 && k <= 11)
                                {
                                    recordSpacing = 7;
                                    worksheet.Range(1, 7, 1, 12).Row(i + 1).Merge();
                                }
                                else if (k >= 12 && k <= 16)
                                {
                                    recordSpacing = 13;
                                    worksheet.Range(1, 13, 1, 17).Row(i + 1).Merge();
                                }
                                else if (k >= 17 && k <= 21)
                                {
                                    recordSpacing = 18;

                                    worksheet.Range(1, 18, 1, 22).Row(i + 1).Merge();
                                }
                                else if (k >= 22 && k <= 25)
                                {
                                    recordSpacing = 23;
                                    worksheet.Range(1, 23, 1, 26).Row(i + 1).Merge();
                                }
                                else if (k == 26 || k == 26)
                                {
                                    recordSpacing = 27;
                                    worksheet.Range(1, 27, 1, 28).Row(i + 1).Merge();
                                }

                                worksheet.Cell(i + 1, recordSpacing).Value = ReadData.Data[i][k];
                            }
                        }
                        else if (isEndOfReport)     //end of report
                        {
                            if (!String.IsNullOrEmpty(ReadData.Data[i][k]))
                            {
                                if (k == 8 || k == 9)
                                {
                                    worksheet.Range(1, 7, 1, 10).Row(i + 1).Merge();

                                }
                                else if (k == 13 || k == 14)
                                {
                                    endOfReportSpacing = 11;
                                    worksheet.Range(1, 11, 1, 16).Row(i + 1).Merge();

                                }
                                else if (k == 16 || k == 17)
                                {
                                    endOfReportSpacing = 17;
                                    worksheet.Range(1, 17, 1, 23).Row(i + 1).Merge();
                                }
                                else
                                {
                                    endOfReportSpacing = 24;
                                    worksheet.Range(1, 24, 1, 27).Row(i + 1).Merge();
                                }
                                worksheet.Cell(i + 1, endOfReportSpacing).Value = ReadData.Data[i][k];
                            }
                        }
                        else
                        {
                            worksheet.Cell(i + 1, k + 1).Value = ReadData.Data[i][k];
                        }
                    }
                }

                worksheet.Columns().AdjustToContents();
                worksheet.Rows().AdjustToContents();
                workbook.SaveAs(OutputFile);
            }
        }

        /// <summary>
        /// Finds the resources folder
        /// </summary>
        /// <param name="currentDirectory"></param>
        /// <param name="isFirst"></param>
        /// <returns></returns>
        private DirectoryInfo GetResourcesFolder(DirectoryInfo currentDirectory, bool isFirst = true)
        {
            DirectoryInfo resourceFolder = null;

            if (!isFirst)
            {
                resourceFolder = currentDirectory.EnumerateDirectories("Resources", SearchOption.TopDirectoryOnly).FirstOrDefault();
            }
            else
            {
                resourceFolder = currentDirectory.EnumerateDirectories("Resources", SearchOption.AllDirectories).FirstOrDefault();
            }

            if (resourceFolder == null)
            {
                resourceFolder = GetResourcesFolder(currentDirectory.Parent, false);
            }


            return resourceFolder;
        }
    }
}
