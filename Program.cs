using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace GiadaExcelDataParser
{
    class Program
    {
        static String magicWordToFind = "Region";
        static void Main(string[] args)
        {
            List<List<String>> data = new List<List<String>>();
            Console.WriteLine("Loading Data...");
            List<String> firstRow = new List<String>();
            firstRow.Add("Type");
            firstRow.Add("NHS_AREA_Code");
            firstRow.Add("ONS_CODE");
            firstRow.Add("Name");
            firstRow.Add("Open_GP_Practices");
            firstRow.Add("Included_GP_Practices");
            firstRow.Add("Appointments");
            firstRow.Add("Face_to_Face");
            firstRow.Add("Home_visit");
            firstRow.Add("Telephone");
            firstRow.Add("Video/Online");
            firstRow.Add("Unkown");
            firstRow.Add("Month");
            firstRow.Add("Year");

            Excel.Application xlApp = new Excel.Application();
            foreach (String path in Directory.GetFiles(Directory.GetCurrentDirectory() + "\\FileInput"))
            {
                Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path, ReadOnly: true);
                //Getting Date
                String dateTime = GetDateTime(xlWorkbook.Worksheets[1]);
                //Getting table 3c
                String tableName;
                Excel.Worksheet xlWorksheet = xlWorkbook.Worksheets[10];
                tableName = xlWorksheet.Name;
                if (tableName!="Table 3c")
                {
                    xlWorksheet = xlWorkbook.Worksheets[11];
                    if(tableName!= "Table 3c")
                    {
                        xlWorksheet = xlWorkbook.Worksheets[12];
                    }
                }
                //Parsing Data and feeding the dataset
                Console.WriteLine("\tParsing Data:" + dateTime);
                GetRowData(xlWorksheet, data,dateTime);
            }
            Console.WriteLine("\t\tWriting File");
            WriteData(firstRow, data);
            Console.WriteLine("\t\t\tFinish");
        }
        private static void GetRowData(Excel.Worksheet worksheet,List<List<String>> data,String dateTime)
        {
            bool notEmpty = true;
            bool firstRow = true;
            int row = 13;
            do
            {
                String t = ((Excel.Range)worksheet.Cells[row,1]).Value+"";
                List<String> tmpList = new List<string>();
                //CurrentRow-Serve a evitare l'errore riga 13 inesistente in alcuni file
                if (firstRow && t!="")
                {
                    if (t == magicWordToFind)
                    {
                        for (int col = 1; col < 14; col++)
                        {
                            var cell= ((Excel.Range)worksheet.Cells[row, col]).Value;
                            if(cell!=null)
                                tmpList.Add(Convert.ToString(cell));
                        }
                        tmpList.Add(dateTime.Split(',')[0]);
                        tmpList.Add(dateTime.Split(',')[1]);
                        data.Add(tmpList);
                    }
                }
                else
                {
                    if (t == magicWordToFind)
                    {
                        for (int r = 1; r < 14; r++)
                        {
                            String cell = ((Excel.Range)worksheet.Cells[row, r]).Value+"";
                            if (cell != null && cell != "")
                                tmpList.Add((String)cell);
                        }
                        tmpList.Add(dateTime.Split(',')[0]);
                        tmpList.Add(dateTime.Split(',')[1]);
                        data.Add(tmpList);
                    }
                }
                row++;
                String nextT= ((Excel.Range)worksheet.Cells[row, 1]).Value+"";
                if (nextT == null || nextT == "")//NextRow-Serve a controllare la fine del dataset
                    notEmpty = false;
            } while (notEmpty);
        }

        private static void WriteData(List<String> firstRow, List<List<String>> data)
        {
            Excel.Application application = new Excel.Application();
            if (application != null)
            {
                Excel.Workbook excelWorkbook = application.Workbooks.Add();
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelWorkbook.Sheets.Add();
                //Qui Sto Aggiungendo gli header
                for (int i = 0; i < firstRow.Count; i++)
                {
                    excelWorksheet.Cells[1, i+1] = firstRow[i];
                }
                //Qui Aggiungo il dataset che è stato precedentemente creato
                for (int i = 0; i < data.Count; i++)
                {
                    List<String> tmp = data[i];//Prendo la riga
                    for (int j = 0; j < tmp.Count; j++)
                    {
                        //Prendo la colonna
                        excelWorksheet.Cells[i+2, j + 1] = tmp[j];
                    }
                }
                String mainPath = Directory.GetCurrentDirectory() + "\\FileOutput";
                Directory.CreateDirectory(mainPath);
                if(File.Exists(mainPath + "\\GP_Appt_stataData.xlsx")) File.Delete(mainPath + "\\GP_Appt_stataData.xlsx");
                application.ActiveWorkbook.SaveAs(mainPath+"\\GP_Appt_stataData.xlsx");

                excelWorkbook.Close();
                application.Quit();

                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorksheet);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelWorkbook);
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(application);
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        private static String GetDateTime(Excel.Worksheet tmpWs)
        {
            //GettingDateTime
            String dateTime = ((Excel.Range)tmpWs.Cells[9, 1]).Value + "";
            String dateTIme = dateTime.Replace(" ", "");
            if (dateTime.Split(",").Length<=1)
            {
                dateTime = ((Excel.Range)tmpWs.Cells[10, 1]).Value + "";
                if (dateTime.Replace(" ", "").Split(",").Length <= 1)
                {
                    dateTime = ((Excel.Range)tmpWs.Cells[11, 1]).Value + "";
                }
            }
            String month = dateTime.Replace(" ", "").Split(",")[0];
            String year = dateTime.Replace(" ", "").Split(",")[1];
            int newMonth = 0;
            switch (month)
            {
                case "January":
                    newMonth = 1;
                    break;
                case "February":
                    newMonth = 2;
                    break;
                case "March":
                    newMonth = 3;
                    break;
                case "April":
                    newMonth = 4;
                    break;
                case "May":
                    newMonth = 5;
                    break;
                case "June":
                    newMonth = 6;
                    break;
                case "July":
                    newMonth = 7;
                    break;
                case "August":
                    newMonth = 8;
                    break;
                case "September":
                    newMonth = 9;
                    break;
                case "October":
                    newMonth = 10;
                    break;
                case "November":
                    newMonth = 11;
                    break;
                case "December":
                    newMonth = 12;
                    break;
                default:
                    break;
            }
            return newMonth+","+year;
        }
    }
}
