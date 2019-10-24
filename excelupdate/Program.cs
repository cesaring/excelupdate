using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace excelupdate
{
    class Program
    {

        String SourceFile;
        String SourceSheet;
        String TargetFile;
        String TargetSheet;
        String OriginFile;
        String OriginTimeStamp;
        String OriginTimeStampNew;

        static void Main(string[] args)
        {//open up source Excel
            Program myProgram;
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBookDest;
            Excel.Worksheet xlWorkSheet;
            Excel.Worksheet xlWorkSheetDest;
            object misValue = System.Reflection.Missing.Value;

            myProgram = new Program();

            myProgram.InitParameters();

            if (myProgram.CopyOriginToSource(myProgram.OriginFile, myProgram.SourceFile) == 1)
            {


                if (myProgram.IsFileLocked(new FileInfo(myProgram.SourceFile)) == false && myProgram.IsFileLocked(new FileInfo(myProgram.TargetFile)) == false)
                {

                    xlApp = new Excel.Application();
                    xlWorkBook = xlApp.Workbooks.Open(myProgram.SourceFile);
                    xlWorkBookDest = xlApp.Workbooks.Open(myProgram.TargetFile);
                    xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets[myProgram.SourceSheet];
                    xlWorkSheetDest = (Excel.Worksheet)xlWorkBookDest.Worksheets[myProgram.TargetSheet];

                    xlWorkSheet.Copy(Before: (Excel.Worksheet)xlWorkBookDest.Worksheets.get_Item(1));
                    xlWorkSheet = (Excel.Worksheet)xlWorkBookDest.Worksheets[myProgram.SourceSheet];

                    //clear all the cell values in the target
                    xlWorkSheetDest.Cells.ClearContents();

                    //get the last row/col in the sourcesheet
                    Excel.Range last = xlWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, misValue);

                    //replacing values
                    for (int x = 1; x <= last.Row; x++)
                    {
                        for (int y = 1; y <= last.Column; y++)
                        {
                            xlWorkSheetDest.Cells[x, y].Value = xlWorkSheet.Cells[x, y].Value;
                        }
                    }

                    //removing the imported sheet
                    xlApp.DisplayAlerts = false;
                    xlWorkSheet.Delete();
                    xlApp.DisplayAlerts = true;


                    xlWorkBookDest.Save();
                    Console.WriteLine("Target Sheet has been replaced:" + myProgram.TargetSheet);

                    myProgram.runMacroOnSheets(xlApp, xlWorkBookDest);
                    Console.WriteLine("Sorts Applied.");
                    xlWorkBookDest.Save();

                    xlWorkBook.Close(true, misValue, misValue);
                    xlWorkBookDest.Close(true, misValue, misValue);
                    xlApp.Quit();

                    myProgram.LogToFile();

                    myProgram.releaseObject(xlWorkSheet);
                    myProgram.releaseObject(xlWorkSheetDest);
                    myProgram.releaseObject(xlWorkBookDest);
                    myProgram.releaseObject(xlWorkBook);
                    myProgram.releaseObject(xlApp);


                } //if IsFileLocked==false
                else
                {
                    Console.WriteLine("SourceFile or TargetFile is not available for writing...");
                }

            }//if FileCopy success==1
           // Console.ReadKey();
            
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                Console.WriteLine("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        void InitParameters()
        {
            this.OriginFile = "";
            this.SourceFile = "";
            this.SourceSheet = "";
            this.TargetFile = "";
            this.TargetSheet = "";
            this.OriginTimeStamp = "";
            this.OriginTimeStampNew = "";

            string line;
            string[] val = new string[2];

            //open file 
            System.IO.StreamReader file = new System.IO.StreamReader(@"excelupdate.cfg");

            //for each line
            while ((line = file.ReadLine()) != null)
            {
                //split the line by =
                val = line.Split('=');

                //if val[0]="VARIABLENAME" set sourcefile
                if (val[0]=="SOURCEFILE") { this.SourceFile = val[1]; };
                if (val[0] == "SOURCESHEET") { this.SourceSheet = val[1]; };
                if (val[0] == "TARGETFILE") { this.TargetFile = val[1]; };
                if (val[0] == "TARGETSHEET") { this.TargetSheet = val[1]; };
                if (val[0] == "ORIGINFILE") { this.OriginFile = val[1]; }
                if (val[0] == "ORIGINTIMESTAMP") { this.OriginTimeStamp = val[1];  this.OriginTimeStampNew = val[1]; };
            }

            //close file
            file.Close();
        }

        void runMacroOnSheets(Excel.Application myApp, Excel.Workbook myBook)
        {
            object misValue = System.Reflection.Missing.Value;


            foreach (Excel.Worksheet mySheet in myBook.Worksheets) {
                if (mySheet.Name != "Raw Data" && mySheet.Name != "Manipulated" && mySheet.Name.StartsWith("Sheet")!=true)
                {
                    mySheet.Select(misValue);
                    myApp.Run("ApplySort");
                }
            }

        }
        int CopyOriginToSource(String Ofile, String Sfile)
        {
            int Success;
            bool tryAgain;
            int attempts;

            Success = 0;
            attempts = 0;
            tryAgain = true;

            FileInfo fileToCopy = new FileInfo(Ofile);
            FileInfo fileToReplace = new FileInfo(Sfile);

            while (tryAgain && attempts<=4) {
                if (IsFileLocked(fileToCopy) || IsFileLocked(fileToReplace))
                {
                    attempts++;
                    tryAgain = true;
                    Console.WriteLine("File is open");
                    Thread.Sleep(5000);
                    if(attempts>4) { tryAgain = false; };
                }
                else { tryAgain = false; }
            }

            if (attempts <= 4)
            {
                String OfileTimeStamp;
                OfileTimeStamp = File.GetLastWriteTime(Ofile).ToString();

                if (OfileTimeStamp != this.OriginTimeStamp) //this file has changed
                {
                    try
                    {

                        File.Copy(Ofile, Sfile, true);
                        this.OriginTimeStampNew = OfileTimeStamp;
                        Success = 1;
                    }

                    catch (Exception e)
                    {
                        Console.WriteLine($"There was an error copying from {Ofile} to {Sfile}.");
                    }
                }
                else { Console.WriteLine("File has not changed since last update."); }

            } else
            {
                Success = 0;
            }

            return Success;
        }

        protected virtual bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException ex)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }

            //file is not locked
            return false;
        }


        void LogToFile()
        {
            System.IO.StreamWriter strWriter = new StreamWriter(@"excelupdate.cfg");
            
            strWriter.WriteLine($"ORIGINFILE={this.OriginFile}" );
            strWriter.WriteLine($"SOURCEFILE={this.SourceFile}");
            strWriter.WriteLine($"SOURCESHEET={this.SourceSheet}");
            strWriter.WriteLine($"TARGETFILE={this.TargetFile}");
            strWriter.WriteLine($"TARGETSHEET={this.TargetSheet}");
            strWriter.WriteLine($"ORIGINTIMESTAMP={this.OriginTimeStampNew}");

            strWriter.Close();

        }
    }
}
