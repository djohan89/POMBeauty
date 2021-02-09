using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelOffice = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using OPM.OPMEnginee;
using ExcelDataReader;

namespace OPM.ExcelHandler
{
    class OpmExcelHandler
    {
        private string _strFileName;


        public string FileName
        {
            set { _strFileName = value; }
            get { return _strFileName; }
        }
        public OpmExcelHandler()
        { }
        ~OpmExcelHandler()
        { }

        public static int fReadExcelFile(string fname, ref Dictionary<int, string> ListTP)
        {
            ExcelOffice.Range xlRange = null;
            ExcelOffice.Workbook xlWorkbook = null;
            ExcelOffice.Application xlApp = null;
            ExcelOffice._Worksheet xlWorksheet = null;
            try
            {
                //Dictionary<string, string> ListTP = new Dictionary<string, string>();

                xlApp = new ExcelOffice.Application();
                xlWorkbook = xlApp.Workbooks.Open(fname);
                xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[2];
                xlRange = xlWorksheet.UsedRange;

                string xName = xlWorksheet.Name.ToString();
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;
                /*Get conntet Danh Sach Tinh/Tp*/
                for (int i = 2; i <= rowCount; i++)
                {
                    string index = Convert.ToString((xlRange.Cells[i, 1] as ExcelOffice.Range).Text);
                    string strName = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                    if (string.Empty != index && string.Empty != strName)
                    {
                        int temp = Int32.Parse(index);
                        ListTP.Add(temp, strName);
                    }
                    else
                    {
                        i = rowCount + 1;
                    }

                }
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 1;
            }
            catch (Exception)
            {
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 0;
            }


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
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        public static int fReadExcelFile3(string fname)
        {

            ExcelOffice.Application xlApp = new ExcelOffice.Application();
            ExcelOffice.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
            int numSheets = xlWorkbook.Sheets.Count;
            //ExcelOffice._Worksheet xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[1];
            List<string> items = new List<string>();
            foreach (ExcelOffice._Worksheet xlWorksheet in xlWorkbook.Sheets)
            {

                string xName = xlWorksheet.Name.ToString();

                items.Add(xName);

            }
            //ExcelOffice.Range xlRange = xlWorksheet.UsedRange;
            //string xName = xlWorksheet.Name.ToString();

            //int rowCount = xlRange.Rows.Count;
            //int colCount = xlRange.Columns.Count;

            //cleanup  
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //rule of thumb for releasing com objects:  
            //  never use two dots, all COM objects must be referenced and released individually  
            //  ex: [somthing].[something].[something] is bad  

            //release com objects to fully kill excel process from running in the background  
            //Marshal.ReleaseComObject(xlRange);
            //Marshal.ReleaseComObject(xlWorksheet);

            //close and release  
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release  
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            return 1;
        }

        public static int fReadPackageListFiles(string[] fnames, ref List<Packagelist> oPackagelist)
        {
            ExcelOffice.Application xlApp = null;
            //List<ExcelOffice.Workbook> xlWorkbookList = null;
            ExcelOffice.Range xlRange = null;
            ExcelOffice.Workbook xlWorkbook = null;
            ExcelOffice._Worksheet xlWorksheet = null;

            ExcelOffice._Worksheet xlWorksheet2 = null;
            ExcelOffice.Range xlRange2 = null;

            List<string> ListSerial = new List<string>();

            try
            {
                int index = 0;
                xlApp = new ExcelOffice.Application();
                //xlWorkbookList = new List<ExcelOffice.Workbook>();
                foreach (string strfilename in fnames)
                {
                    Packagelist temp = new Packagelist();
                    /*Read FileName Infor*/
                    int ret = fReadInforFromFileName(strfilename, ref temp);

                    /*Open Workbook*/
                    xlWorkbook = xlApp.Workbooks.Open(strfilename);

                    /*Read PO Info */
                    xlWorksheet2 = (ExcelOffice._Worksheet)xlWorkbook.Sheets[2];
                    xlRange2 = xlWorksheet2.UsedRange;
                    string xName2 = xlWorksheet2.Name.ToString();
                    string strtype = Convert.ToString((xlRange2.Cells[3, 4] as ExcelOffice.Range).Text);
                    string strPOtemp = strtype.Substring(20, 2);
                    temp.PO_number = strPOtemp;
                    /*Read Hang 2 %*/
                    if (strtype.Contains("("))
                    {
                        temp.Type = "Hang_Bao_Hanh";
                    }

                    /*Open Sheet Serrial*/
                    xlWorksheet = (ExcelOffice._Worksheet)xlWorkbook.Sheets[3];
                    xlRange = xlWorksheet.UsedRange;
                    string xName = xlWorksheet.Name.ToString();
                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    ret = fCheckExistProvince(oPackagelist, temp, ref index);

                    if (1 == ret)
                    {
                        /*Get content Danh Sach Tinh/Tp*/
                        for (int j = 2; j <= rowCount; j++)
                        {
                            //string strCase = Convert.ToString((xlRange.Cells[i, 1] as ExcelOffice.Range).Text);
                            //string strStorage = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            string strSerial = "'"+Convert.ToString((xlRange.Cells[j, 3] as ExcelOffice.Range).Text);
                            //string strMAC = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            //string strSerial_gpon = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            if (string.Empty != strSerial)
                            {
                                oPackagelist[index].SetSerial(strSerial);
                            }
                            else
                            {
                                j = rowCount + 1;
                            }

                        }
                        /*Read PO Info %*/
                        oPackagelist[index].PO_number = temp.PO_number;
                    }
                    else
                    {
                        /*Get content Danh Sach Tinh/Tp*/
                        for (int j = 2; j <= rowCount; j++)
                        {
                            //string strCase = Convert.ToString((xlRange.Cells[i, 1] as ExcelOffice.Range).Text);
                            //string strStorage = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            string strSerial = Convert.ToString((xlRange.Cells[j, 3] as ExcelOffice.Range).Text);
                            //string strMAC = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            //string strSerial_gpon = Convert.ToString((xlRange.Cells[i, 3] as ExcelOffice.Range).Text);
                            if (string.Empty != strSerial)
                            {
                                temp.SetSerial(strSerial);
                            }
                            else
                            {
                                j = rowCount + 1;
                            }

                        }
                        oPackagelist.Add(temp);
                    }

                }

                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlRange2);
                Marshal.ReleaseComObject(xlWorksheet2);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 1;
            }
            catch (Exception)
            {
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();
                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                Marshal.ReleaseComObject(xlRange2);
                Marshal.ReleaseComObject(xlWorksheet2);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);

                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 0;
            }



        }

        public static int fCheckExistProvince(List<Packagelist> oPackagelist, Packagelist temp, ref int index)
        {
            int icount = oPackagelist.Count;
            if (0 == icount)
            {
                return 0;
            }
            else
            {
                for (int j = 0; j < icount; j++)
                {
                    if ((oPackagelist[j].Province == temp.Province) && (temp.Type != "Hang_Bao_Hanh"))
                    {
                        index = j;
                        return 1;
                    }
                }
            }
            return 0;

        }

        public static int fReadInforFromFileName(string strFilename, ref Packagelist oPackagelist)
        {

            try
            {
                string[] strInfo = strFilename.Split('-');
                string[] strDP = strInfo[6].Split(' ');
                string[] strProvince = strInfo[8].Split(' ');
                oPackagelist.DP = strDP[1];
                oPackagelist.Province = strProvince[0] + " " + strProvince[1];
                oPackagelist.Year = strInfo[1] + "-" + strInfo[2];
                return 1;
            }
            catch (Exception)
            {
                return 0;
            }
        }

        public static int fWriteExcelfile(Dictionary<int, string> ListTP, List<Packagelist> oPackagelists, string strFolder)
        {


            int number = ListTP.Count();
            for (int j = 1; j <= number; j++)
            {

                string strprovince = ListTP[j];
                int ret = fWriteProvince(strprovince, oPackagelists, strFolder);
                if (ret == 0)
                {
                    return 0;
                }

            }

            return 0;
        }

        public static int fWriteProvince(string strprovince, List<Packagelist> oPackagelists, string strFolder)
        {
            string strNosignProvince = RemoveSign4VietnameseString(strprovince).TrimEnd();
            ExcelOffice.Application xlApp = null;
            try
            {
                
                /*Check Province in the oPackageList*/
                foreach (Packagelist oPackage in oPackagelists)
                {
                    string strProvincePL = oPackage.Province.TrimEnd();
                    if (strProvincePL == strNosignProvince)
                    {
                        /*Write Excel File*/
                        xlApp = new ExcelOffice.Application();
                        xlApp.StandardFont = "Times New Roman";
                        xlApp.StandardFontSize = 9;
                        ExcelOffice.Workbook xlWorkBook;

                        if (xlApp == null)
                        {
                            MessageBox.Show("Excel is not properly installed!!");
                            return 0;
                        }
                        object misValue = System.Reflection.Missing.Value;
                        xlWorkBook = xlApp.Workbooks.Add(misValue);
                        int countSerial = oPackage.GetListSerial().Count;
                        List<string> temp = oPackage.GetListSerial();
                        string strName = oPackage.Year + "_PO" + oPackage.PO_number + "_" + oPackage.Province + "_" + oPackage.Type;
                        string strFullSNname = strFolder + "\\" + "PHULUC_S_N_" + oPackage.Year + "_PO" + oPackage.PO_number + "_" + oPackage.Province + "_" + oPackage.Type;
                        ExcelOffice._Worksheet xlWorkSheet = (ExcelOffice.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                        for (int k = 0; k < countSerial; k++)
                        {

                            xlWorkSheet.Cells[1, 1] = "STT";
                            xlWorkSheet.get_Range("a1", "a1").EntireRow.Font.Bold = true;
                            xlWorkSheet.get_Range("a1", "a1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            xlWorkSheet.get_Range("a1", "a1").BorderAround();

                            xlWorkSheet.Cells[1, 2] = "Serial";
                            xlWorkSheet.get_Range("b1", "d1").Merge(false);
                            xlWorkSheet.get_Range("b1", "d1").WrapText = true;
                            xlWorkSheet.get_Range("b1", "d1").EntireRow.Font.Bold = true;
                            xlWorkSheet.get_Range("b1", "d1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            xlWorkSheet.get_Range("b1", "d1").BorderAround();

                            xlWorkSheet.Cells[k + 2, 1] = k + 1;
                            string strindex = (k + 2).ToString();
                            xlWorkSheet.get_Range("a" + strindex, "a" + strindex).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            xlWorkSheet.get_Range("a" + strindex, "a" + strindex).BorderAround();



                            string strindexSerial = (k + 2).ToString();
                            xlWorkSheet.Cells[k + 2, 2] = temp[k];
                            xlWorkSheet.get_Range("b" + strindexSerial, "d" + strindexSerial).Merge(false);
                            xlWorkSheet.get_Range("b" + strindexSerial, "d" + strindexSerial).WrapText = true;
                            xlWorkSheet.get_Range("b" + strindexSerial, "d" + strindexSerial).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                            xlWorkSheet.get_Range("b" + strindexSerial, "d" + strindexSerial).BorderAround();

                        }
                        xlWorkSheet.PageSetup.CenterFooter = strName;
                        xlWorkSheet.PageSetup.CenterHeader = strName;
                        /*Calculate Number of Sheet*/
                        //int numberofsheet1 = countSerial / 172;
                        //int numberofsheet2 = countSerial % 172;
                        //if(numberofsheet2 !=0)
                        //{
                        //    numberofsheet1 = numberofsheet1 + 1;
                        //} 
                        //for (int k =0; k< numberofsheet1-1; k++)
                        //{
                        //    ExcelOffice._Worksheet xlWorkSheet = (ExcelOffice.Worksheet)xlWorkBook.Worksheets.get_Item(k+1);
                        //    xlWorkSheet.Cells[1, 1] = "STT";
                        //    xlWorkSheet.get_Range("a1", "a1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("a1", "a1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("a1", "a1").BorderAround();


                        //    xlWorkSheet.Cells[1, 2] = "Serial";
                        //    xlWorkSheet.get_Range("b1", "d1").Merge(false);
                        //    xlWorkSheet.get_Range("b1", "d1").WrapText = true;
                        //    xlWorkSheet.get_Range("b1", "d1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("b1", "d1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("b1", "d1").BorderAround();


                        //    xlWorkSheet.Cells[1, 5] = "STT";
                        //    xlWorkSheet.get_Range("e1", "e1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("e1", "e1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("e1", "e1").BorderAround();


                        //    xlWorkSheet.Cells[1, 6] = "Serial";
                        //    xlWorkSheet.get_Range("f1", "h1").Merge(false);
                        //    xlWorkSheet.get_Range("f1", "h1").WrapText = true;
                        //    xlWorkSheet.get_Range("f1", "h1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("f1", "h1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("f1", "h1").BorderAround();


                        //    xlWorkSheet.Cells[1, 9] = "STT";
                        //    xlWorkSheet.get_Range("i1", "i1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("i1", "i1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("i1", "i1").BorderAround();


                        //    xlWorkSheet.Cells[1, 10] = "Serial";
                        //    xlWorkSheet.get_Range("j1", "k1").Merge(false);
                        //    xlWorkSheet.get_Range("j1", "k1").WrapText = true;
                        //    xlWorkSheet.get_Range("j1", "k1").EntireRow.Font.Bold = true;
                        //    xlWorkSheet.get_Range("j1", "k1").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                        //    xlWorkSheet.get_Range("j1", "k1").BorderAround();

                        //    for(int t = 1; t<= countSerial; t++)
                        //    {
                        //        if()
                        //        xlWorkSheet.Cells[1+t, 1] = t;
                        //        else
                        //        {

                        //        }    
                        //    }   

                        //    xlWorkSheet.PageSetup.CenterFooter = "AAAAAAAAAAAAAAAAAAAAAAAAAAAAA";
                        //}    



                        //string locationsavefilesaveas = strFilename + "\\" + "abcdef";
                        //xlWorkBook.SaveAs(locationsavefilesaveas);
                        xlWorkBook.SaveAs(strFullSNname, ExcelOffice.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, ExcelOffice.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);

                        //close and release 
                        xlWorkBook.Close();

                        Marshal.ReleaseComObject(xlWorkBook);

                        //quit and release  
                        xlApp.Quit();
                        Marshal.ReleaseComObject(xlApp);
                    }
                }
                return 1;
            } catch (Exception)
            {
                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
                return 0;
            }
        }

        public static readonly string[] VietnameseSigns = new string[]

        {

            "aAeEoOuUiIdDyY",

            "áàạảãâấầậẩẫăắằặẳẵ",

            "ÁÀẠẢÃÂẤẦẬẨẪĂẮẰẶẲẴ",

            "éèẹẻẽêếềệểễ",

            "ÉÈẸẺẼÊẾỀỆỂỄ",

            "óòọỏõôốồộổỗơớờợởỡ",

            "ÓÒỌỎÕÔỐỒỘỔỖƠỚỜỢỞỠ",

            "úùụủũưứừựửữ",

            "ÚÙỤỦŨƯỨỪỰỬỮ",

            "íìịỉĩ",

            "ÍÌỊỈĨ",

            "đ",

            "Đ",

            "ýỳỵỷỹ",

            "ÝỲỴỶỸ"

        };
        public static string RemoveSign4VietnameseString(string str)

        {

            //Tiến hành thay thế , lọc bỏ dấu cho chuỗi

            for (int i = 1; i < VietnameseSigns.Length; i++)

            {

                for (int j = 0; j < VietnameseSigns[i].Length; j++)

                    str = str.Replace(VietnameseSigns[i][j], VietnameseSigns[0][i - 1]);

            }

            return str;

        }

    }
}
