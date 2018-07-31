using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office;
using Excel = Microsoft.Office.Interop.Excel;




namespace ExamplePLIF
{
    class Functions
    {
        enum Columns { CPR = 1, COMPANY = 2, TYPE = 3, AMOUNT = 4 }; // select relevant columns from working file
        private static int SHEET = 2; // specific sheet
        private static int FROM_LINE = 2;
        

        private Dictionary<string, int> typesRef = new Dictionary<string, int>();
        //only the ref of the list 
        private Dictionary<string, Dictionary<string, List<double>>> companies = new Dictionary<string, Dictionary<string, List<double>>>();

               
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;


        public Functions()
        {
            xlApp = new Microsoft.Office.Interop.Excel.Application();
        }

        private void donotremove(string filepath) {
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SHEET); // second sheet for this file
                                                                                                           // the entire table :
            Microsoft.Office.Interop.Excel.Range range = xlWorkSheet.UsedRange; // range.Rows.Count, range.Columns.Count

            for (int col = 1; col <= range.Columns.Count; col++)
            {
                // this line does only one COM interop call for the whole column
                object[,] currentColumn = (object[,])range.Columns[col, Type.Missing].Value;
                double sum = 0;
                for (int i = 1; i < currentColumn.Length; i++)
                {
                    object val = currentColumn[i, 1];
                    //sum += (double)val; // only if you know that the values are numbers 
                    // if it was a string column :
                    Console.WriteLine(val);
                }
            }
        }
        public void insertPFA(string filepath, string outFile)
        {
            xlWorkBook = xlApp.Workbooks.Open(filepath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SHEET); // second sheet for this file

            readTypes(xlWorkSheet);
            Excel.Range range = xlWorkSheet.UsedRange;
            object[,] cpnyColumn = (object[,])range.Columns[Columns.COMPANY, Type.Missing].Value;
            object[,] cprColumn = (object[,])range.Columns[Columns.CPR, Type.Missing].Value;
            object[,] typeColumn = (object[,])range.Columns[Columns.TYPE, Type.Missing].Value;
            object[,] amountColumn = (object[,])range.Columns[Columns.AMOUNT, Type.Missing].Value;

            int row = FROM_LINE;
            // read row by row
            while (row < cpnyColumn.Length)
            {
                string companyName = (string)cpnyColumn[row, 1];
                Dictionary<string, List<double>> thisCompany;
                List<double> thisCPR;
                // checks if company already exists, if not, we create it
                if (!companies.ContainsKey(companyName))
                {
                    Console.WriteLine("new company " + companyName);
                    thisCompany = new Dictionary<string, List<double>>();
                    companies.Add(companyName, thisCompany);
                    

                }

                // checks if cpr already exists, if not we create it and put all values to 0
                thisCompany = companies[companyName];
                string cprValue = cprColumn[row, 1].ToString();
                if (!thisCompany.ContainsKey(cprValue)) {
                    thisCPR = new List<double>();
                    // put everything to 0
                    for (int i = 0; i < typesRef.Count; i++) {
                        thisCPR.Add(0);
                    }

                    thisCompany.Add(cprValue, thisCPR);
                }

                thisCPR = thisCompany[cprValue];

                // gather the type, the amount and put it in the list

                string typeName = (string)typeColumn[row, 1];
                double amount = (double)amountColumn[row, 1];
                thisCPR[typesRef[typeName]] += amount;

                row++;
            }

            xlWorkBook.Close();



            //insert obtained data into outputfile
            xlWorkBook = xlApp.Workbooks.Open(outFile, 0, false, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(SHEET); // first sheet for this file
            Boolean _continue = true;
            row = 5; // check the file befor deciding from where to start!!

                foreach (Dictionary<string, List<double>> nCompany in companies.Values) {
                    foreach( string cpr in nCompany.Keys)
                    {
                        xlWorkSheet.Cells[row, 2] = cpr; 
                        row++;
                    }
                    //////////////////////remember me////////////////
                }

            xlWorkBook.SaveAs(outFile, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkBook.Close();

        }
        public void readTypes(Excel.Worksheet sheet) {
            Excel.Range range = xlWorkSheet.UsedRange; 
            object[,] cprColumn = (object[,])range.Columns[Columns.TYPE, Type.Missing].Value;
            for (int i = FROM_LINE; i < cprColumn.Length; i++)
            {
                string typeName = (string) cprColumn[i, 1];
                if (!typesRef.ContainsKey(typeName)) {
                    Console.WriteLine("Inserting type " + typeName);
                    typesRef.Add(typeName, typesRef.Count);
                }
            }
            Console.WriteLine("There is a total of " + typesRef.Count + " different types");
        }

        public void readCPRs(string filepath) {
           
                                                                                                           // the entire table :
            Excel.Range range = xlWorkSheet.UsedRange; // range.Rows.Count, range.Columns.Count
            object[,] cprColumn = (object[,])range.Columns[Columns.CPR, Type.Missing].Value;
            for (int i = 1; i < cprColumn.Length; i++)
            {
                object val = cprColumn[i, 1];
                Console.WriteLine(val);
            }
        }

       
    }

    //initialize classes for extraction data from the first test file -> skip this one 
    /*    public class ClientModelPFA
    {
        public int CPR { get; set; }
        public ColumnModel ColumnModel { get; set; }

    }

    public class ColumnModel
    {
        public string ColumnName { get; set; }
        public int ColumnValue { get; set; }
    }

    public class ConverterPFA
    {
        Dictionary<string, List<ClientModelPFA>> InputModel = new Dictionary<string, List<ClientModelPFA>>
            {
                {"test.xls", new List<ClientModelPFA>
                    {
                        new ClientModelPFA {CPR = 60606043, ColumnModel = new ColumnModel() {ColumnName = "VederlagEtablering", ColumnValue = -4000}},
                         new ClientModelPFA {CPR = 60606043, ColumnModel = new ColumnModel() {ColumnName = "VederlagLøbende", ColumnValue = 0}},
                    }

                }
            };
    }*/
}

