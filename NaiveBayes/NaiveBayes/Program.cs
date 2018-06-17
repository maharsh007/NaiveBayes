using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;


namespace NaiveBayes
{
    class Program
    {
        List<string> FirstProbabilityColumn=new List<string>();
        List<string> SecondProbabilityColumn=new List<string>();
        List<string> FirstExcelColumn = new List<string>();
        List<string> SecondExcelColumn = new List<string>();

        string probabilityFirst;
        string probabilitySecond;

        static char[] charSeparator = { ' ' };

        Dictionary<int, List<string>> TotalData=new Dictionary<int, List<string>>();
        static void Main(string[] argsOld)
        {
            var naiveBayes = new NaiveBayes.Program();

            naiveBayes.GetDataFromExcel();

            Console.WriteLine("Please enter your inputs.\n");

            var arguments = Console.ReadLine();

            string[] args = arguments.Split(charSeparator);

            if (args == null)
            {
                Console.WriteLine("Please enter the arguments properly.\n" +
                    " For example if you want a probability of x if y is present , please write x first and y later.\n" +
                    " The x and y should be fetched from the excel sheet and they should be present in the excel sheet in different columns");
                Console.ReadLine();

                return;
            }
            else
            {
                var firstArgumentFound = false;

                var countofArguments = 0;

                for (int i = 0; i < args.Length; i++)
                {
                    if (firstArgumentFound)
                    {
                        if(naiveBayes.FirstProbabilityColumn.Contains(args[i]))
                        {
                            naiveBayes.probabilityFirst = args[i];
                            countofArguments++;
                            break;
                        }
                    }
                    else
                    {
                        if (naiveBayes.FirstExcelColumn.Contains(args[i]))
                        {
                            naiveBayes.SecondProbabilityColumn.AddRange(naiveBayes.FirstExcelColumn);
                            naiveBayes.FirstProbabilityColumn.AddRange(naiveBayes.SecondExcelColumn);
                            naiveBayes.probabilitySecond = args[i];
                            countofArguments++;
                            firstArgumentFound = true;
                        }
                        else
                        {
                            if (naiveBayes.SecondExcelColumn.Contains(args[i]))
                            {
                                naiveBayes.FirstProbabilityColumn.AddRange(naiveBayes.FirstExcelColumn);
                                naiveBayes.SecondProbabilityColumn.AddRange(naiveBayes.SecondExcelColumn);
                                naiveBayes.probabilitySecond = args[i];
                                countofArguments++;
                                firstArgumentFound = true;
                            }
                        }
                    }
                }

                if (countofArguments<2)
                {
                    Console.WriteLine("We could not find sufficient number of arguments .\n");
                    Console.ReadLine();
                    return;
                }

            }

            Console.WriteLine($"The following arguments are taken {naiveBayes.probabilityFirst} and {naiveBayes.probabilitySecond}. Rest are ignored.\n");
            Console.ReadLine();

            decimal ProbabilityFirstTypesCount = naiveBayes.GetTotalOccuranceofProbabilityFirstTypes();
            decimal ProbabilitySecondTypesCount = naiveBayes.GetTotalOccuranceofProbabilitySecondTypes();
            decimal ProbabilityFirstinFirstColumn = naiveBayes.GetOccuranceofProbabilityFirstinFirstColumn();
            decimal ProbabilitySecondinSecondColumn = naiveBayes.GetOccuranceofProbablitySecondinSecondColumn();
            decimal ProbabilityFirstinProbabilitySecond = naiveBayes.GetOccuranceofProbabilityFirstinProbabilitySecond();
            decimal ProbabilitySecondinProbabilityFirst = naiveBayes.GetOccuranceofProbabilitySecondinProbabilityFirst();

            Console.WriteLine($"The probability of {naiveBayes.probabilitySecond} when {naiveBayes.probabilityFirst} is present " + naiveBayes.FindProbabilityofCWhenB(ProbabilitySecondTypesCount, ProbabilitySecondinSecondColumn, ProbabilityFirstinFirstColumn, ProbabilityFirstTypesCount, ProbabilityFirstinProbabilitySecond).ToString("0.00"));
            Console.ReadLine();
        }

        private void GetDataFromExcel()
        {
            var xlApp = new Application();
            var xlWorkBook = xlApp.Workbooks.Open(Filename: @"C:\Users\maharshi.choudhury\Data - Copy.xlsx");
            var xlWorkSheet = (Worksheet)xlWorkBook.Sheets["Sheet1"];
            var range = xlWorkSheet.UsedRange;
            var rowCount = range.Rows.Count;
            var columnsCount = range.Columns.Count;

            for (int row = 1; row <= rowCount; row++)
            {
                var str = (string)(range.Cells[row, 1] as Range).Value2;
                FirstExcelColumn.Add(str.ToLower());
            }

            for (int row = 1; row <= rowCount; row++)
            {
                var str = (string)(range.Cells[row, 2] as Range).Value2;
                SecondExcelColumn.Add(str.ToLower());
            }

            for (int row = 1; row <= rowCount; row++)
            {
                TotalData.Add(row, new List<string>());
                for (int col = 1; col <= columnsCount; col++)
                {
                    var str = (string)(range.Cells[row, col] as Range).Value2;
                    TotalData[row].Add(str.ToLower());
                }
            }


            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);
        }

        private void PopulateLocalObject()
        {

        }

        private decimal GetTotalOccuranceofProbabilityFirstTypes()
        {
            return FirstProbabilityColumn.Count;
        }

        private decimal GetTotalOccuranceofProbabilitySecondTypes()
        {
            return SecondProbabilityColumn.Count;
        }

        private decimal GetOccuranceofProbablitySecondinSecondColumn()
        {
            return SecondProbabilityColumn.Where(x => x.Equals(probabilitySecond,StringComparison.CurrentCultureIgnoreCase)).Count();
        }

        private decimal GetOccuranceofProbabilityFirstinFirstColumn()
        {
            return FirstProbabilityColumn.Where(x => x.Equals(probabilityFirst, StringComparison.CurrentCultureIgnoreCase)).Count();
        }

        private decimal GetOccuranceofProbabilityFirstinProbabilitySecond()
        {
            return TotalData.Where(x => x.Value.Contains(probabilitySecond.ToLower())).Where(x => x.Value.Contains(probabilityFirst.ToLower())).Count();
        }

        private decimal GetOccuranceofProbabilitySecondinProbabilityFirst()
        {
            return TotalData.Where(x => x.Value.Contains(probabilityFirst.ToLower())).Where(x => x.Value.Contains(probabilitySecond.ToLower())).Count();
        }

        private decimal FindProbabilityOfC(decimal totalCountC,decimal countC)
        {
            return (countC / totalCountC);
        }
        private decimal FindProbabilityOfBWhenC(decimal countC, decimal countBC)
        {
            return countBC / countC;
        }

        private decimal FindProbabilityofB(decimal totalCountB, decimal countB)
        {
            return countB / totalCountB;
        }

        private decimal FindProbabilityofCWhenB(decimal totalCountC, decimal countC, decimal countB, decimal countBC, decimal totalCountB)
        {
            return (FindProbabilityOfC(totalCountC, countC) * FindProbabilityOfBWhenC(countC, countBC)) / FindProbabilityOfC(totalCountB, countB);
        }
    }


}
