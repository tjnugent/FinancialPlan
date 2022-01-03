using System;
using System.IO;
using System.Data;
using System.Windows.Forms;
using System.Globalization;

namespace FinancialPlan
{
    public partial class FinPlan : Form
    {

        public const double MAXCPP = 2898;
        public const double cpprate = 0.0525;
        public const double MAXEI = 856.36;
        public const double eirate = 0.0158;
        public const double taxTier1 = 48535;
        public const double taxTier2 = 97069;
        public const double taxTier3 = 150476;
        public const double taxTier4 = 214368;

        private double MyDouble(string s)
        {
            NumberStyles style = NumberStyles.Float | NumberStyles.AllowLeadingSign | NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint |
                NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite | NumberStyles.AllowTrailingSign | NumberStyles.AllowThousands ;
            try
            {
                if (Double.TryParse(s, style, culture, out double value))
                    return (value);
                else if (s.Length == 0)
                    return (Double)0;
                else
                    return Double.Parse(s, style);
            }
            catch (Exception ex)
            {
                statusLabelValue.Text = ex.Message;
                return (Double)0;
            }
        }

        private int MyInt(string s)
        {
            NumberStyles style = NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite;

            if (Int32.TryParse(s, style, culture, out int value))
                return (value);
            else if (s.Length == 0)
                return (0);
            else
                return Int32.Parse(s);
        }

        private double MyRate(string s)
        {
            NumberStyles style = NumberStyles.AllowTrailingSign | NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingWhite | NumberStyles.AllowTrailingWhite;

            if (Double.TryParse(s, style, culture, out double value))
                return value;
            else if (s.Length == 0)
                return 0;
            else
                return MyDouble(s.Substring(0, s.Length - 1)) / 100;
        }

        private double CPP(double income)
        {
            return Math.Min(income * cpprate, MAXCPP);
        }

        private double EI(double income)
        {
            return Math.Min(income * eirate, MAXEI);
        }

        private double IncomeTax(double income)
        {
            double tax;

            if (income <= taxTier1)
                tax = income * 0.15;
            else
            {
                tax = taxTier1 * 0.15;
                if (income <= taxTier2)
                    tax += (taxTier2 - income) * 0.205;
                else
                {
                    tax += (taxTier2 - taxTier1) * 0.205;
                    if (income < taxTier3)
                        tax += (taxTier3 - income) * 0.26;
                    else
                    {
                        tax += (taxTier4 - taxTier3) * .29;
                        if (income > taxTier4)
                            tax += (income - taxTier4) * .33;
                    }

                }
            }
            return tax;
        }

        protected virtual bool IsFileLocked(string file)
        {
            FileStream stream = null;

            try
            {
                stream = File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to, or being processed by another thread, or does not exist (has already been processed)
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

        public DataTable GenerateTransposedTable(DataTable inputTable, string tableName)
        {
            DataTable outputTable = new DataTable();

            outputTable.Columns.Add(tableName);
            foreach (DataRow inRow in inputTable.Rows)
                outputTable.Columns.Add(inRow[0].ToString());

            for (int i = 1; i < inputTable.Columns.Count; i++)
                outputTable.Rows.Add(inputTable.Columns[i].ColumnName);
 
            for (int i = 1; i < outputTable.Rows.Count + 1; i++)
            {
                for (int j = 0; j < outputTable.Columns.Count - 1; j++)
                {
                    outputTable.Rows[i - 1][j+1] = inputTable.Rows[j][i];
                }
            }

            return outputTable;
        }
    }
}
