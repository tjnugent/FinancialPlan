using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace FinancialPlan
{
    public partial class FinPlan : Form
    {
        public readonly OpenFileDialog ofd = new OpenFileDialog();

        public int ReadExcelFile(string fileName, int sheetNum)
        {
            List<string> rowList = new List<string>();
            NPOI.SS.UserModel.ISheet sheet;

            if (IsFileLocked(fileName))
                return 1;

            using (var stream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                NPOI.SS.UserModel.IWorkbook book;
                int headerRowNum = 0;
                stream.Position = 0;
                book = new XSSFWorkbook(stream);
                sheet = book.GetSheetAt(sheetNum);
                NPOI.SS.UserModel.IRow headerRow = sheet.GetRow(0);
                //XSSFFormulaEvaluator formula = new XSSFFormulaEvaluator(book);
                int cellCount = headerRow.LastCellNum;

                if (dtTable != null)
                {
                    dtTable.Rows.Clear();
                    dtTable.Columns.Clear();
                }

                for (int k = 0; k < cellCount; k++)
                {
                    var headerCell = headerRow.GetCell(k);
                    if (headerCell != null && headerCell.CellType != CellType.Error && headerCell.CellType != CellType.Blank)
                    {
                        dtTable.Columns.Add(headerCell.StringCellValue);
                    }
                }

                for (int i = (sheet.FirstRowNum + headerRowNum + 1); i <= sheet.LastRowNum; i++)
                {
                    NPOI.SS.UserModel.IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        for (int j = row.FirstCellNum; j < cellCount; j++)
                        {
                            var cell = row.GetCell(j);
                            if (cell != null && cell.CellType != CellType.Error && cell.CellType != CellType.Blank)
                            {
                                if ((j == 0) && (cell.StringCellValue[0] == '^' || cell.StringCellValue[0] == '#'))
                                    break;

                                if (cell.CellType == CellType.Formula)
                                {
                                    try
                                    {
                                        //formula.EvaluateInCell(cell);
                                        XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(book);
                                        //fe.ClearAllCachedResultValues();
                                        //fe.EvaluateFormulaCell(cell);
                                        fe.EvaluateInCell(cell);
                                        rowList.Add(cell.NumericCellValue.ToString());
                                    }
                                    catch (Exception ex)
                                    {
                                        statusLabelValue.Text = ex.Message;
                                        rowList.Add(cell.NumericCellValue.ToString());
                                    }

                                }
                                else
                                    rowList.Add(cell.ToString());
                            }
                            else
                                break;
                        }
                    }
                    if (rowList.Count > 0)
                        dtTable.Rows.Add(rowList.ToArray());
                    rowList.Clear();
                }
                stream.Close();
            }
            return 0;
        }
    }
}