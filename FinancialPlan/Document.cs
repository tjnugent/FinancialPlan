using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Excel.FinancialFunctions;
using System.Globalization;
using Microsoft.Office.Interop.Word;

namespace FinancialPlan
{
    public partial class FinPlan : Form
    {
        List<string> subsections = new List<string> {
            "Younger single/couple, no children",
            "Single/Couple, Dependent Children",
            "Older Single/Couple, dependent children",
            "Older Single/Couple, Children Independent or nearly so",
            "Single/Couple, Retired, Empty Nest"
        };
        Microsoft.Office.Interop.Word.Application app = new Microsoft.Office.Interop.Word.Application();
        DataGridView dgvNetWorth = new DataGridView();
        readonly NumberStyles styles = NumberStyles.AllowCurrencySymbol | NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign | NumberStyles.AllowTrailingSign | NumberStyles.AllowThousands | NumberStyles.AllowTrailingWhite;
        object styleHeading1 = "Heading 1";
        object styleHeading2 = "Heading 2";
        object styleHeading3 = "Heading 3";
        object styleDocument = "Normal";
        object missing = System.Reflection.Missing.Value;
        Document document = null;

        private void CreateDocument()
        {
            object filename = fName;
            object behavior = WdDefaultListBehavior.wdWord10ListBehavior;

            statusLabelValue.Text = "Starting word document";
            
            app.Visible = false;
            document = new Document();

            SetHeaderFooter();
            document.Content.SetRange(0, 0);
            document.Content.Text = "A planning document for financial success. " + Environment.NewLine;
            document.Words.Last.InsertBreak(WdBreakType.wdPageBreak);

            if (IsFileLocked(fName))
            {
                statusLabelValue.Text = "File is locked, cannot save";
                return;
            }
            else
                document.SaveAs2(ref filename);

            ExecutiveSummary();
            document.Save();
            Assumptions();
            document.Save();
            Networth();
            document.Save();
            AccommodationsAndDebt();
            document.Save();
            PrintInsurance();
            document.Save();
            FinancingRetirement();
            document.Save();
            Conclusion();
            document.Save();
            document.Close(ref missing, ref missing, ref missing);
            document = null;
            app.Quit(ref missing, ref missing, ref missing);
            app = null;
        }

        private void Conclusion()
        {
            statusLabelValue.Text = "Setting up Conclusion";
            PrintStatement(styleHeading1, "Conclusion");
            PrintStatement(styleDocument, "We can see that if " + clientNameTextBox.Text + " can stick to the plan here, they will be in great shape for retirement. " +
                "As " + clientNameTextBox.Text + " enters retirement, we can easily see how retirement is funded for an excellent standard of living. " + clientNameTextBox.Text +
                " will have a significant legacy if they stick to the plan.");
            statusLabelValue.Text = "Document created successfully!";
        }
        private void Networth()
        {
            //System.Data.DataTable dt;

            statusLabelValue.Text = "Setting up Current Net Worth";
            PrintStatement(styleHeading1, "Current Net Worth");
            PrintStatement(styleDocument, "Net worth is the value of the assets a person or couple owns, minus the liabilities they owe. It is an important " +
                "metric to gauge a person's financial health, providing a useful snapshot of their current financial position. The information provided in the table below " +
                "is as discussed in the initial client interview and follow up conversations. ");
            //dt = BuildNetWorthTable();
            InsertTable(BuildNetWorthTable(), "Net Worth");
            //document.Content.Paragraphs.Add(ref missing).Range.InsertBreak(WdBreakType.wdPageBreak);
        }

        private void Assumptions()
        {
            //System.Data.DataTable dt;

            statusLabelValue.Text = "Setting up Assumptions";
            //myParagraph = document.Content.Paragraphs.Add(ref missing);
            PrintStatement(styleHeading1, "Assumptions");
            PrintStatement(styleDocument, "For planning purposes some assumptions have had to be made. At time of writing these are most realistic and conservative. " +
                "Any forward-looking statements and assumptions can change and would require updating this document. The following is a list of assumptions used in this document:");
            string[] bullets =
            {
                "The minimum income level while remaining within the employment phase is now. Income will only increase over time.",
                "There will be no debt at the time of retirement.",
                "The cost-of-living increase will be " + cpprate.ToString("P02", culture) + " year-over-year.",
                "Government of Canada Pensions will be indexed at " + cpprate.ToString("P02", culture) + ".",
                "Government rules will not radically change over the course of this plan.",
                "Non-registered income is just moving money and therefore has no tax consequences.",
                clientName.Text + " will live until they both reach age 90 at which time, they will expire peacefully in their sleep."
            };
            PrintBullets(bullets);
            PrintStatement(styleDocument, "The table below shows the summary of assumptions:");
            InsertTable(BuildAssumptionsTable(), "Assumptions");
            //document.Content.Paragraphs.Add(ref missing).Range.InsertBreak(WdBreakType.wdPageBreak);
        }
        private void ExecutiveSummary()
        {
            statusLabelValue.Text = "Setting up Preliminaries";

            PrintStatement(styleHeading1, "Executive Summary");
            PrintStatement(styleDocument, "This document has been written to provide a very high-level financial plan. While there may be many decisions made " +
                "differently, this document shows one method for planning " + clientName.Text + "'s financial future, when to begin old age security, and how much to add/remove " +
                "from their various investment sources. While the plan will change from time to time, the basic outline will likely remain the same.");
            PrintStatement(styleDocument, "The plan has " + clientName.Text + " continuing to live in their existing home, taking Canada Pension at age 65.");
            //document.Content.Paragraphs.Add(ref missing).Range.InsertBreak(WdBreakType.wdPageBreak);
        }

        private void SetHeaderFooter()
        {
            foreach (Section section in document.Sections)
            {
                //Get the header range and add the header details.  
                Range headerRange = section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                //headerRange.Fields.Add(headerRange, WdFieldType.wdFieldNumPages);
                headerRange.Fields.Add(headerRange, WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft;
                //headerRange.PageSetup.DifferentFirstPageHeaderFooter = true;

                headerRange.Font.Name = "Franklin Gothic Book";
                headerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Text = "Financial Plan";

                Range footerRange = section.Footers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = WdColorIndex.wdBlue;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = "Confidential";
            }
        }

        private void FinancingRetirement()
        {
            bool dependants = !dependantChildrenTextBox.Text.Equals("0");
            bool married = marriedCheckBox.Checked;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            int index = 0;
            string roi = MyDouble(roiTextBox.Text).ToString("P02", culture);

            statusLabelValue.Text = "Setting up Financing Retirement";
            PrintStatement(styleHeading1, "Financing Retirement");
            if (age < 40 && !dependants)                                // Younger single or couple, no children
                index = 0;
            else if (age < 40 && dependants)                            // Single or Couple, Children
                index = 1;
            else if (age >= 40 && age <= retirementAge && !dependants)  // Single or Couple, dependent children
                index = 2;
            else if (age >= 40 && age <= retirementAge && dependants)   // Older Single/Couple, children independent or nearly so
                index = 3;
            else if (age > retirementAge)                               // Single/Couple, retired
                index = 4;

            
            PrintStatement(styleHeading2, subsections[index]);
            PrintStatement(styleDocument, "Our assumption of Market investments providing a return of " + roi + " annually on average is key to living off the proceeds of the " +
                "RSP. Less than " + roi + " changes the plan as the RSP will run out sooner. In addition, the plan assumes minimum contributions of " +
                monthlyContributions.ToString("C02", culture) + " per month from now to " + clientName.Text + "’s retirement. This savings is key to the plan as it seeds the TFSA " +
                "and adds to the RSP significantly. Should either of these assumptions prove false, we should revisit the plan. ");
            PrintStatement(styleDocument, "Decreasing discretionary spending, paying less in taxes, and possibly attaining the guaranteed income supplement (GIS) would all leave " +
                "room for flexibility in retirement. Another assumption is that " + clientName.Text + " puts off taking CPP until retirement at age " + retirementAgeTextBox.Text + 
                " wouldn’t make much sense to be collecting CPP while still in a very high tax margin. Better to let it increase by 0.6% per month until " + clientName.Text + " retires.");
            PrintStatement(styleDocument, "Leading into retirement, " + clientName.Text + " will " +
                "contribute $XXX per month into TFSA accounts, and $XXXX per month into his RSP accounts. We’ve skewed the higher amount to the RSP to maximum RSP employer matching. " +
                "This allows a balance of current-year tax savings combined with future - year tax relief. For simplicity, the RESP contributions end upon " + clientName.Text + "’s " +
                "retirement.");

            Range p = document.Content.Paragraphs.Add(ref missing).Range;
            p.InsertBreak(WdBreakType.wdSectionBreakContinuous);
            p.PageSetup.Orientation = WdOrientation.wdOrientLandscape;
            p.InsertParagraphAfter();

            statusLabelValue.Text = "Setting up Financing Retirement Pre-retirement";
            if (age <= retirementAge)
                 PrintRetirementPhase(age, age, retirementAge + 1, index, "Pre-Retirement");
            document.Save();
            statusLabelValue.Text = "Setting up Financing Retirement post-retirement";
            if (age < rifAge)
                PrintRetirementPhase(age, retirementAge + 1, rifAge, index, "Post Retirement/Pre RIF-Conversion");
            document.Save();
            statusLabelValue.Text = "Setting up Financing Retirement post rif-conversion";
            if (age < deathAge)
                PrintRetirementPhase(age, rifAge, deathAge+1, index, "Post RIF-Conversion");

            p = document.Content.Paragraphs.Add(ref missing).Range;
            p.InsertBreak(WdBreakType.wdSectionBreakContinuous);
            p.PageSetup.Orientation = WdOrientation.wdOrientPortrait;
            p.InsertParagraphAfter(); 
        }

        private void PrintRetirementPhase(int age1, int age2, int age3, int index, string heading)
        {
            PrintStatement(styleHeading2, heading);
            PrintStatement(styleHeading3, "Budget");
            InsertTableSection(age1, age2, age3, budgetdt, "Budget");
            PrintStatement(styleHeading3, "Pensions");
            InsertTableSection(age1, age2, age3, BasicPensions, "Pensions");
            PrintStatement(styleHeading3, "The LIRA/LIF Account");
            InsertTableSection(age1, age2, age3, LIRALIFTable, "LIRA/LIF");
            PrintStatement(styleHeading3, "The RSP/RIF Account");
            InsertTableSection(age1, age2, age3, rspTable, "RSP/RIF");
            PrintStatement(styleHeading3, "The TFSA Account");
            InsertTableSection(age1, age2, age3, tfsaTable, "TFSA");
            PrintStatement(styleHeading3, "The Non-Registered Account");
            InsertTableSection(age1, age2, age3, nonRegTable, "Non-Registered");
        }

        private void AccommodationsAndDebt()
        {
            string s;
            int age = MyInt(currentAgeTextBox.Text);
            int retireAge = MyInt(retirementAgeTextBox.Text);
            //System.Data.DataTable dt;
            bool dependents = (MyInt(dependantChildrenTextBox.Text) > 0 || MyDouble(spouseIncomeTextBox.Text) > 0 || MyDouble(spousalRSPBalanceTextBox.Text) > 0 || MyDouble(spousalTFSABalanceTextBox.Text) > 0);

            statusLabelValue.Text = "Setting up Living Accommodations & Debt";
            PrintStatement(styleHeading1, "Living Accommodations & Debt");
            PrintStatement(styleDocument, "To create wealth, the years with the largest income are assumed to be prior to age 60. It is during these years that " +
                    "the largest income will support contributions to costs such as Mortgages, vacation homes, retirement home, Children, Education, and retirement savings. The " +
                    "primary home will likely be maintained until age " + retirementAgeTextBox.Text + " when it will be downsized, and the profits added to the non-registered " +
                    "investment.");

            if (age < 30 && !dependents)
            {
                PrintStatement(styleDocument, "Plans for a vacation property or later, a semi - retirement home in a warmer climate have not " +
                    "been identified yet.It should be noted that all assets should be paid for in full prior to retirement, or else sold and/ or downsized such " +
                    "that no debt remains. Entering retirement while paying a monthly amount for debt, unless attempted to take advantage of the cost of money, " +
                    "will often create a retirement too meager to be enjoyed. There are no children living in the primary home, nor is there any education funding " +
                    "requirements. ");
                if (HomeOwnedCheckBox.Checked)
                    PrintStatement(styleDocument, "At the current stage, we are in our starter home, which has " +
                    MyDouble(primaryMortgageAmortTextBox.Text).ToString("C02", culture) + " " + "remaining on the mortgage. There is likely a home upgrade " +
                    "planned in the future. ");
                else
                    PrintStatement(styleDocument, "At the current stage, we are renting our home. ");
                PrintStatement(styleDocument, "In most plans, the most significant debt holdings are for housing. Debt with an asset to back " +
                    "it is fine and is often part of the plan. Debt with no asset, such as credit card debt, should be paid off as the highest priority. ");
            }
            s = clientName.Text;
            if (MyDouble(lineOfCreditAmtTextBox.Text) == 0 && MyDouble(creditCardDebtTextBox.Text) == 0)
                s += " has no debt, which is excellent. ";
            else if ((!primaryLocCheckBox.Checked && MyDouble(lineOfCreditAmtTextBox.Text) != 0) || MyDouble(creditCardDebtTextBox.Text) != 0)
                s += " have debt. If possible, try to move debt to a lower interest style of loan, such as a home equity line of credit. ";
            else
                s += " have no debt without an asset, which is excellent. ";
            PrintStatement(styleDocument, s);

            if (age >= 30 && age < 55 && !dependents)
                s = "At the current stage, " + clientName.Text + " are likely wanting to contribute to tax deferment plans. ";
            else if (age >= 30 && age < 55 && dependents)
                s = "At the current stage, " + clientName.Text + " likely don't have any cash left over at the end of each month. ";
            else if (age >= 55 && age < retireAge && !dependents)
                s = "At the current stage, " + clientName.Text + " should be reaching, or have already reached financial independence. ";
            else if (age >= 55 && age < retireAge && dependents)
                s = "At the current stage, " + clientName.Text + " should be approaching the empty nest stage, and reaching financial independence. ";
            else if (age > retireAge)
                s = "At the current stage, " + clientName.Text + " should be retired. ";
            PrintStatement(styleDocument, s);

            //dt = BuildDebtTable();
            InsertTable(BuildDebtTable(), "Debt");

            if (!primaryLocCheckBox.Checked && HomeOwnedCheckBox.Checked && (MyDouble(tfsaBalanceTextBox.Text) + MyDouble(cashBalanceTextBox.Text)) < 20000)
                PrintStatement(styleDocument, "Consider opening a Home Equity Line of Credit (HELOC) to use for an emergency fund. Once your " +
                    "TFSA and/or non-registered savings are built up you can close this out.");
            if (MyDouble(primaryMortgageAmortTextBox.Text) > 0)
            {
                PrintStatement(styleDocument, "The only significant debt holdings should be for housing. Debt that is covered with an asset " +
                    "is not necessarily bad and, in this case, is part of the plan. The table below shows the mortgage debt decrease over time, and the tables " +
                    "in the next section show the assets increase in value at a much faster rate.");
                PrintStatement(styleDocument, "As you can see the mortgage is paid off in " + mortgageTable.Rows.Count + " years. Prepayment " +
                    "is not recommended as we believe investment returns will continue to be higher than the cost of the mortgage debt.");
                InsertMortgageTable(mortgageTable, "Mortgage");
                PrintStatement(styleDocument, "Choosing between paying off debt vs investing can be as simple as looking at the interest rate. " +
                    "With 19.5% interest on debt, and 9% return on investment, it’s a no-brainer. When we shift the paradigm to 3% interest on debt, it is far " +
                    "less of a concern to pay it down quickly.");
            }
            //document.Content.Paragraphs.Add(ref missing).Range.InsertBreak(WdBreakType.wdPageBreak);
        }
        private void PrintStatement(object stmtStyle, string statement)
        {
            Range q = document.Content.Paragraphs.Add(ref missing).Range;
            q.set_Style(ref stmtStyle);
            q.Text = statement;
            q.InsertParagraphAfter();
        }

        private void PrintBullets(string[] statement)
        {
            object behavior = WdDefaultListBehavior.wdWord10ListBehavior;
            Paragraph p = null;
            bool appliedListFormat = false;

            foreach (string s in statement)
            {
                p = document.Content.Paragraphs.Add();
                p.Range.Text = s;
                if (!appliedListFormat)
                {
                    //p.Range.ListFormat.ApplyBulletDefault(WdDefaultListBehavior.wdWord10ListBehavior);
                    appliedListFormat = true;
                }

                p.Outdent();
                p.Range.InsertParagraphAfter();
            }
            p.Range.Delete();
        }

        private void InsertMortgageTable(System.Data.DataTable which, string label)
        {
            Range q = document.Content.Paragraphs.Add(ref missing).Range;

            Table table = document.Tables.Add(q, which.Rows.Count, which.Columns.Count, ref missing, ref missing);
            table.set_Style("List Table 7 Colorful - Accent 5");
            table.AllowAutoFit = true;
            table.Borders.Enable = 1;
            table.Borders.InsideColor = WdColor.wdColorBlue;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            table.Title = label;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            table.ApplyStyleRowBands = true;

            foreach (Cell cell in table.Rows[1].Cells)
            {
                cell.Range.Text = which.Columns[cell.ColumnIndex - 1].ColumnName;
                cell.Range.Font.Bold = 1;
                cell.Range.Font.Name = "Arial (Body)";
                cell.Range.Font.Size = 8;
            }

            try
            {
                for (int i = 2; i < table.Rows.Count; i++)
                {
                    foreach (Cell cell in table.Rows[i].Cells)
                    {
                        switch (cell.ColumnIndex - 1)
                        {
                            case 0:
                            case 1:
                                cell.Range.Text = MyInt(which.Rows[i - 2][cell.ColumnIndex - 1].ToString()).ToString("#", culture);
                                break;
                            default:
                                cell.Range.Text = MyDouble(which.Rows[i - 2][cell.ColumnIndex - 1].ToString()).ToString("C0", culture);
                                break;
                        }
                        cell.Range.Font.Bold = 0;
                        cell.Range.Font.Name = "Arial (Body)";
                        cell.Range.Font.Size = 7;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }

                Column firstCol = table.Columns[1];
                firstCol.AutoFit(); // force fit sizing
                Single firstColAutoWidth = firstCol.Width; // store autofit width
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent); // fill page width
                firstCol.SetWidth(firstColAutoWidth, WdRulerStyle.wdAdjustFirstColumn);
                table.ApplyStyleDirectFormatting("List Table 7 Colorful - Accent 5");
            }
            catch (Exception ex)
            {
                statusLabelValue.Text = ex.Message;
            }
            q.InsertParagraphAfter();
        }

        private void InsertTableSection(int currentAge, int startAge, int endAge, System.Data.DataTable dt, string label)
        {
            dt.Columns["Year"].SetOrdinal(0);
            dt.Columns["Age"].SetOrdinal(1); 
            System.Data.DataTable which = GenerateTransposedTable(dt, label);
            Range q = document.Content.Paragraphs.Add(ref missing).Range;

            statusLabelValue.Text = "Setting up Section " + label;
            Table table = document.Tables.Add(q, which.Rows.Count + 1, endAge - startAge + 1, ref missing, ref missing);
            table.set_Style("List Table 7 Colorful - Accent 5");
            table.AllowAutoFit = true;
            table.Borders.Enable = 1;
            table.Borders.InsideColor = WdColor.wdColorBlue;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            table.Title = label;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            table.ApplyStyleRowBands = true;

            try
            {
                for (int j = 1; j < table.Rows[1].Cells.Count + 1; j++)             // Column Labels
                {
                    table.Rows[1].Cells[j].Range.Text = which.Columns[startAge - currentAge + j - 1].ColumnName;
                    table.Rows[1].Cells[j].Range.Font.Bold = 1;
                    table.Rows[1].Cells[j].Range.Font.Name = "Arial (Header)";
                    table.Rows[1].Cells[j].Range.Font.Size = 9;
                }

                for (int i = 2; i < table.Rows.Count + 1; i++)                      // Copy Row Labels + Data
                {
                    int year = MyInt(which.Columns[startAge - currentAge + 1].ColumnName);
                    table.Rows[i].Cells[1].Range.Text = which.Rows[i - 2][0].ToString();
                    table.Rows[i].Cells[1].Range.Font.Bold = 1;
                    table.Rows[i].Cells[1].Range.Font.Name = "Arial (Header)";
                    table.Rows[i].Cells[1].Range.Font.Size = 9;

                    for (int j = 2; j < table.Rows[i].Cells.Count + 1; j++)
                    {
                        table.Rows[i].Cells[j].Range.Text = MyDouble(which.Rows[i - 2][year++.ToString()].ToString()).ToString("#", culture);
                        table.Rows[i].Cells[j].Range.Font.Bold = 0;
                        table.Rows[i].Cells[j].Range.Font.Name = "Arial (Body)";
                        table.Rows[i].Cells[j].Range.Font.Size = 8;
                        table.Rows[i].Cells[j].Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }
                table.Rows[1].Cells[1].Range.Text = label;
                table.Columns[1].AutoFit();
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                //table.Columns[1].SetWidth(table.Columns[1].Width, WdRulerStyle.wdAdjustFirstColumn);
                table.ApplyStyleDirectFormatting("List Table 7 Colorful - Accent 5");
            }
            catch (Exception ex)
            {
                statusLabelValue.Text = ex.Message;
            }
            q.InsertParagraphAfter();
            //q.InsertBreak(WdBreakType.wdPageBreak);
        }
        
        private System.Data.DataTable BuildDebtTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            Double value, rate;

            dt.Columns.Add("Debt");
            dt.Columns.Add("Amount");
            dt.Columns.Add("Rate");
            dt.Columns.Add("Comment");

            value = MyDouble(creditCardDebtTextBox.Text);
            rate = MyRate(creditCardRateTextBox.Text);
            dt.Rows.Add(new Object[] { "Credit Card", value.ToString("C02", culture), rate.ToString("P02", culture), String.Empty });

            value = MyDouble(lineOfCreditAmtTextBox.Text);
            rate = MyRate(lineOfCreditRateTextBox.Text);
            dt.Rows.Add(new Object[] { "Loans", value.ToString("C02", culture), rate.ToString("P02", culture), String.Empty });

            value = MyDouble(carPmtTextBox.Text);
            dt.Rows.Add(new Object[] { "Car Payment", String.Empty, String.Empty, value.ToString("C02", culture) + " Monthly" });

            return (dt);
        }

        private System.Data.DataTable BuildAssumptionsTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            double pmt = 0, rate = 0, value;
            int period;
            string s;

            dt.Columns.Add(clientName.Text);
            dt.Columns.Add("Current Amount");
            dt.Columns.Add("Interest Rate");
            dt.Columns.Add("Monthly Payment");

            value = MyDouble(primaryMortgageAmtTextBox.Text); 
            rate = MyRate(primaryMortgageRateTextBox.Text);
            period = (DateTime.Now.Month - primaryHomePurchaseDatePicker.Value.Month) + (DateTime.Now.Year - primaryHomePurchaseDatePicker.Value.Year) * 12;
            /*
            if (value > 0)
                pmt = Financial.Pmt(rate, period, value, 0, PaymentDue.EndOfPeriod);
            */
            dt.Rows.Add(new Object[] { "Mortgage", value, rate.ToString("P02", culture), primaryMortgagePmt.ToString("C02", culture) });

            value = MyDouble(lineOfCreditAmtTextBox.Text);
            rate = MyRate(lineOfCreditRateTextBox.Text);
            if (value > 0)
                pmt = value * rate / 365 * DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month);
            dt.Rows.Add(new Object[] { "Line of Credit", value.ToString("C00", culture), rate.ToString("P02", culture), pmt.ToString("C02", culture) });

            s = roiCheckBox.Checked ? "Age Based" : roiTextBox.Text;
            dt.Rows.Add(new Object[] { "Return on Investment", String.Empty, s, String.Empty });

            dt.Rows.Add(new Object[] { "Age", "Current: " + currentAgeTextBox.Text, "Retire: " + retirementAgeTextBox.Text, "Death: " + deathAgeTextBox.Text });

            value = MyDouble(incomeTextBox.Text) + MyDouble(spouseIncomeTextBox.Text);
            dt.Rows.Add(new Object[] { "Current Income", value.ToString("C00", culture), String.Empty, value / 12 });

            value = MyDouble(rspBalanceTextBox.Text) + MyDouble(spousalRSPBalanceTextBox.Text);
            dt.Rows.Add(new Object[] { "RSP", value.ToString("C00", culture), String.Empty, String.Empty });

            value = MyDouble(tfsaBalanceTextBox.Text) + MyDouble(spousalTFSABalanceTextBox.Text);
            dt.Rows.Add(new Object[] { "TFSA", value.ToString("C00", culture), String.Empty, String.Empty });

            value = MyDouble(cashBalanceTextBox.Text) + MyDouble(spouseCashBalanceTextBox.Text) + MyDouble(jointCashBalanceTextBox.Text);
            dt.Rows.Add(new Object[] { "Non-Registered", value.ToString("C00", culture), String.Empty, String.Empty });

            dt.Rows.Add(new Object[] { "Cost of Living", String.Empty, costOfLiving.ToString("P02", culture), String.Empty });
            dt.Rows.Add(new Object[] { "Government Claw back/Index", "$70,954", "2.50%", String.Empty });

            return (dt);
        }

        private System.Data.DataTable BuildNetWorthTable()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            double assets = 0, liabilities = 0;

            dt.Columns.Add("Name");
            dt.Columns.Add("Assets");
            dt.Columns.Add("Liabilities");
            if (HomeOwnedCheckBox.Checked)
            {
                assets = String.IsNullOrEmpty(primaryHomeValueTextBox.Text) ? 0 : Double.Parse(primaryHomeValueTextBox.Text);
                liabilities = String.IsNullOrEmpty(primaryMortgageAmtTextBox.Text) ? 0 : Double.Parse(primaryMortgageAmtTextBox.Text);
                dt.Rows.Add(new Object[] { "Primary Home", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });
            }

            assets = 0;
            liabilities = String.IsNullOrEmpty(lineOfCreditAmtTextBox.Text) ? 0 : Double.Parse(lineOfCreditAmtTextBox.Text);
            dt.Rows.Add(new Object[] { "Line of Credit", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            assets = 0;
            liabilities = String.IsNullOrEmpty(creditCardDebtTextBox.Text) ? 0 : Double.Parse(creditCardDebtTextBox.Text);
            dt.Rows.Add(new Object[] { "Credit Card Debt", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            if (investPropertyCheckBox.Checked)
            {
                assets = String.IsNullOrEmpty(investPropertyValueTextBox.Text) ? 0 : Double.Parse(investPropertyValueTextBox.Text);
                liabilities = String.IsNullOrEmpty(investMortgageAmountTextbox.Text) ? 0 : Double.Parse(investMortgageAmountTextbox.Text);
                dt.Rows.Add(new Object[] { "Investment Property", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });
            }

            assets = 0;
            liabilities = String.IsNullOrEmpty(emergencyFundTextBox.Text) ? 0 : Double.Parse(emergencyFundTextBox.Text);
            dt.Rows.Add(new Object[] { "Emergency Fund", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            assets = String.IsNullOrEmpty(rspBalanceTextBox.Text) ? 0 : Double.Parse(rspBalanceTextBox.Text);
            liabilities = 0;
            dt.Rows.Add(new Object[] { "RSP Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            assets = String.IsNullOrEmpty(spousalRSPBalanceTextBox.Text) ? 0 : Double.Parse(spousalRSPBalanceTextBox.Text);
            liabilities = 0;
            if (marriedCheckBox.Checked)
                dt.Rows.Add(new Object[] { "Spouses RSP Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            assets = String.IsNullOrEmpty(tfsaBalanceTextBox.Text) ? 0 : Double.Parse(tfsaBalanceTextBox.Text);
            liabilities = 0;
            dt.Rows.Add(new Object[] { "TFSA Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            if (marriedCheckBox.Checked)
            {
                assets = String.IsNullOrEmpty(spousalTFSABalanceTextBox.Text) ? 0 : Double.Parse(spousalTFSABalanceTextBox.Text);
                liabilities = 0;
                dt.Rows.Add(new Object[] { "Spouses TFSA Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });
            }
            assets = String.IsNullOrEmpty(cashBalanceTextBox.Text) ? 0 : Double.Parse(cashBalanceTextBox.Text);
            liabilities = 0;
            dt.Rows.Add(new Object[] { "Cash Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

            if (marriedCheckBox.Checked)
            {
                assets = String.IsNullOrEmpty(spouseCashBalanceTextBox.Text) ? 0 : Double.Parse(spouseCashBalanceTextBox.Text);
                liabilities = 0;
                dt.Rows.Add(new Object[] { "Spouses Cash Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });

                assets = String.IsNullOrEmpty(jointCashBalanceTextBox.Text) ? 0 : Double.Parse(jointCashBalanceTextBox.Text);
                liabilities = 0;
                dt.Rows.Add(new Object[] { "Joint Cash Balance", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });
            }

            assets = liabilities = 0;
            foreach (DataRow r in dt.Rows)
            {
                if (!string.IsNullOrEmpty(r["Assets"].ToString()))
                    assets += Double.Parse(r["Assets"].ToString(), styles);
                if (!string.IsNullOrEmpty(r["Liabilities"].ToString()))
                    liabilities += Double.Parse(r["Liabilities"].ToString(), styles);
            }
            dt.Rows.Add(new Object[] { "Totals", assets.ToString("C0", culture), liabilities.ToString("C0", culture) });
            if (assets - liabilities > 0)
                dt.Rows.Add(new Object[] { "Net Worth", (assets - liabilities).ToString("C0", culture), "" });
            else
                dt.Rows.Add(new Object[] { "Net Worth", "", (liabilities - assets).ToString("C0", culture) });

            return(dt);
        }

        private void InsertTable(System.Data.DataTable dgv, string title)
        {
            Range q = document.Content.Paragraphs.Add(ref missing).Range;

            Table table = document.Tables.Add(q, dgv.Rows.Count+1, dgv.Columns.Count, ref missing, ref missing);
            table.set_Style("List Table 7 Colorful - Accent 5");
            table.AllowAutoFit = true;
            table.Borders.Enable = 1;
            table.Borders.InsideColor = WdColor.wdColorBlue;
            table.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone;
            table.Title = title;
            table.Rows.Alignment = WdRowAlignment.wdAlignRowCenter;
            table.ApplyStyleRowBands = true;

            try
            {
                foreach (Cell cell in table.Rows[1].Cells)
                {
                    cell.Range.Text = dgv.Columns[cell.ColumnIndex - 1].ColumnName;
                    cell.Range.Font.Bold = 1;
                    cell.Range.Font.Name = "Arial (Body)";
                    cell.Range.Font.Size = 8;
                }

                for (int i = 2; i < table.Rows.Count + 1; i++)
                {
                    foreach (Cell cell in table.Rows[i].Cells)
                    {
                        cell.Range.Text = dgv.Rows[cell.RowIndex - 2][cell.ColumnIndex - 1].ToString();
                        cell.Range.Font.Bold = 0;
                        cell.Range.Font.Name = "Arial (Body)";
                        cell.Range.Font.Size = 7;
                        cell.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight;
                    }
                }
                table.Rows[1].Cells[1].Range.Text = string.Empty;
                table.Columns[1].AutoFit();
                table.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitContent);
                table.ApplyStyleDirectFormatting("List Table 7 Colorful - Accent 5");
            }
            catch (Exception ex)
            {
                statusLabelValue.Text = ex.Message;
            }
            q.InsertParagraphAfter();
        }

        private void PrintInsurance()
        {
            statusLabelValue.Text = "Setting up Insurance";
            PrintStatement(styleHeading1, "Insurance");
            PrintStatement(styleHeading2, "Term Life Insurance");
            PrintStatement(styleDocument, "If one party passes and the spouse or the children are left with present and future debts, it holds that the amount of term life " +
                "coverage combined with savings should be enough to cover these debts plus any living costs required. It makes sense then that the more wealth accumulated, as well " +
                "as the more self-sufficient the children, the less term life insurance is necessary. Certainly, all term life premiums should be gone at the retirement stage, " +
                "especially when we follow the methodology of no debt carried into retirement. Finally, Universal life insurance policies need to be heavily funded in the early days " +
                "to build enough capital to overcome the increased insurance costs as we grow older. Without early funding, the term-life costs overcome the benefit of the savings vehicle.");
            PrintStatement(styleHeading2, "Accident Insurance");
            PrintStatement(styleDocument, "Accident insurance does not often make sense to carry. Typically, the requirements for a payout are so stringent that we rarely get " +
                "the benefit of this anyway. Likely the cost outweighs the benefit, and so there is no consideration for Accident insurance in this plan.");
            PrintStatement(styleHeading2, "Disability Insurance");
            PrintStatement(styleDocument, "We consider this type of insurance to be income replacement should something happen, that would restrict us from working. It should " +
                "be purchased in its basic form, and the cost - of - living adjustment portion should be included. Usually, this is available, and often not optional, from your " +
                "employer and that amount is considered sufficient.");
            PrintStatement(styleHeading2, "Long Term Care Insurance");
            PrintStatement(styleDocument, "If you need extra care such that someone qualified needs to come into the home regularly, or perhaps you need a higher care facility, " +
                "this insurance can help with, or even take over the cost.");
            PrintStatement(styleDocument, "Long Term Care Insurance covers what your provincial health plan does not. This coverage ensures that the cost of long - term care does not impact " +
                "your savings and retirement income as well as become a financial strain on your loved ones. Long Term Care Insurance can cover expenses such as:");
            string[] bullets = { 
                "care by a certified nurse",
                "rehabilitation and therapy" ,
                "personal care and home care services (assistance with daily activities such as",
                "dressing, cooking, cleaning",
                "supervision by another individual"
            };
            PrintBullets(bullets);
            PrintStatement(styleDocument, "A place to look further: https://canada-life-insurance.org/Canada/long-term-careinsurance.php");
            PrintStatement(styleHeading2, "End of Life Insurance");
            PrintStatement(styleDocument, "End - of - life insurance is likely something to buy in the intermediate future. It is a way to lock in the cost of services including " +
                "funeral and graveside services in today’s prices covered by a single, or a small number of premiums. This way the only expenses and decisions needed to be made by " +
                "those we leave behind are far simpler. The policy maps out your desire for internment and prepays it. Your trustee / executor simply makes a phone call, and the " + 
                "mortician takes over.");
            PrintStatement(styleDocument, "A place to look further: https://canada-life-insurance.org/Canada/funeral-insurance.php");
            document.Content.Paragraphs.Add(ref missing).Range.InsertBreak(WdBreakType.wdPageBreak);
        }
    }
}
