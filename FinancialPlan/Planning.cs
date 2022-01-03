using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excel.FinancialFunctions;
using System.Globalization;

namespace FinancialPlan
{
    public partial class FinPlan : Form
    {
        public int rifAge = 72;
        public static DataTable LIFRates = new DataTable
        {
            Columns = { { "Age", typeof(int) }, { "Minimum", typeof(double) }, { "Prov1", typeof(double) }, { "Prov2", typeof(double) }, { "Federal", typeof(double) } },
            Rows = {
                { 50, 0.0250 , 0.0627 , 0.0610 , 0.0413 },
                { 51, 0.0256 , 0.0631 , 0.0610 , 0.0416 },
                { 52, 0.0263 , 0.0635 , 0.0610 , 0.0420 },
                { 53, 0.0270 , 0.0640 , 0.0610 , 0.0424 },
                { 54, 0.0278 , 0.0645 , 0.0610 , 0.0428 },
                { 55, 0.0286 , 0.0651 , 0.0640 , 0.0433 },
                { 56, 0.0294 , 0.0657 , 0.0650 , 0.0438 },
                { 57, 0.0303 , 0.0663 , 0.0650 , 0.0443 },
                { 58, 0.0313 , 0.0670 , 0.0660 , 0.0449 },
                { 59, 0.0323 , 0.0677 , 0.0670 , 0.0455 },
                { 60, 0.0333 , 0.0685 , 0.0670 , 0.0462 },
                { 61, 0.0345 , 0.0694 , 0.0680 , 0.0470 },
                { 62, 0.0357 , 0.0704 , 0.0690 , 0.0478 },
                { 63, 0.0370 , 0.0714 , 0.0700 , 0.0487 },
                { 64, 0.0385 , 0.0726 , 0.0710 , 0.0498 },
                { 65, 0.0400 , 0.0738 , 0.0720 , 0.0509 },
                { 66, 0.0417 , 0.0752 , 0.0730 , 0.0521 },
                { 67, 0.0435 , 0.0767 , 0.0740 , 0.0535 },
                { 68, 0.0455 , 0.0783 , 0.0760 , 0.0551 },
                { 69, 0.0476 , 0.0802 , 0.0770 , 0.0568 },
                { 70, 0.0500 , 0.0822 , 0.0790 , 0.0588 },
                { 71, 0.0528 , 0.0845 , 0.0810 , 0.0610 },
                { 72, 0.0540 , 0.0871 , 0.0830 , 0.0636 },
                { 73, 0.0553 , 0.0900 , 0.0850 , 0.0666 },
                { 74, 0.0567 , 0.0934 , 0.0880 , 0.0701 },
                { 75, 0.0582 , 0.0971 , 0.0910 , 0.0742 },
                { 76, 0.0598 , 0.1015 , 0.0940 , 0.0789 },
                { 77, 0.0617 , 0.1066 , 0.0980 , 0.0843 },
                { 78, 0.0636 , 0.1125 , 0.1030 , 0.0907 },
                { 79, 0.0658 , 0.1196 , 0.1080 , 0.0982 },
                { 80, 0.0682 , 0.1282 , 0.1150 , 0.1072 },
                { 81, 0.0708 , 0.1387 , 0.1210 , 0.1182 },
                { 82, 0.0738 , 0.1519 , 0.1290 , 0.1319 },
                { 83, 0.0771 , 0.1690 , 0.1380 , 0.1496 },
                { 84, 0.0808 , 0.1919 , 0.1480 , 0.1732 },
                { 85, 0.0851 , 0.2240 , 0.1600 , 0.2063 },
                { 86, 0.0899 , 0.2723 , 0.1730 , 0.2559 },
                { 87, 0.0955 , 0.3529 , 0.1890 , 0.3385 },
                { 88, 0.1021 , 0.5146 , 0.2000 , 0.5039 },
                { 89, 0.1099 , 1.0000 , 0.2000 , 1.0000 },
                { 90, 0.1192 , 1.0000 , 0.2000 , 1.0000 }
            }
        };

        public static DataTable RIFMinimums = new DataTable
        {
            Columns = { { "Age", typeof(int) }, { "Minimum", typeof(double) } },
            Rows = {
                { 65, 0.0400 },
                { 66, 0.0417 },
                { 67, 0.0435 },
                { 68, 0.0455 },
                { 69, 0.0476 },
                { 70, 0.0500 },
                { 71, 0.0528 },
                { 72, 0.0540 },
                { 73, 0.0553 },
                { 74, 0.0567 },
                { 75, 0.0582 },
                { 76, 0.0598 },
                { 77, 0.0617 },
                { 78, 0.0636 },
                { 79, 0.0658 },
                { 80, 0.0682 },
                { 81, 0.0708 },
                { 82, 0.0738 },
                { 83, 0.0771 },
                { 84, 0.0808 },
                { 85, 0.0851 },
                { 86, 0.0899 },
                { 87, 0.0955 },
                { 88, 0.1021 },
                { 89, 0.1099 },
                { 90, 0.1192 },
                { 91, 0.1306 },
                { 92, 0.1449 },
                { 93, 0.1634 },
                { 94, 0.1879 },
                { 95, 0.2000 }
            }
        };

        string fName = @"C:\Temp\FinancialPlan.docx";
        public static TextBox clientName;
        public static double respBalance = 0;
        public static double funeralExpenses = 20000;
        public static double costOfLiving = 0.012, returnOnInvestment = 0.09;
        public static bool preRetirement = true, retirementThruRIFAge = true, postRIFAge = true;
        public static bool preUniversityAgeChildren = false, divorced = true;
        public static double spousalPaymentAmount = 0, childCarePaymentAmount = 0;
        public static double savingsRate = 0.1;
        public static double lico = 25920;
        public const double calgaryMillRate = 0.0075223;
        public const double maxOAS = 635.26 * 12;
        public const int minLIFAge = 50;
        //public double monthlyBudgetAmt;
        public static double cppIncome = 1151.51 + 663.86;
        public static double oasIncome = 614.14 + 614.14;
        DataTable dtTable = new DataTable();
        readonly CultureInfo culture = CultureInfo.CreateSpecificCulture("en-CA");
        bool unlock = true;
        double rspAddLif = 0;
        int rspAddlifAge = 0;
        double tfsaContributions, rspContributions;
        double tfsaPerYear = 5500;
        double maxtfsaPerYear = 5500;
        double firstYearCPP = (1203.75 * 12);
        public enum Milestones { FIRSTHOME, MARRIAGE, CHILDREN, SETBACKS, RETIREMENT };
        public static DataTable AgeROI = new DataTable
        {
            Columns = { { "ROI", typeof(double) }, { "Equity", typeof(double) }, { "Fixed Income", typeof(double) } },
            Rows = { { 0.0818, 0, 100 }, { 0.0857, 10, 90 }, { 0.0894, 20, 80 }, { 0.0928, 30, 70 }, { 0.0960, 40, 60 }, { 0.0988, 50, 50 },
                    { 0.1015, 60, 40 }, { 0.1038, 70, 30 }, { 0.1058, 80, 20 }, { 0.1075, 90, 10 }, { 0.1090, 100, 0 } }
        };

        public FinPlan()
        {
            InitializeComponent();
            roiTextBox.Text = returnOnInvestment.ToString("P02", culture);
            clientName = clientNameTextBox;
        }

        private void Married_Click(object sender, EventArgs e)
        {
            if (marriedCheckBox.Checked)
            {
                spouseSalaryLabel.Enabled = true;
                spouseIncomeTextBox.Enabled = true;
                spouseRSPBalanceLabel.Enabled = true;
                spousalRSPBalanceTextBox.Enabled = true;
                spousalRSPRoomLabel.Enabled = true;
                spousalRSPRoomTextBox.Enabled = true;
                spouseTFSABalanceLabel.Enabled = true;
                spousalTFSABalanceTextBox.Enabled = true;
                spouseTFSARoomLabel.Enabled = true;
                spousalTFSARoomTextbox.Enabled = true;
                spouseCashBalanceLabel.Enabled = true;
                spouseCashBalanceTextBox.Enabled = true;
                jointCashBalanceLabel.Enabled = true;
                jointCashBalanceTextBox.Enabled = true;
            }
            else
            {
                spouseSalaryLabel.Enabled = false;
                spouseIncomeTextBox.Enabled = false;
                spouseIncomeTextBox.Text = "0";
                spouseRSPBalanceLabel.Enabled = false;
                spouseRSPBalanceLabel.Text = "0";
                spousalRSPBalanceTextBox.Enabled = false;
                spousalRSPBalanceTextBox.Text = "0";
                spousalRSPRoomLabel.Enabled = false;
                spousalRSPRoomTextBox.Enabled = false;
                spousalRSPRoomTextBox.Text = "0";
                spouseTFSABalanceLabel.Enabled = false;
                spousalTFSABalanceTextBox.Enabled = false;
                spousalTFSABalanceTextBox.Text = "0";
                spouseTFSARoomLabel.Enabled = false;
                spousalTFSARoomTextbox.Enabled = false;
                spousalTFSARoomTextbox.Text = "0";
                spouseCashBalanceLabel.Enabled = false;
                spouseCashBalanceTextBox.Enabled = false;
                spouseCashBalanceTextBox.Text = "0";
                jointCashBalanceLabel.Enabled = false;
                jointCashBalanceTextBox.Enabled = false;
                jointCashBalanceTextBox.Text = "0";
            }
        }

        private void ROIAgeBasedCheck_Changed(object sender, EventArgs e)
        {
            if (roiCheckBox.Checked)
            {
                roiTextBox.Clear();
                roiLabel.Enabled = false;
                roiTextBox.Enabled = false;
            }
            else
            {
                roiLabel.Enabled = true;
                roiTextBox.Text = returnOnInvestment.ToString("F04", culture);
                roiTextBox.Enabled = true;
            }
        }

        private void HomeOwned_Click(object sender, EventArgs e)
        {
            if (HomeOwnedCheckBox.Checked)
            {
                primaryMortgageAmtLabel.Enabled = true;
                primaryMortgageAmtTextBox.Enabled = true;
                primaryMortgageRateLabel.Enabled = true;
                primaryMortgageRateTextBox.Enabled = true;
                primaryMortgageTermLabel.Enabled = true;
                primaryMortgageTermTextBox.Enabled = true;
                primaryMortgageAmortLabel.Enabled = true;
                primaryMortgageAmortTextBox.Enabled = true;
                primaryHomeValueLabel.Enabled = true;
                primaryHomeValueTextBox.Enabled = true;
                primaryLocCheckBox.Enabled = true;
                primaryHOALabel.Enabled = true;
                primaryHoaFeesTextBox.Enabled = true;
                rentLabel.Enabled = false;
                rentTextBox.Enabled = false;
            }
            else
            {
                primaryMortgageAmtLabel.Enabled = false;
                primaryMortgageAmtTextBox.Enabled = false;
                primaryMortgageRateLabel.Enabled = false;
                primaryMortgageRateTextBox.Enabled = false;
                primaryMortgageTermLabel.Enabled = false;
                primaryMortgageTermTextBox.Enabled = false;
                primaryMortgageAmortLabel.Enabled = false;
                primaryMortgageAmortTextBox.Enabled = false;
                primaryHomeValueLabel.Enabled = false;
                primaryHomeValueTextBox.Enabled = false;
                primaryLocCheckBox.Enabled = false;
                primaryHOALabel.Enabled = false;
                primaryHoaFeesTextBox.Enabled = false;
                rentLabel.Enabled = true;
                rentTextBox.Enabled = true;
            }
        }

        private void InvestmentProperty_Click(object sender, EventArgs e)
        {
            if (investPropertyCheckBox.Checked)
            {
                investmentRentLabel.Enabled = true;
                rentalIncomeTextBox.Enabled = true;
                investPropertyHOAlabel.Enabled = true;
                investPropertyHOATextBox.Enabled = true;
                investMortgageAmountLabel.Enabled = true;
                investMortgageAmountTextbox.Enabled = true;
                investMortgageRateLabel.Enabled = true;
                investMortgageRateTextbox.Enabled = true;
                investMortgageTermLabel.Enabled = true;
                investMortgageTermTextBox.Enabled = true;
                investMortgageAmortLabel.Enabled = true;
                investMortgageAmortTextBox.Enabled = true;
                investPropertyValueLabel.Enabled = true;
                investPropertyValueTextBox.Enabled = true;
                investPropertyDatePickerLabel.Enabled = true;
                investPropertyDateTimePicker.Enabled = true;
                investPropertyLOCLabel.Enabled = true;
                investPropertyLOCTextBox.Enabled = true;
                investPropertyRateLabel.Enabled = true;
                investPropertyRateTextBox.Enabled = true;
                investPropertyHELOCCheckBox.Enabled = true;
            }
            else
            {
                investmentRentLabel.Enabled = false;
                rentalIncomeTextBox.Enabled = false;
                investPropertyHOAlabel.Enabled = false;
                investPropertyHOATextBox.Enabled = false;
                investMortgageAmountLabel.Enabled = false;
                investMortgageAmountTextbox.Enabled = false;
                investMortgageRateLabel.Enabled = false;
                investMortgageRateTextbox.Enabled = false;
                investMortgageTermLabel.Enabled = false;
                investMortgageTermTextBox.Enabled = false;
                investMortgageAmortLabel.Enabled = false;
                investMortgageAmortTextBox.Enabled = false;
                investPropertyValueLabel.Enabled = false;
                investPropertyValueTextBox.Enabled = false;
                investPropertyDatePickerLabel.Enabled = false;
                investPropertyDateTimePicker.Enabled = false;
                investPropertyLOCLabel.Enabled = false;
                investPropertyLOCTextBox.Enabled = false;
                investPropertyRateLabel.Enabled = false;
                investPropertyRateTextBox.Enabled = false;
                investPropertyHELOCCheckBox.Enabled = false;
            }
        }

        private void ClearTables()
        {
            mortgageTable.Rows.Clear();
            LIRALIFTable.Rows.Clear();
            oPension.Rows.Clear();
            cPension.Rows.Clear();
            oasTable.Rows.Clear();
            budgetdt.Rows.Clear();
            BasicPensions.Rows.Clear();
            rspTable.Rows.Clear();
            tfsaTable.Rows.Clear();
            nonRegTable.Rows.Clear();
        }

        private void GeneratePlanButton_Click(object sender, EventArgs e)
        {
            if (MyDouble(spousalTFSARoomTextbox.Text) > 0)
                tfsaPerYear *= 2;
            ClearTables();
            GenerateMortgageTable();
            GenerateLIRALIFTable();
            GenerateNonRegTable();
            GenerateOtherPensionTables();
            GenerateBasicPensionsTable();
            GenerateFirstPassBudget();               // Need LIF & Pensions Table for income in Budget
            GeneratePreRetirementTFSATable();
            GeneratePreRetirementRSPTable();

            if (retirementThruRIFAge && MyInt(currentAgeTextBox.Text) < rifAge)
            {
                GenerateRetirementThruRIFAgeRSPTable();
                GenerateRetirementThruRIFAgeTFSATable();
            }

            if (postRIFAge)
            {
                if (!GeneratePostRIFAgeRSPTable())
                {
                    tfsaPerYear = 0;
                    GenerateRetirementThruRIFAgeRSPTable();
                    GenerateRetirementThruRIFAgeTFSATable();
                    GeneratePostRIFAgeRSPTable();
                }
                GeneratePostRIFAgeTFSATable();
            }
            DisplayDGV(dgvStage1, LIRALIFTable);
            DisplayDGV(dgvStage2, rspTable);
            DisplayDGV(dgvStage3, tfsaTable);

            CreateDocument();
            LifeInsurance();
        }

        
        private void DisplayDGV(DataGridView dgv, DataTable dt)
        {
            dgv.DataSource = dt.DefaultView;
            dgv.RowsDefaultCellStyle.BackColor = Color.White;
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;
            dgv.ReadOnly = true;
            dgv.MultiSelect = false;
            dgv.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            for (int i = 0; i < dgv.Columns.Count; i++)
                dgv.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv.Columns[0].DefaultCellStyle.Format = "c";
            dgv.Columns[3].DefaultCellStyle.Format = "c";
            dgv.Columns[4].DefaultCellStyle.Format = "c";

            if (dgv.Columns.Count > 6)
            {
                dgv.Columns[5].DefaultCellStyle.Format = "##.##%";
                dgv.Columns[6].DefaultCellStyle.Format = "##.##%";
            }
            else
            {
                dgv.Columns[5].DefaultCellStyle.Format = "c";
            }
        }

        private void GenerateRetirementThruRIFAgeTFSATable()
        {
            int age = MyInt(retirementAgeTextBox.Text) + 1;
            int currentAge = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int year = startDate.Year + retirementAge - currentAge + 1;
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double tfsaGrowth, contribution = tfsaPerYear;
            double balance;
            double tfsaRoom, wthAmt;
            bool exists = false;

            if (currentAge >= retirementAge)
            {
                tfsaRoom = MyDouble(tfsaRoomTextBox.Text) + MyDouble(spousalTFSABalanceTextBox.Text) + tfsaPerYear;
                balance = MyDouble(tfsaBalanceTextBox.Text) + MyDouble(spousalTFSABalanceTextBox.Text);
            }
            else
            {
                tfsaRoom = MyDouble(tfsaTable.Rows[age - currentAge - 1]["Limit"].ToString()) + tfsaPerYear;
                balance = MyDouble(tfsaTable.Rows[age - currentAge - 1]["Balance"].ToString());
            }

            for (int i = retirementAge; i < rifAge && i < deathAge; i++)
            {
                DataRow r;

                if (contribution > tfsaRoom)
                    contribution = tfsaRoom;
                wthAmt = -contribution;
                if (tfsaTable.Rows.Count > i - currentAge)
                {
                    exists = true;
                    r = tfsaTable.Rows[i - currentAge];
                }
                else
                    r = tfsaTable.NewRow();

                r["Balance"] = balance;
                r["Limit"] = tfsaRoom - contribution;
                tfsaRoom -= contribution;
                r["Withdrawals"] = wthAmt;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = tfsaGrowth = balance * growthRate;
                balance += (tfsaGrowth - wthAmt);
                if (balance < 0)
                    balance = 0;
                if (!exists)
                    tfsaTable.Rows.Add(r);
                exists = false;
                tfsaRoom += contribution;
            }
        }

        private void GenerateOtherPensionTables()
        {
            double cPPension = 0, oasAmt = 0, otherPension = otherPensions;
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int currentAge = MyInt(currentAgeTextBox.Text);
            int year = startDate.Year;
            int deathAge = MyInt(deathAgeTextBox.Text);

            double monthlyCPP = firstYearCPP / 12;
            double monthlyCPPPenalty = monthlyCPP * cpptiming * 0.006;
            monthlyCPP -= monthlyCPPPenalty;
            firstYearCPP = monthlyCPP * 12;

            for (int i = 0; i <= deathAge - currentAge; i++)
            {

                if (currentAge + i >= retirementAge)
                {
                    if (cPPension == 0)
                        cPPension = firstYearCPP;
                    if (oasAmt == 0)
                        oasAmt = maxOAS;
                }
                DataRow oPP = oPension.NewRow();
                DataRow cPP = cPension.NewRow();
                DataRow oas = oasTable.NewRow();
                oPP["Year"] = cPP["Year"] = oas["Year"] = year + i;
                oPP["Age"] = cPP["Age"] = oas["Age"] = currentAge + i;
                oPP["Withdrawals"] = otherPension;
                cPP["WithDrawals"] = cPPension;
                oas["Amount"] = oasAmt;
                oPension.Rows.Add(oPP);
                cPension.Rows.Add(cPP);
                oasTable.Rows.Add(oas);
                otherPension = otherPension * (1 + costOfLiving);
                cPPension = cPPension * (1 + costOfLiving);
            }
        }

        private void GenerateRetirementThruRIFAgeRSPTable()
        {
            // This is where we take the budget amount + tfsa deposits and try to meet it with LIF and RSP withdrawals
            int age = MyInt(retirementAgeTextBox.Text) + 1;
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int currentAge = MyInt(currentAgeTextBox.Text);
            int year = startDate.Year + retirementAge - currentAge + 1;
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double rspGrowth;
            double balance;
            double wthAmt;//, cPPension, otherPension;
            double rspRoom;
            bool exists = false;

            /* Need work here. cPension start year and amount should be read from FinancialData
             * Also, withdraw funding for TFSA from age retirementAget
             */

            if (currentAge >= retirementAge)
            {
                rspRoom = MyDouble(rspRoomTextBox.Text) + MyDouble(spousalRSPRoomTextBox.Text);
                balance = MyDouble(rspBalanceTextBox.Text) + MyDouble(spousalRSPBalanceTextBox.Text);
            }
            else
            {
                rspRoom = MyDouble(rspTable.Rows[age - currentAge - 1]["Limit"].ToString());
                balance = MyDouble(rspTable.Rows[age - currentAge - 1]["Balance"].ToString());
            }

            for (int i = retirementAge; i < rifAge && i < deathAge; i++)
            {
                double expenses = MyDouble(budgetdt.Rows[i - currentAge]["Expenses"].ToString()) + tfsaPerYear;
                double incomes  = MyDouble(budgetdt.Rows[i - currentAge]["Incomes"].ToString());
                wthAmt = Math.Max(0, expenses - incomes);
                if (wthAmt > balance)
                {
                    wthAmt -= tfsaPerYear;
                    tfsaPerYear = 0;
                    if (wthAmt > balance)
                    {
                        statusLabelValue.Text = "Invalid: Not enough to retire";
                        wthAmt = balance;
                    }
                }

                DataRow r;
                if (rspTable.Rows.Count > i - currentAge)
                {
                    exists = true;
                    r = rspTable.Rows[i - currentAge];
                }
                else
                    r = rspTable.NewRow();
                
                r["Withdrawals"] = wthAmt;
                budgetdt.Rows[i - currentAge]["RSP/RIF"] = wthAmt;
                budgetdt.Rows[i - currentAge]["Incomes"] = incomes + wthAmt;
                r["Limit"] = rspRoom;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = rspGrowth = balance * growthRate;
                balance += (rspGrowth - wthAmt);
                r["Balance"] = balance;
                if (!exists)
                    rspTable.Rows.Add(r);
                exists = false;
            }
        }

        private void GeneratePostRIFAgeTFSATable()
        {
            int age = rifAge;
            int currentAge = MyInt(currentAgeTextBox.Text);
            int year = startDate.Year + rifAge - currentAge + 1;
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double tfsaGrowth;
            double balance = MyDouble(tfsaTable.Rows[age - currentAge - 1]["Balance"].ToString());
            double contribution = tfsaPerYear;
            double tfsaRoom = MyDouble(tfsaTable.Rows[tfsaTable.Rows.Count - 1]["Limit"].ToString()) + tfsaPerYear;
            double rspBalance = MyDouble(rspTable.Rows[age - currentAge - 1]["Balance"].ToString());

            for (int i = rifAge; i <= deathAge; i++)
            {
                
                if (contribution > tfsaRoom)
                    contribution = tfsaRoom;
                
                DataRow r = tfsaTable.NewRow();
                double expenses = MyDouble(budgetdt.Rows[i - currentAge]["Expenses"].ToString()) + tfsaPerYear;
                double incomes = MyDouble(budgetdt.Rows[i - currentAge]["Incomes"].ToString());

                double shortFall = Math.Max(0, incomes - expenses);
                r["Withdrawals"] = shortFall - contribution;
                budgetdt.Rows[i - currentAge]["TFSA"] = shortFall - contribution;
                budgetdt.Rows[i - currentAge]["Incomes"] = incomes + shortFall - contribution;
                r["Limit"] = tfsaRoom - contribution;
                tfsaRoom -= contribution;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = tfsaGrowth = balance * growthRate;
                balance += (tfsaGrowth - shortFall + contribution);
                if (balance < 0)
                    balance = 0;
                r["Balance"] = balance;
                tfsaTable.Rows.Add(r);
                tfsaRoom += maxtfsaPerYear;
            }
        }

        /* At this stage the RSP is converted to a RIF and we're bound to take out at least the minimum
         */

        private bool GeneratePostRIFAgeRSPTable()
        {
            int age = rifAge;
            int currentAge = MyInt(currentAgeTextBox.Text);
            int year = startDate.Year + rifAge - currentAge + 1;
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double rspGrowth;
            double balance = MyDouble(rspTable.Rows[age - currentAge - 1]["Balance"].ToString());
            double wthAmt;
            double rspRoom = 0; // MyDouble(rspTable.Rows[age - currentAge - 1]["Limit"].ToString());

            /* Need work here. cPension start year and amount should be read from FinancialData
             * Also, withdraw funding for TFSA from age retirementAget
             */

            for (int i = rifAge; i <= deathAge; i++)
            {
                double expenses = MyDouble(budgetdt.Rows[i - currentAge]["Expenses"].ToString()) + tfsaPerYear;
                double incomes = MyDouble(budgetdt.Rows[i - currentAge]["Incomes"].ToString());
                wthAmt = Math.Max(0, incomes - expenses);
                double rifMinimum = MyDouble(RIFMinimums.Rows[i - 70 + 5]["Minimum"].ToString()) * balance;
                if (wthAmt > balance)
                {
                    wthAmt -= tfsaPerYear;
                    if (wthAmt > balance)
                    {
                        // Need to start over with no TFSA contributions after retirement
                        if (tfsaPerYear > 0)
                            statusLabelValue.Text = "Not enough to Contribute";
                        else
                            statusLabelValue.Text = "Not enough to retire";
                        wthAmt = balance;
                        if (tfsaPerYear != 0)
                        {
                            tfsaPerYear = 0;
                            return (false);
                        }
                    }
                }
                else if (wthAmt < rifMinimum)
                    wthAmt = rifMinimum;

                DataRow r = rspTable.NewRow();
                
                r["Withdrawals"] = wthAmt;
                budgetdt.Rows[i - currentAge]["RSP/RIF"] = wthAmt;
                budgetdt.Rows[i - currentAge]["Incomes"] = incomes + wthAmt;
                r["Limit"] = rspRoom;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = rspGrowth = balance * growthRate;
                balance += (rspGrowth - wthAmt);
                r["Balance"] = balance;
                rspTable.Rows.Add(r);
            }
            return (true);
        }

        private void GeneratePreRetirementRSPTable()
        {
            double rspRoom = MyDouble(rspRoomTextBox.Text) + MyDouble(spousalRSPRoomTextBox.Text);
            double balance = MyDouble(rspBalanceTextBox.Text) + MyDouble(spousalRSPBalanceTextBox.Text);
            double contribution = rspContributions;
            int year = startDate.Year;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double rspGrowth, wthAmt = 0;
            double annualIncome = MyDouble(incomeTextBox.Text);

            for (int i = age; i <= retirementAge; i++)
            {
                if (rspAddlifAge == i)
                    balance += rspAddLif;
                DataRow r = rspTable.NewRow();
                if (contribution > rspRoom)
                    contribution = rspRoom;
                wthAmt = -contribution;
                r["Balance"] = balance;
                r["Limit"] = rspRoom - contribution;
                rspRoom -= contribution;
                r["Withdrawals"] = -contribution;
                //double incomes = MyDouble(budgetdt.Rows[i - retirementAge]["Incomes"].ToString());
                //budgetdt.Rows[i - retirementAge]["RSP/RIF"] = wthAmt;
                //budgetdt.Rows[i - retirementAge]["Incomes"] = incomes + wthAmt;
                double incomes = MyDouble(budgetdt.Rows[i - age]["Incomes"].ToString());
                budgetdt.Rows[i - age]["RSP/RIF"] = wthAmt;
                budgetdt.Rows[i - age]["Incomes"] = incomes + wthAmt;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = rspGrowth = balance * growthRate;
                balance += (rspGrowth - wthAmt);
                rspTable.Rows.Add(r);
                rspRoom += annualIncome * 0.18;
            }
        }

        private void GeneratePreRetirementTFSATable()
        {
            double tfsaRoom = MyDouble(tfsaRoomTextBox.Text) + MyDouble(spousalTFSARoomTextbox.Text);
            double balance = MyDouble(tfsaBalanceTextBox.Text) + MyDouble(spousalTFSABalanceTextBox.Text);
            double contribution = tfsaContributions;
            int year = startDate.Year;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double tfsaGrowth, wthAmt = 0;

            for (int i = age; i <= retirementAge; i++)
            {
                if (contribution > tfsaRoom)
                    contribution = tfsaRoom;
                wthAmt = -contribution;
                DataRow r = tfsaTable.NewRow();
                r["Withdrawals"] = wthAmt;
                //double incomes = MyDouble(budgetdt.Rows[i - retirementAge]["Incomes"].ToString());
                double incomes = MyDouble(budgetdt.Rows[i - age]["Incomes"].ToString());
                budgetdt.Rows[i - age]["TFSA"] = wthAmt;
                //budgetdt.Rows[i - retirementAge]["Incomes"] = incomes + wthAmt;
                budgetdt.Rows[i - age]["Incomes"] = incomes + wthAmt;
                r["Balance"] = balance;
                r["Limit"] = tfsaRoom - contribution;
                tfsaRoom -= contribution;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Growth"] = tfsaGrowth = balance * growthRate;
                balance += (tfsaGrowth - wthAmt);
                if (balance < 0)
                    balance = 0;
                tfsaTable.Rows.Add(r);
                tfsaRoom += maxtfsaPerYear;
            }
        }

        public DataTable mortgageTable = new DataTable
        {
            Columns = { { "Age", typeof(Int16) }, { "Year", typeof(Int16) }, { "Payment", typeof(double) }, { "Interest", typeof(double) }, { "Balance", typeof(double) } }
        };

        // Assume 25 year total amort
        private void GenerateMortgageTable()
        {
            DataTable dt = new DataTable
            {
                Columns = { { "Age", typeof(Int16) }, { "Year", typeof(Int16) }, { "Num", typeof(Int16) }, { "Date", typeof(DateTime) }, { "Payment", typeof(double) },
                    { "Interest", typeof(double) }, { "Principle", typeof(double) }, {"Balance", typeof(double) } }
            };
            int pmtsPerYear = MyInt(primaryMortgageTermTextBox.Text);
            int pmtsLeft = MyInt(primaryMortgageAmortTextBox.Text);
            int firstPaymentNumber = (25 * pmtsPerYear) - pmtsLeft;
            DateTime t = startDate;
            double daysBetweenPmts = Math.Round(365D / pmtsPerYear);
            double balance = MyDouble(primaryMortgageAmtTextBox.Text);
            double rate = MyRate(primaryMortgageRateTextBox.Text);

            if (pmtsLeft == 0)
            {
                System.Data.DataRow r = mortgageTable.NewRow();
                r["Age"] = t.Year - dob.Year;
                r["Year"] = t.Year;
                r["Payment"] = 0;
                r["Interest"] = 0;
                r["Balance"] = 0;
                mortgageTable.Rows.Add(r);
                return;
            }
            double pv = balance;

            for (int i = 0; i < pmtsLeft; i++)
            {
                DataRow r;

                r = dt.NewRow();
                r["Num"] =  i;
                r["Date"] = t;
                r["Year"] = t.Year;
                r["Payment"] = primaryMortgagePmt;
                if (t.Month <= dob.Month) {
                    if (t.Month == dob.Month && t.Day <= dob.Day)
                        r["Age"] = t.Year - dob.Year - 1;
                    else if (t.Month == dob.Month && t.Day > dob.Day)
                        r["Age"] = t.Year - dob.Year;
                    else
                        r["Age"] = t.Year - dob.Year - 1;
                }
                else
                    r["Age"] = t.Year - dob.Year;
                t = t.AddDays(daysBetweenPmts);
                double ppmt = -Financial.PPmt(rate / pmtsPerYear, i + 1, pmtsLeft, pv, 0, PaymentDue.EndOfPeriod);
                balance -= ppmt;
                r["Balance"] = balance;
                r["Principle"] = ppmt;
                r["Interest"] = primaryMortgagePmt - ppmt;
                dt.Rows.Add(r);
            }

            DateTime end = t;
            t = startDate;
            for (int i = 0; i < end.Year - startDate.Year + 1; i++)
            {
                DataRow r;

                r = mortgageTable.NewRow();
                r["Age"] = t.Year - dob.Year;
                r["Year"] = t.Year;
                r["Payment"] = dt.Compute("SUM(Payment)", "Year = " + t.Year);
                r["Interest"] = dt.Compute("SUM(Interest)", "Year = " + t.Year);
                r["Balance"] = dt.Compute("Min(Balance)", "Year = " + t.Year);
                t = t.AddYears(1);
                mortgageTable.Rows.Add(r);
            }
        }

        public DataTable LIRALIFTable = new DataTable
        {
            Columns = { { "Balance", typeof(double) }, { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) }, { "Growth", typeof(double) }, { "Min Withdrawal", typeof(double) }, { "Max Withdrawal", typeof(double) } },
        };

        public DataTable rspTable = new DataTable
        {
            Columns = { { "Balance", typeof(double) }, { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) }, { "Growth", typeof(double) }, { "Limit", typeof(double) } }
        };

        public DataTable tfsaTable = new DataTable
        {
            Columns = { { "Balance", typeof(double) }, { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) }, { "Growth", typeof(double) }, { "Limit", typeof(double) } }
        };

        public DataTable nonRegTable = new DataTable
        {
            Columns = { { "Balance", typeof(double) }, { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) }, { "Growth", typeof(double) } }
        };

        public DataTable oPension = new DataTable
        {
            Columns = { { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) } }
        };

        public DataTable cPension = new DataTable
        {
            Columns = { { "Year", typeof(int) }, { "Age", typeof(int) }, { "Withdrawals", typeof(double) } }
        };

        public DataTable oasTable = new DataTable
        {
            Columns = { { "Year", typeof(int) }, { "Age", typeof(int) }, { "Amount", typeof(double) } }
        };

        private void GenerateNonRegTable()
        {
            double balance, nonRegGrowth;
            int year = startDate.Year;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);

            balance = MyDouble(cashBalanceTextBox.Text) + MyDouble(spouseCashBalanceTextBox.Text) + MyDouble(jointCashBalanceTextBox.Text);

            for (int i = age; i <= deathAge; i++)
            {
                DataRow r = nonRegTable.NewRow();
                r["Balance"] = balance;
                r["Year"] = year++;
                r["Age"] = age++;
                r["Withdrawals"] = 0;
                r["Growth"] = nonRegGrowth = balance * growthRate;
                balance += nonRegGrowth;
                nonRegTable.Rows.Add(r);
            }
        }

        private void GenerateLIRALIFTable()
        {
            double balance, LIFminRate, LIFmaxRate, lifGrowth;
            int year = startDate.Year;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            double growthRate = MyDouble(roiTextBox.Text);
            double wthAmt = 0;

            if (liraLIFAmt > 0)
                wthAmt = -Financial.Pmt(growthRate, deathAge - age, liraLIFAmt, 0, PaymentDue.BeginningOfPeriod);

            bool unlocked = false;
            if (unlock && age >= minLIFAge)
            {
                double rsp = MyDouble(rspBalanceTextBox.Text);
                rsp += liraLIFAmt / 2;
                rspBalanceTextBox.Text = rsp.ToString();
                liraLIFAmt /= 2;
                unlocked = true;
            }
            balance = liraLIFAmt + spouseLiraLIFAmt;

            for (int i = age; i <= deathAge ; i++) {
                DataRow r = LIRALIFTable.NewRow();
                r["Balance"] = balance;
                r["Year"] = year++;
                //r["Age"] = age++;
                //if (age >= minLIFAge)
                r["Age"] = i;
                if (i >= minLIFAge)
                {
                    //r["Min Withdrawal"] = LIFminRate = MyDouble(LIFRates.Rows[age - minLIFAge]["Minimum"].ToString());
                    //r["Max Withdrawal"] = LIFmaxRate = MyDouble(LIFRates.Rows[age - minLIFAge][LIRALIFLegislation].ToString());
                    r["Min Withdrawal"] = LIFminRate = MyDouble(LIFRates.Rows[i - minLIFAge]["Minimum"].ToString());
                    r["Max Withdrawal"] = LIFmaxRate = MyDouble(LIFRates.Rows[i - minLIFAge][LIRALIFLegislation].ToString());
                    if (!unlocked && unlock)
                    {
                        //double rsp = MyDouble(rspBalanceTextBox.Text) + balance / 2;
                        //rspBalanceTextBox.Text = rsp.ToString();
                        rspAddLif = balance / 2;
                        rspAddlifAge = age;
                        balance /= 2;
                        unlocked = true;
                    }
                    double minwthdrwl = balance * LIFminRate;
                    double maxwthdrwl = balance * LIFmaxRate;

                    if (balance > 0 && i > retirementAge)
                        wthAmt = -Financial.Pmt(growthRate, deathAge - i + 1, balance, 0, PaymentDue.BeginningOfPeriod);
                    else if (balance > 0)
                        wthAmt = minwthdrwl;
                    
                    if (wthAmt < minwthdrwl)
                        wthAmt = minwthdrwl;
                    else if (wthAmt > maxwthdrwl)
                        wthAmt = maxwthdrwl;

                    if (wthAmt > balance)
                        wthAmt = balance;
                    r["WithDrawals"] = wthAmt;
                }
                r["Growth"] = lifGrowth = balance * growthRate;
                balance += (lifGrowth - wthAmt);
                LIRALIFTable.Rows.Add(r);
            }
        }

        System.Data.DataTable BasicPensions = new System.Data.DataTable
        {
            Columns = { { "Year", typeof(int) }, { "Age", typeof(int) }, { "CPP", typeof(double) }, { "OAS", typeof(double) }, { "Other", typeof(double) }, { "Total", typeof(double) } }
        };

        public void GenerateBasicPensionsTable()
        {
            int age = MyInt(currentAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            

            for (int i = age; i <= deathAge; i++)
            {
                System.Data.DataRow r = BasicPensions.NewRow();

                r["Year"] = cPension.Rows[i - age]["Year"];
                r["Age"] = cPension.Rows[i - age]["Age"];
                r["CPP"] = cPension.Rows[i - age]["Withdrawals"];
                r["OAS"] = oasTable.Rows[i - age]["Amount"];
                r["Other"] = oPension.Rows[i - age]["Withdrawals"];
                r["Total"] = MyDouble(r["CPP"].ToString()) + MyDouble(r["OAS"].ToString()) + MyDouble(r["Other"].ToString());
                BasicPensions.Rows.Add(r);
            }
        }

        DataTable budgetdt = new DataTable
        {
            Columns = { 
                { "Year", typeof(int) }, 
                { "Age", typeof(int) },
                { "Salary", typeof(double) },
                { "CPP", typeof(double) },
                { "OAS", typeof(double) },
                { "Other Pension", typeof(double) },
                { "LIRA/LIF", typeof(double) },
                { "RSP/RIF", typeof(double) },
                { "TFSA", typeof(double) },
                { "Non-Registered", typeof(double) },
                { "Incomes", typeof(double) },
                { "Food & Dining", typeof(double) },
                { "Bills & Utilities", typeof(double) },
                { "Income Tax", typeof(double) },
                { "Personal Care", typeof(double) },
                { "Leisure", typeof(double) },
                { "Auto", typeof(double) },
                { "Shopping", typeof(double) },
                { "Vacation", typeof(double) },
                { "Medical", typeof(double) },
                { "LOC Debt", typeof(double) },
                { "CC Debt", typeof(double) },
                { "Housing", typeof(double) },
                { "Expenses", typeof(double) }
            },
        };
        public void GenerateFirstPassBudget()
        {
            const int maxIncomeIndex = 10;
            int dependantChildren = MyInt(dependantChildrenTextBox.Text);
            int persons = (1 + (marriedCheckBox.Checked ? 1 : 0)) + dependantChildren;
            double salary = MyDouble(incomeTextBox.Text) + MyDouble(spouseIncomeTextBox.Text);
            int year = startDate.Year;
            int age = MyInt(currentAgeTextBox.Text);
            int retirementAge = MyInt(retirementAgeTextBox.Text);
            int deathAge = MyInt(deathAgeTextBox.Text);
            double expenses, incomes;
            double locDebt = MyDouble(lineOfCreditAmtTextBox.Text);
            double locDebtRate = MyRate(lineOfCreditRateTextBox.Text);
            double locDebtPer;
            double ccDebt = MyDouble(creditCardDebtTextBox.Text);
            double ccDebtRate = MyRate(creditCardRateTextBox.Text);
            double ccDebtPer = 2;
 
            if (locDebt < salary * 0.1)
                locDebtPer = 1;
            else if (locDebt < salary * 0.2)
                locDebtPer = 2;
            else
                locDebtPer = 5;

            if (ccDebt < salary * 0.1)
                ccDebtPer = 1;
            else if (ccDebt < salary * .2)
                ccDebtPer = 2;
            else
                ccDebtPer = 5;

            DataRow r = budgetdt.NewRow();
            r["Age"] = age;
            r["Year"] = year;
            r["Salary"] = salary;
            r["LIRA/LIF"] = r["RSP/RIF"] = r["TFSA"] = r["Non-Registered"] = 0;
            if (age <= retirementAge)
            {
                r["CPP"] = MyDouble(cPension.Rows[retirementAge - age]["Withdrawals"].ToString());
                r["OAS"] = MyDouble(oasTable.Rows[retirementAge - age]["Amount"].ToString());
                r["Other Pension"] = MyDouble(oPension.Rows[retirementAge - age]["Withdrawals"].ToString());
            }
            else
                r["CPP"] = r["OAS"] = r["Other Pension"] = 0;
            incomes = 0;
            for (int j = 2; j < maxIncomeIndex; j++)
                incomes += MyDouble(r[j].ToString());
            r["Incomes"] = incomes;
            r["Income Tax"] = CPP(salary) + EI(salary) + IncomeTax(incomes);

            r["Food & Dining"] = 1000 * 12 * (1 + (marriedCheckBox.Checked ? 0.5 : 0)) + (300 * 12 * dependantChildren);
            r["Bills & Utilities"] = (50 + 85 + 25 + (persons * 20) + 200 + 200 + 40 + 20) * 12;             // cable + internet + phone + mobile phone + gas + electric + water + garbage pickup
            r["Personal Care"] = 200 * (1 + (marriedCheckBox.Checked ? 1 : 0)) * 12 + (200 * 12 * dependantChildren / 10);
            r["Leisure"] = 300 * (1 + (marriedCheckBox.Checked ? 1 : 0)) * 12 + (300 * 12 * dependantChildren / 10);
            r["Auto"] = 6000 * (1 + (marriedCheckBox.Checked ? 1 : 0));
            r["Housing"] = HousingCosts(0);
            r["Shopping"] = 100 * (1 + (marriedCheckBox.Checked ? 1 : 0)) * 12 + (100 * 12 * dependantChildren / 10);
            r["Vacation"] = 2000 * (1 + (marriedCheckBox.Checked ? 1 : 0)) + (2000 * dependantChildren / 10);
            r["Medical"] = 100 * (1 + (marriedCheckBox.Checked ? 1 : 0)) * 12 + (100 * 12 * dependantChildren / 10);
            r["LOC Debt"] = (locDebt > 0) ? -Financial.Pmt(locDebtRate / locDebtPer, locDebtPer, locDebt, 0, PaymentDue.EndOfPeriod) : 0;
            r["CC Debt"] = (ccDebt > 0) ? -Financial.Pmt(ccDebtRate / ccDebtPer, ccDebtPer, ccDebt, 0, PaymentDue.EndOfPeriod) : 0;
            expenses = 0;
            for (int j = maxIncomeIndex+1; j < budgetdt.Columns.Count - 1; j++)
                expenses += MyDouble(r[j].ToString());
            r["Expenses"] = expenses;
            budgetdt.Rows.Add(r);

            for (int i = age + 1; i <= deathAge; i++)           // Stage 1 Budget - Preretirement
            {
                r = budgetdt.NewRow();
                r["Age"] = i;
                //r["Year"] = year + i - age;
                r["Year"] = MyInt(budgetdt.Rows[i - age - 1]["Year"].ToString()) + 1;
                r["Salary"] = MyDouble(budgetdt.Rows[i - age - 1]["Salary"].ToString()) * (1 + costOfLiving);
                r["LIRA/LIF"] = MyDouble(LIRALIFTable.Rows[i - age]["Withdrawals"].ToString());
                r["RSP/RIF"] = r["TFSA"] = r["Non-Registered"] = 0;
                r["CPP"] = r["OAS"] = r["Other Pension"] = 0;
                incomes = MyDouble(r["Salary"].ToString()) + MyDouble(r["LIRA/LIF"].ToString());
                r["Incomes"] = incomes;
                r["Income Tax"] = CPP(MyDouble(r["Salary"].ToString())) + EI(MyDouble(r["Salary"].ToString())) + IncomeTax(incomes);

                r["Food & Dining"] = MyDouble(budgetdt.Rows[i - age - 1]["Food & Dining"].ToString()) * (1 + costOfLiving);
                r["Bills & Utilities"] = MyDouble(budgetdt.Rows[i - age - 1]["Bills & Utilities"].ToString()) * (1 + costOfLiving);
                r["Personal Care"] = MyDouble(budgetdt.Rows[i - age - 1]["Personal Care"].ToString()) * (1 + costOfLiving);
                r["Leisure"] = MyDouble(budgetdt.Rows[i - age - 1]["Leisure"].ToString()) * (1 + costOfLiving);
                r["Auto"] = MyDouble(budgetdt.Rows[i - age - 1]["Auto"].ToString()) * (1 + costOfLiving);
                r["Housing"] = HousingCosts(i - age);
                r["Shopping"] = MyDouble(budgetdt.Rows[i - age - 1]["Shopping"].ToString()) * (1 + costOfLiving);
                r["Vacation"] = MyDouble(budgetdt.Rows[i - age - 1]["Vacation"].ToString()) * (1 + costOfLiving);
                r["Medical"] = MyDouble(budgetdt.Rows[i - age - 1]["Medical"].ToString()) * (1 + costOfLiving);
                r["LOC Debt"] = (--locDebtPer > 0) ? MyDouble(budgetdt.Rows[i - age - 1]["LOC Debt"].ToString()) : 0;
                r["CC Debt"] = (--ccDebtPer > 0) ? MyDouble(budgetdt.Rows[i - age - 1]["CC Debt"].ToString()) : 0;
                expenses = 0;
                for (int j = maxIncomeIndex + 1; j < budgetdt.Columns.Count - 1; j++)
                    expenses += MyDouble(r[j].ToString());
                r["Expenses"] = expenses;
                budgetdt.Rows.Add(r);
            }

            /*
            for (int i = age + 1; i <= retirementAge; i++)           // Stage 1 Budget - Preretirement
            {
                r = budgetdt.NewRow();
                r["Age"] = i;
                //r["Year"] = year + i - age;
                r["Year"] = MyInt(budgetdt.Rows[i - age - 1]["Year"].ToString()) + 1;
                r["Salary"] = MyDouble(budgetdt.Rows[i - age - 1]["Salary"].ToString()) * (1 + costOfLiving);
                r["LIRA/LIF"] = MyDouble(LIRALIFTable.Rows[i - age]["Withdrawals"].ToString());
                r["RSP/RIF"] = r["TFSA"] = r["Non-Registered"] = 0;
                r["CPP"] = r["OAS"] = r["Other Pension"] = 0;
                incomes = MyDouble(r["Salary"].ToString()) + MyDouble(r["LIRA/LIF"].ToString());
                r["Incomes"] = incomes;
                r["Income Tax"] = CPP(MyDouble(r["Salary"].ToString())) + EI(MyDouble(r["Salary"].ToString())) + IncomeTax(incomes);

                r["Food & Dining"] = MyDouble(budgetdt.Rows[i - age - 1]["Food & Dining"].ToString()) * (1 + costOfLiving);
                r["Bills & Utilities"] = MyDouble(budgetdt.Rows[i - age - 1]["Bills & Utilities"].ToString()) * (1 + costOfLiving);
                r["Personal Care"] = MyDouble(budgetdt.Rows[i - age - 1]["Personal Care"].ToString()) * (1 + costOfLiving);
                r["Leisure"] = MyDouble(budgetdt.Rows[i - age - 1]["Leisure"].ToString()) * (1 + costOfLiving);
                r["Auto"] = MyDouble(budgetdt.Rows[i - age - 1]["Auto"].ToString()) * (1 + costOfLiving);
                r["Housing"] = HousingCosts(i - age);
                r["Shopping"] = MyDouble(budgetdt.Rows[i - age - 1]["Shopping"].ToString()) * (1 + costOfLiving);
                r["Vacation"] = MyDouble(budgetdt.Rows[i - age - 1]["Vacation"].ToString()) * (1 + costOfLiving);
                r["Medical"] = MyDouble(budgetdt.Rows[i - age - 1]["Medical"].ToString()) * (1 + costOfLiving);
                r["LOC Debt"] = (--locDebtPer > 0) ? MyDouble(budgetdt.Rows[i - age - 1]["LOC Debt"].ToString()) : 0;
                r["CC Debt"] = (--ccDebtPer > 0) ? MyDouble(budgetdt.Rows[i - age - 1]["CC Debt"].ToString()) : 0;
                expenses = 0;
                for (int j = maxIncomeIndex+1; j < budgetdt.Columns.Count - 1; j++)
                    expenses += MyDouble(r[j].ToString());
                r["Expenses"] = expenses;
                budgetdt.Rows.Add(r);
            }

            int adjust = 1;
            for (int i = retirementAge + 1; i < rifAge; i++)           // Stage 2 Budget - Retired
            {
                if (i > age)
                {
                    r = budgetdt.NewRow();
                    r["Age"] = i;
                    //r["Year"] = year + i - age;
                    r["Year"] = MyInt(budgetdt.Rows[i - retirementAge - adjust]["Year"].ToString()) + 1;
                    r["Salary"] = 0;
                    r["LIRA/LIF"] = MyDouble(LIRALIFTable.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                    r["RSP/RIF"] = r["TFSA"] = r["Non-Registered"] = 0;
                    r["CPP"] = MyDouble(cPension.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                    r["OAS"] = MyDouble(oasTable.Rows[i - retirementAge - adjust + 1]["Amount"].ToString());
                    r["Other Pension"] = MyDouble(oPension.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                    incomes = 0;
                    for (int j = 2; j < maxIncomeIndex; j++)
                        incomes += MyDouble(r[j].ToString());
                    r["Incomes"] = incomes;
                    r["Income Tax"] = IncomeTax(incomes);

                    r["Food & Dining"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Food & Dining"].ToString()) * (1 + costOfLiving);
                    r["Bills & Utilities"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Bills & Utilities"].ToString()) * (1 + costOfLiving);
                    r["Personal Care"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Personal Care"].ToString()) * (1 + costOfLiving);
                    r["Leisure"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Leisure"].ToString()) * (1 + costOfLiving);
                    r["Auto"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Auto"].ToString()) * (1 + costOfLiving);
                    r["Housing"] = HousingCosts(i - retirementAge - adjust);
                    r["Shopping"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Shopping"].ToString()) * (1 + costOfLiving);
                    r["Vacation"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Vacation"].ToString()) * (1 + costOfLiving);
                    r["Medical"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Medical"].ToString()) * (1 + costOfLiving);
                    r["LOC Debt"] = (--locDebtPer > 0) ? MyDouble(budgetdt.Rows[i - retirementAge - adjust]["LOC Debt"].ToString()) : 0;
                    r["CC Debt"] = (--ccDebtPer > 0) ? MyDouble(budgetdt.Rows[i - retirementAge - adjust]["CC Debt"].ToString()) : 0;

                    expenses = 0;
                    for (int j = maxIncomeIndex + 1; j < budgetdt.Columns.Count - 1; j++)
                        expenses += MyDouble(r[j].ToString());
                    r["Expenses"] = expenses;

                    budgetdt.Rows.Add(r);
                }
                else
                    adjust++;
            }

            for (int i = rifAge; i <= deathAge; i++)                // Stage 3 Budget - Retired + CPP et. al.
            {
                r = budgetdt.NewRow();
                r["Age"] = i;
                //r["Year"] = year + i - rifAge;
                r["Year"] = MyInt(budgetdt.Rows[i - retirementAge - adjust]["Year"].ToString()) + 1;
                r["Salary"] = 0;
                r["LIRA/LIF"] = MyDouble(LIRALIFTable.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                r["RSP/RIF"] = r["TFSA"] = r["Non-Registered"] = 0;
                r["CPP"] = MyDouble(cPension.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                r["OAS"] = MyDouble(oasTable.Rows[i - retirementAge - adjust + 1]["Amount"].ToString());
                r["Other Pension"] = MyDouble(oPension.Rows[i - retirementAge - adjust + 1]["Withdrawals"].ToString());
                incomes = 0;
                for (int j = 2; j < maxIncomeIndex; j++)
                    incomes += MyDouble(r[j].ToString());
                r["Incomes"] = incomes;
                r["Income Tax"] = IncomeTax(incomes);

                r["Food & Dining"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Food & Dining"].ToString()) * (1 + costOfLiving);
                r["Bills & Utilities"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Bills & Utilities"].ToString()) * (1 + costOfLiving);
                r["Personal Care"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Personal Care"].ToString()) * (1 + costOfLiving);
                r["Leisure"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Leisure"].ToString()) * (1 + costOfLiving);
                r["Auto"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Auto"].ToString()) * (1 + costOfLiving);
                r["Housing"] = HousingCosts(i - retirementAge - adjust);
                r["Shopping"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Shopping"].ToString()) * (1 + costOfLiving);
                r["Vacation"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Vacation"].ToString()) * (0.8 + costOfLiving);
                r["Medical"] = MyDouble(budgetdt.Rows[i - retirementAge - adjust]["Medical"].ToString()) * (2 + costOfLiving);
                r["LOC Debt"] = (--locDebtPer > 0) ? MyDouble(budgetdt.Rows[i - retirementAge - adjust]["LOC Debt"].ToString()) : 0;
                r["CC Debt"] = (--ccDebtPer > 0) ? MyDouble(budgetdt.Rows[i - retirementAge - adjust]["CC Debt"].ToString()) : 0;

                expenses = 0;                                          // Given that income is based on expenses, so too then is income tax
                for (int j = maxIncomeIndex; j < budgetdt.Columns.Count - 1; j++)
                    expenses += MyDouble(r[j].ToString());
                r["Expenses"] = expenses;

                budgetdt.Rows.Add(r);
            }
            */
        }

        public void ReadButton_Click(object sender, EventArgs e)
        {
            ofd.Filter = "Funds|*.xlsx";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                //statusLabelValue.Text = ofd.FileName;

                if ((ReadExcelFile(ofd.FileName, 0)) == 0)
                {
                    generatePlanButton.Enabled = true;
                    UpdatePlanInfo();
                }
                else
                    statusLabelValue.Text = "Couldn't Read File";
            }
        }

        double propertyTax, primaryMortgagePmt, carPmtRate, liraLIFAmt, spouseLiraLIFAmt, monthlyContributions, otherPensions, cpptiming;
        DateTime startDate, dob, spouseDob;
        string LIRALIFLegislation;

        private void UpdatePlanInfo()
        {
            clientNameTextBox.Text = dtTable.Rows[0]["Current Amount"].ToString();
            dob = DateTime.Parse(dtTable.Rows[1]["Current Amount"].ToString());
            if (dtTable.Rows[1]["Rate/Spouse Amt"].ToString().Length <= 0)
                marriedCheckBox.Checked = false;
            else
            {
                spouseDob = DateTime.Parse(dtTable.Rows[1]["Rate/Spouse Amt"].ToString());
                spouseIncomeTextBox.Text = dtTable.Rows[15]["Rate/Spouse Amt"].ToString();
                spouseLiraLIFAmt = double.Parse(dtTable.Rows[16]["Rate/Spouse Amt"].ToString());
                spousalRSPBalanceTextBox.Text = dtTable.Rows[17]["Rate/Spouse Amt"].ToString();
                spousalTFSABalanceTextBox.Text = dtTable.Rows[18]["Rate/Spouse Amt"].ToString();
                spousalTFSARoomTextbox.Text = dtTable.Rows[18]["Amort Remaining"].ToString();
                spouseCashBalanceTextBox.Text = dtTable.Rows[19]["Rate/Spouse Amt"].ToString();
                spouseCashBalanceTextBox.Text = dtTable.Rows[19]["Rate/Spouse Amt"].ToString();
            }
            currentAgeTextBox.Text = dtTable.Rows[3]["Current Amount"].ToString();
            retirementAgeTextBox.Text = dtTable.Rows[4]["Current Amount"].ToString();
            deathAgeTextBox.Text = dtTable.Rows[5]["Current Amount"].ToString();
            startDate = DateTime.Parse(dtTable.Rows[6]["Current Amount"].ToString());
            primaryHomeValueTextBox.Text = dtTable.Rows[7]["Current Amount"].ToString();
            if (!(HomeOwnedCheckBox.Checked = dtTable.Rows[7]["Current Amount"].ToString().Length > 0))
                rentTextBox.Text = dtTable.Rows[7]["Payment"].ToString();
            propertyTax = double.Parse(dtTable.Rows[8]["Current Amount"].ToString());
            primaryHoaFeesTextBox.Text = dtTable.Rows[9]["Current Amount"].ToString();
            //rentTextBox.Text = dtTable.Rows[9]["Rate/Spouse Amt"].ToString();
            primaryMortgageAmtTextBox.Text = dtTable.Rows[10]["Current Amount"].ToString();
            primaryMortgageRateTextBox.Text = dtTable.Rows[10]["Rate/Spouse Amt"].ToString();
            primaryMortgageTermTextBox.Text = dtTable.Rows[10]["Payments/Year"].ToString();
            primaryMortgageAmortTextBox.Text = dtTable.Rows[10]["Amort Remaining"].ToString();
            primaryMortgagePmt = double.Parse(dtTable.Rows[10]["Payment"].ToString());
            lineOfCreditAmtTextBox.Text = dtTable.Rows[11]["Current Amount"].ToString();
            lineOfCreditRateTextBox.Text = dtTable.Rows[11]["Rate/Spouse Amt"].ToString();
            creditCardDebtTextBox.Text = dtTable.Rows[12]["Current Amount"].ToString();
            creditCardRateTextBox.Text = dtTable.Rows[12]["Rate/Spouse Amt"].ToString();
            carPmtTextBox.Text = dtTable.Rows[13]["Current Amount"].ToString();
            carPmtRate = double.Parse(dtTable.Rows[13]["Rate/Spouse Amt"].ToString());
            roiTextBox.Text = dtTable.Rows[14]["Rate/Spouse Amt"].ToString();
            roiCheckBox.Checked = false;
            incomeTextBox.Text = dtTable.Rows[15]["Current Amount"].ToString();
            liraLIFAmt = double.Parse(dtTable.Rows[16]["Current Amount"].ToString());
            LIRALIFLegislation = dtTable.Rows[16]["Payments/Year"].ToString();
            rspBalanceTextBox.Text = dtTable.Rows[17]["Current Amount"].ToString();
            rspContributions = MyDouble(dtTable.Rows[17]["Payment"].ToString());
            tfsaBalanceTextBox.Text = dtTable.Rows[18]["Current Amount"].ToString();
            tfsaRoomTextBox.Text = dtTable.Rows[18]["Payments/Year"].ToString();
            tfsaContributions = MyDouble(dtTable.Rows[18]["Payment"].ToString());
            cashBalanceTextBox.Text = dtTable.Rows[19]["Current Amount"].ToString();
            jointCashBalanceTextBox.Text = dtTable.Rows[19]["Payments/Year"].ToString();
            monthlyContributions = double.Parse(dtTable.Rows[21]["Current Amount"].ToString());
            otherPensions = double.Parse(dtTable.Rows[26]["Current Amount"].ToString());
            cpptiming = int.Parse(dtTable.Rows[27]["Current Amount"].ToString());
        }

        private double CalculateRESP()
        {
            double resp = 0;
            /* 
Canada education savings grant summary chart Adjusted income for 2019 	$47,630 or less 	more than $47,630
but less than $95,259 	More than $95,259
CESG on the first $500 of annual RESP contribution 	20% = $100 	10% = $50 	Beneficiary is not eligible
Basic CESG on the first $2,500 of annual RESP contribution 	20% = $500 	20% = $500 	20% = $500
Maximum yearly CESG depending on income and contributions 	$600 	$550 	$500
Lifetime maximum CESG for which you may qualify 	$7,200 	$7,200 	$7,200

Every child under age 18 who is a Canadian resident will accumulate $400 (for 1998 to 2006) and $500 (from 2007 and subsequent years) of CESG contribution room. Unused CESG contribution room is carried forward and used when RESP contributions are made in future years provided that the specific contribution requirements for beneficiaries who attain 16 or 17 years of age are met.

Beneficiaries qualify for a grant on the contributions made on their behalf up to the end of the calendar year in which they turn 17 years of age.
            lifetime limit per child = 50000;
            */

            return resp;
        }
 
        private double HousingCosts(int years)
        {
            double returnVal, fees;
            int mortgageIndex = years == 0 && mortgageTable.Rows.Count > 0 ? 1 : years;

            if (HomeOwnedCheckBox.Checked) {
                fees = propertyTax + MyDouble(primaryHoaFeesTextBox.Text);
                returnVal = fees + (fees * costOfLiving * years);
                if (years < mortgageTable.Rows.Count && mortgageTable.Rows.Count > 1)
                    returnVal += MyDouble(mortgageTable.Rows[mortgageIndex]["Payment"].ToString());
            }
            else
                returnVal = MyDouble(rentTextBox.Text) * 12;
            return (returnVal);
        }

        private double LifeInsurance()
        {
            bool haveDependants = MyInt(dependantChildrenTextBox.Text) > 0;
            double totalDebt = MyDouble(creditCardDebtTextBox.Text) + MyDouble(lineOfCreditAmtTextBox.Text) + funeralExpenses;
            double totalBalance = MyDouble(tfsaBalanceTextBox.Text) + MyDouble(rspBalanceTextBox.Text) + MyDouble(cashBalanceTextBox.Text) + MyDouble(spouseCashBalanceTextBox.Text) +
                MyDouble(spousalTFSABalanceTextBox.Text) + MyDouble(spousalRSPBalanceTextBox.Text) + MyDouble(jointCashBalanceTextBox.Text);

            if (MyInt(dependantChildrenTextBox.Text) > 0 && !divorced && MyDouble(spouseIncomeTextBox.Text) < lico)
                haveDependants = true;
            var haveSavings = totalBalance >= 500000;
            return haveDependants ? totalDebt - totalBalance : 0;
        }

        private void ExitButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
