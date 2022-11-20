using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using ClosedXML.Excel;
using System.Data.Odbc;
using static RecipeIPlacementCount.Model;

namespace RecipeIPlacementCount
{
    public partial class Form1 : Form
    {
        List<WO_Baan> WOlist = new List<WO_Baan>();
        public DataTable Pn1 = new DataTable();
        public DataTable Pn2 = new DataTable();
        public DataTable Pn3 = new DataTable();
        public DataTable Pn4 = new DataTable();
        public DataTable Pn5 = new DataTable();
        public DataTable Pn6 = new DataTable();
        public DataTable Pn7 = new DataTable();
        public DataTable Pn8 = new DataTable();
        public Dictionary<string, int> QtyRecipe = new Dictionary<string, int>();
        public Dictionary<string, int> QtyRecipeNew = new Dictionary<string, int>();
        public Dictionary<string, int> PLCRecipe = new Dictionary<string, int>();
        public Dictionary<string, int> DicPnLeft = new Dictionary<string, int>();
        public Dictionary<string, int> DicPnRight = new Dictionary<string, int>();
        public Dictionary<string, int> DicPnLeftSingle = new Dictionary<string, int>();
        public Dictionary<string, int> DicPnRightSingle = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn1 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn2 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn3 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn4 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn5 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn6 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn7 = new Dictionary<string, int>();
        public Dictionary<string, int> DicPn8 = new Dictionary<string, int>();
        int Qty_recipe;
        string recipe;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnRecipe_Click(object sender, EventArgs e)
        {
            Pn1.Clear();
            Pn2.Clear();
            Pn3.Clear();
            Pn4.Clear();
            Pn5.Clear();
            Pn6.Clear();
            Pn7.Clear();
            Pn8.Clear();

            fillingTable4Sipl1();
            fillingTable3Sipl1();
            fillingTable4Sipl2();
            fillingTable3Sipl2();
            fillingTable1Sipl1();
            fillingTable2Sipl1();
            fillingTable1Sipl2();
            fillingTable2Sipl2();
            int sums = Convert.ToInt32(lblt1s1.Text) + Convert.ToInt32(lblt1s2.Text) + Convert.ToInt32(lblt2s1.Text) + Convert.ToInt32(lblt2s2.Text) + Convert.ToInt32(lblt3s1.Text) + Convert.ToInt32(lblt3s2.Text) + Convert.ToInt32(lblt4s1.Text) + Convert.ToInt32(lblt4s2.Text);
            lblSMTSum.Text = sums.ToString();
            int smtSum = Convert.ToInt32(lblSMTSum.Text);
            int baanSum = Convert.ToInt32(lblBaanSum.Text);
            if(smtSum != baanSum)
            {
                MessageBox.Show("Error!! Quantity in the databases does not match");
                
            }

        }
        private void fillingTable4Sipl1()
        {
            /////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table4, location 0, Sipl1 to dataGridView1
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT       TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%')AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 0) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl1')AND (dbo.CRecipe.lLotSize =" + Convert.ToInt32(QtyRecipeNew[txtRecipe.Text]) + ")   GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn1.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn1.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn1.Rows.Count; i++)
            {
                if (Convert.ToInt32(Pn1.Rows[i][1]) % 2 == 0)
                {
                    Pn1.Rows[i][1] = (Convert.ToInt32(Pn1.Rows[i][1]) / 2) * Qty_recipe;
                }
                else
                {
                    Pn1.Rows[i][1] = ((Convert.ToInt32(Pn1.Rows[i][1]) + 1) / 2) * Qty_recipe;
                }


            }
            dataGridView1.DataSource = Pn1;
            db.closeConnection();
            DicPn1 = Pn1.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnLeft = DicPnLeft.Concat(DicPn1)
                  .GroupBy(x => x.Key)
                  .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView1.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView1.Rows[x].Cells[1].Value);
            }
            lblt4s1.Text = sum.ToString();
            Qty_PN_1.Text = (dataGridView1.Rows.Count - 1).ToString();
        }
        private void fillingTable3Sipl1()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 3, Location 2, Sipl1 to dataGridView2
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%') AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 2) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl1')AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ") GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn2.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn2.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn2.Rows.Count; i++)
            {
                if (Convert.ToInt32(Pn2.Rows[i][1]) % 2 == 0)
                {
                    Pn2.Rows[i][1] = (Convert.ToInt32(Pn2.Rows[i][1]) / 2) * Qty_recipe;
                }
                else
                {
                    Pn2.Rows[i][1] = ((Convert.ToInt32(Pn2.Rows[i][1]) + 1) / 2) * Qty_recipe;
                }


            }
            dataGridView2.DataSource = Pn2;
            db.closeConnection();
            DicPn2 = Pn2.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnLeft = DicPnLeft.Concat(DicPn2)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView2.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView2.Rows[x].Cells[1].Value);
            }
            lblt3s1.Text = sum.ToString();
            Qty_PN_2.Text = (dataGridView2.Rows.Count - 1).ToString();
        }
        private void fillingTable4Sipl2()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 4, Location 0, Sipl2 to dataGridView3
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%') AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 0) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl2')AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ") GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn3.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn3.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn3.Rows.Count; i++)
            {
                if (Convert.ToInt32(Pn3.Rows[i][1]) % 2 == 0)
                {
                    Pn3.Rows[i][1] = (Convert.ToInt32(Pn3.Rows[i][1]) / 2) * Qty_recipe;
                }
                else
                {
                    Pn3.Rows[i][1] = ((Convert.ToInt32(Pn3.Rows[i][1]) + 1) / 2) * Qty_recipe;
                }


            }
            dataGridView3.DataSource = Pn3;
            db.closeConnection();
            DicPn3 = Pn3.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnLeft = DicPnLeft.Concat(DicPn3)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView3.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView3.Rows[x].Cells[1].Value);
            }
            lblt4s2.Text = sum.ToString();
            Qty_PN_3.Text = (dataGridView3.Rows.Count - 1).ToString();
        }
        private void fillingTable3Sipl2()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 3, Location 2, Sipl2 to dataGridView4
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%')AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 2)AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ") AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl2') GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn4.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn4.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn4.Rows.Count; i++)
            {
                Pn4.Rows[i][1] = (Convert.ToInt32(Pn4.Rows[i][1])) * Qty_recipe;

            }
            dataGridView4.DataSource = Pn4;
            db.closeConnection();
            DicPn4 = Pn4.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnLeftSingle = DicPnLeftSingle.Concat(DicPn4)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView4.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView4.Rows[x].Cells[1].Value);
            }
            lblt3s2.Text = sum.ToString();
            Qty_PN_4.Text = (dataGridView4.Rows.Count - 1).ToString();

        }
        private void fillingTable1Sipl1()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 1, Location 1, Sipl1 to dataGridView5
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%') AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ") AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 1) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl1') GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn5.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn5.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn5.Rows.Count; i++)
            {
                Pn5.Rows[i][1] = (Convert.ToInt32(Pn5.Rows[i][1]) / 2) * Qty_recipe;

            }
            dataGridView5.DataSource = Pn5;
            db.closeConnection();
            DicPn5 = Pn5.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnRight = DicPnRight.Concat(DicPn5)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView5.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView5.Rows[x].Cells[1].Value);
            }
            lblt1s1.Text = sum.ToString();
            Qty_PN_5.Text = (dataGridView5.Rows.Count - 1).ToString();

        }
        private void fillingTable2Sipl1()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 2, Location 3, Sipl1 to dataGridView6
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%')AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ") AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 3) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl1') GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn6.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn6.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn6.Rows.Count; i++)
            {
                Pn6.Rows[i][1] = (Convert.ToInt32(Pn6.Rows[i][1]) / 2) * Qty_recipe;

            }
            dataGridView6.DataSource = Pn6;
            db.closeConnection();
            DicPn6 = Pn6.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnRight = DicPnRight.Concat(DicPn6)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView6.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView6.Rows[x].Cells[1].Value);
            }
            lblt2s1.Text = sum.ToString();
            Qty_PN_6.Text = (dataGridView6.Rows.Count - 1).ToString();
        }
        private void fillingTable1Sipl2()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 1, Location 1, Sipl2 to dataGridView7
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%')AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ")AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 1) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl2') GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn7.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn7.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn7.Rows.Count; i++)
            {
                Pn7.Rows[i][1] = (Convert.ToInt32(Pn7.Rows[i][1]) / 2) * Qty_recipe;

            }
            dataGridView7.DataSource = Pn7;
            db.closeConnection();
            DicPn7 = Pn7.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnRight = DicPnRight.Concat(DicPn7)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView7.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView7.Rows[x].Cells[1].Value);
            }
            lblt1s2.Text = sum.ToString();
            Qty_PN_7.Text = (dataGridView7.Rows.Count - 1).ToString();
        }
        private void fillingTable2Sipl2()
        {
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            ///Table 2, Location 3, Sipl2 to dataGridView8
            ////////////////////////////////////////////////////////////////////////////////////////////////////
            db.openConnection();
            db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN, COUNT(*) AS CNT  FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(dbo.AliasName.ObjectName LIKE N'%" + recipe + "%')AND (dbo.CRecipe.lLotSize =" + QtyRecipeNew[txtRecipe.Text] + ")AND (AliasName_3.ObjectName LIKE N'%" + txtSetup.Text.Trim() + "%') AND (dbo.CHeadSchedule.lHeadIndex = 3) AND (SUBSTRING(AliasName_1.ObjectName, 1, 5) = 'Sipl2') GROUP BY AliasName_2.ObjectName HAVING        (COUNT(*) > 0) ORDER BY PN";
            db.cmd.CommandText = db.sql;
            db.rd = db.cmd.ExecuteReader();
            Pn8.Load(db.rd);
            foreach (System.Data.DataColumn col in Pn8.Columns) col.ReadOnly = false;

            for (int i = 0; i < Pn8.Rows.Count; i++)
            {
                Pn8.Rows[i][1] = (Convert.ToInt32(Pn8.Rows[i][1])) * Qty_recipe;

            }
            dataGridView8.DataSource = Pn8;
            db.closeConnection();
            DicPn8 = Pn8.AsEnumerable().ToDictionary<DataRow, string, int>(row => row.Field<string>("PN"), row => row.Field<int>("CNT"));
            DicPnRightSingle = DicPnRightSingle.Concat(DicPn8)
                 .GroupBy(x => x.Key)
                 .ToDictionary(x => x.Key, x => x.Sum(y => y.Value));
            int sum = 0;
            for (int x = 0; x < dataGridView8.Rows.Count; x++)
            {
                sum += Convert.ToInt32(dataGridView8.Rows[x].Cells[1].Value);
            }
            lblt2s2.Text = sum.ToString();
            Qty_PN_8.Text = (dataGridView8.Rows.Count - 1).ToString();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.dataGridView11.EditMode = DataGridViewEditMode.EditOnEnter;
            txtSetup.Focus();
        }

        private void btnSetup_Click(object sender, EventArgs e)
        {
            string PL = txtSetup.Text.Trim();
            if (PL.Length != 6)
            {
                MessageBox.Show("PL Must Be 6 Digits.", "Warning");
                return;
            }
            try
            {
                db.openConnection();

                DataTable Pnt = new DataTable();
                db.sql = "SELECT        TOP (100) PERCENT  AliasName_2.ObjectName AS PN,  AliasName_1.ObjectName AS Station, dbo.CHeadSchedule.lHeadIndex AS Location, dbo.AliasName.ObjectName AS Recipe, AliasName_3.ObjectName AS Setup FROM            dbo.CRecipe INNER JOIN dbo.AliasName ON dbo.CRecipe.OID = dbo.AliasName.PID INNER JOIN dbo.CHeadSchedule ON dbo.CRecipe.OID = dbo.CHeadSchedule.PID INNER JOIN dbo.AliasName AS AliasName_1 ON dbo.CHeadSchedule.spStation = AliasName_1.PID INNER JOIN dbo.CHeadStep ON dbo.CHeadSchedule.OID = dbo.CHeadStep.PID INNER JOIN dbo.CPickupLink ON dbo.CRecipe.OID = dbo.CPickupLink.PID AND dbo.CHeadStep.lPickupLink = dbo.CPickupLink.lIndex INNER JOIN dbo.AliasName AS AliasName_2 ON dbo.CPickupLink.spComponentRef = AliasName_2.PID INNER JOIN dbo.CPlacementLink ON dbo.CRecipe.OID = dbo.CPlacementLink.PID AND dbo.CHeadStep.lPlacementLink = dbo.CPlacementLink.lIndex INNER JOIN dbo.CComponentPlacement ON dbo.CPlacementLink.spComponentPlacement = dbo.CComponentPlacement.OID INNER JOIN dbo.AliasName AS AliasName_3 ON dbo.CRecipe.spSetupRef = AliasName_3.PID WHERE(AliasName_3.ObjectName LIKE N'%" + txtSetup.Text + "%') ORDER BY PN";

                db.cmd.CommandText = db.sql;
                db.rd = db.cmd.ExecuteReader();
                Pnt.Load(db.rd);
                dataGridView9.DataSource = Pnt;
                List<string> list = dataGridView9.Rows
                                     .OfType<DataGridViewRow>()
                                     .Select(r => r.Cells[3].Value.ToString())
                                     .ToList();
                List<string> noDup = list.Distinct().ToList();
                foreach (string l in noDup)
                {
                    checkedListBox1.Items.Add(l);
                }
                db.closeConnection();

                GetDataFromBAAN(PL);
            }
            catch
            {
                return;
            }
        }
        private void GetDataFromBAAN(string p)
        {
            WOlist = new List<WO_Baan>();

            using (OdbcConnection DbConnection = new OdbcConnection("DSN=BAAN"))
            {
                try
                {
                    DbConnection.Open();
                    OdbcCommand DbCommand = DbConnection.CreateCommand();

                    string q = string.Format(@"SELECT tticst910400.t_pdno AS WO,
                                                      ttiitm001400.t_item AS PN,
                                                      ttisfc001400.t_qrdr AS Qty
                                            FROM baandb.ttccom010400 ttccom010400,
                                                 baandb.tticst910400 tticst910400,
                                                 baandb.ttiitm001400 ttiitm001400,
                                                 baandb.ttiitm200400 ttiitm200400,
                                                 baandb.ttiitm950400 ttiitm950400,
                                                 baandb.ttisfc001400 ttisfc001400,
                                                 baandb.ttisfc050400 ttisfc050400
                                            WHERE tticst910400.t_pdno = ttisfc001400.t_pdno AND
                                                  ttisfc001400.t_mitm = ttiitm001400.t_item AND
                                                  ttiitm001400.t_item = ttiitm200400.t_item AND
                                                  ttiitm200400.t_cuno = ttccom010400.t_cuno AND
                                                  ttiitm001400.t_item = ttiitm950400.t_item AND
                                                  ttisfc001400.t_npif = ttisfc050400.t_npif AND
                                                  ((tticst910400.t_pino='{0}') AND (ttiitm950400.t_mnum='999') AND (ttiitm950400.t_exdt <= Date('01/01/2001')) )
                                            ORDER BY tticst910400.t_pdno", p);

                    DbCommand.CommandText = q;
                    OdbcDataReader DbReader = DbCommand.ExecuteReader();

                    while (DbReader.Read())
                    {
                        WO_Baan wo = new WO_Baan();
                        wo.PL = p;
                        wo.WO = DbReader.GetString(0);
                        wo.PN = DbReader.GetString(1).Trim();
                        wo.Qty = DbReader.GetString(2).Trim();

                        WOlist.Add(wo);
                    }
                    DbReader.Close();

                    foreach (var w in WOlist)
                    {
                        q = string.Format(@"SELECT cst.t_sitm AS pcb,
                                                   cst.t_revi AS pcbRev,
                                                   cst.t_cwar AS WH,
                                                   itm.t_oqmf AS Pattern,
                                                   cst.t_qucs + cst.t_issu + cst.t_subd AS Plc,
                                                   itm.t_csgp
                                                   FROM ttisfc001400 AS sfc
                                                   INNER JOIN tticst001400 AS cst
                                                   ON sfc.t_pdno = cst.t_pdno
                                                   INNER JOIN ttiitm001400 AS itm
                                                   ON cst.t_sitm = itm.t_item
                                                   WHERE sfc.t_pdno = '{0}'
                                                   AND cst.t_opno = 10", w.WO);

                        DbCommand.CommandText = q;
                        DbReader = DbCommand.ExecuteReader();

                        while (DbReader.Read())
                        {
                            if (DbReader.GetString(5).Trim() != "PCB")
                                w.Placements += Convert.ToInt32(DbReader.GetValue(4).ToString().Trim().Replace(".", ""));
                            else if (DbReader.GetString(5).Trim() == "PCB")
                            {

                                w.Pattern = DbReader.GetValue(3).ToString().Trim();
                            }
                        }
                        DbReader.Close();
                    }
                    DbReader.Close();
                    DbCommand.Dispose();
                    DbConnection.Close();

                    foreach (var w in WOlist)
                    {
                        if (QtyRecipe.ContainsKey(w.PN.ToString().Trim()))
                        {
                            if (Convert.ToInt32(w.Qty) % Convert.ToInt32(w.Pattern) == 0)
                            {
                                int qtyNew = Convert.ToInt32(w.Qty) / Convert.ToInt32(w.Pattern);
                                string str = w.PN.ToString().Trim() + "_" + w.WO;
                                QtyRecipeNew[w.PN.ToString().Trim()] = Convert.ToInt32(w.Qty);
                                QtyRecipe[str] = qtyNew;
                            }
                            else
                            {
                                int qtyNew = (Convert.ToInt32(w.Qty) / Convert.ToInt32(w.Pattern)) + 1;
                                string str = w.PN.ToString().Trim() + "_" + w.WO;
                                QtyRecipeNew[w.PN.ToString().Trim()] = Convert.ToInt32(w.Qty);
                                QtyRecipe[str] = qtyNew;
                            }

                        }
                        else
                        {
                            if (Convert.ToInt32(w.Qty) % Convert.ToInt32(w.Pattern) == 0)
                            {
                                int qtyNew = Convert.ToInt32(w.Qty) / Convert.ToInt32(w.Pattern);
                                QtyRecipe.Add(w.PN.ToString().Trim() + "_" + w.WO, qtyNew);
                                QtyRecipeNew[w.PN.ToString().Trim()] = Convert.ToInt32(w.Qty);

                            }
                            else
                            {
                                int qtyNew = (Convert.ToInt32(w.Qty) / Convert.ToInt32(w.Pattern)) + 1;
                                QtyRecipe.Add(w.PN.ToString().Trim() + "_" + w.WO, qtyNew);
                                QtyRecipeNew[w.PN.ToString().Trim()] = Convert.ToInt32(w.Qty);
                            }

                        }
                        if (PLCRecipe.ContainsKey(w.PN.ToString().Trim()))
                        {
                            int plcNew = Convert.ToInt32(w.Placements);
                            string str = w.PN.ToString().Trim() + "_" + w.WO;
                            PLCRecipe[str] = plcNew;
                        }
                        else
                        {
                            int plcNew = Convert.ToInt32(w.Placements);
                            PLCRecipe.Add(w.PN.ToString().Trim() + "_" + w.WO, plcNew);
                        }
                    }
                    dataGridView11.Columns.Add("Key", "Recipe");
                    dataGridView11.Columns.Add("Values", "Qty");
                    foreach (KeyValuePair<string, int> item in QtyRecipe)
                    {
                        dataGridView11.Rows.Add(item.Key, item.Value);
                    }

                }
                catch (OdbcException ex)
                {
                    MessageBox.Show("connection to the DSN '" + "BAAN" + "' failed." + ex.Message);
                    return;
                }
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            var wb = new XLWorkbook();
            var ws1 = wb.Worksheets.Add("RecipePlacementCountWH");

            ws1.Cell(1, 2).Value = "LEFT";
            ws1.Range(1, 2, 1, 3).Merge().AddToNamed("Titles");
            ws1.Cell(2, 2).Value = DicPnLeft.AsEnumerable();

            ws1.Cell(1, 9).Value = "LEFT SINGLE";
            ws1.Range(1, 9, 1, 10).Merge().AddToNamed("Titles");
            ws1.Cell(2, 9).Value = DicPnLeftSingle.AsEnumerable();



            ws1.Cell(1, 6).Value = "RIGHT";
            ws1.Range(1, 6, 1, 7).Merge().AddToNamed("TitlesS");
            ws1.Cell(2, 6).Value = DicPnRight.AsEnumerable();

            ws1.Cell(1, 12).Value = "RIGHT SINGLE";
            ws1.Range(1, 12, 1, 13).Merge().AddToNamed("TitlesS");
            ws1.Cell(2, 12).Value = DicPnRightSingle.AsEnumerable();



            var ws = wb.Worksheets.Add("I-PLC_" + txtRecipe.Text);
            var dataTable1 = Pn1;
            ws.Cell(4, 1).Value = "Table 4 Sipl1-X4S_S";
            ws.Range(4, 1, 4, 2).Merge().AddToNamed("Titles");
            ws.Cell(5, 1).Value = dataTable1.AsEnumerable();

            var dataTable2 = Pn2;
            ws.Cell(4, 4).Value = "Table 3 Sipl1-X4S_S";
            ws.Range(4, 4, 4, 5).Merge().AddToNamed("Titles");
            ws.Cell(5, 4).Value = dataTable2.AsEnumerable();

            var dataTable3 = Pn3;
            ws.Cell(4, 7).Value = "Table 4 Sipl2-X4S_S";
            ws.Range(4, 7, 4, 8).Merge().AddToNamed("Titles");
            ws.Cell(5, 7).Value = dataTable3.AsEnumerable();

            var dataTable4 = Pn4;
            ws.Cell(4, 10).Value = "Table 3 Sipl2-X4S_S";
            ws.Range(4, 10, 4, 11).Merge().AddToNamed("Titles");
            ws.Cell(5, 10).Value = dataTable4.AsEnumerable();

            var dataTable5 = Pn5;
            ws.Cell(4, 13).Value = "Table 1 Sipl1-X4S_S";
            ws.Range(4, 13, 4, 14).Merge().AddToNamed("Titles");
            ws.Cell(5, 13).Value = dataTable5.AsEnumerable();

            var dataTable6 = Pn6;
            ws.Cell(4, 16).Value = "Table 2 Sipl1-X4S_S";
            ws.Range(4, 16, 4, 17).Merge().AddToNamed("Titles");
            ws.Cell(5, 16).Value = dataTable6.AsEnumerable();

            var dataTable7 = Pn7;
            ws.Cell(4, 19).Value = "Table 1 Sipl2-X4S_S";
            ws.Range(4, 19, 4, 20).Merge().AddToNamed("Titles");
            ws.Cell(5, 19).Value = dataTable7.AsEnumerable();

            var dataTable8 = Pn8;
            ws.Cell(4, 22).Value = "Table 2 Sipl2-X4S_S";
            ws.Range(4, 22, 4, 23).Merge().AddToNamed("Titles");
            ws.Cell(5, 22).Value = dataTable8.AsEnumerable();






            var titlesStyle = wb.Style;
            titlesStyle.Font.Bold = true;
            titlesStyle.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle.Fill.BackgroundColor = XLColor.Cyan;
            wb.NamedRanges.NamedRange("Titles").Ranges.Style = titlesStyle;

            var titlesStyle1 = wb.Style;
            titlesStyle1.Font.Bold = true;
            titlesStyle1.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
            titlesStyle1.Fill.BackgroundColor = XLColor.Tomato;
            wb.NamedRanges.NamedRange("TitlesS").Ranges.Style = titlesStyle1;


            //ws.Columns().AdjustToContents();


            //wb.SaveAs(@"C:\Users\migkbron\Desktop\temp\" + txtRecipe.Text +"_"+ DateTime.Now.ToString("yyyy-dd-M--HH-mm-ss")+".xlsx");
            var saveFileDialog = new SaveFileDialog
            {
                FileName = "PL-" + txtSetup.Text + ".xlsx",
                Filter = "Excel files|*.xlsx",
                Title = "Save an Excel File"
            };

            saveFileDialog.ShowDialog();

            if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))

                wb.SaveAs(saveFileDialog.FileName);

        }

        private void dataGridView11_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = e.RowIndex;
            DataGridViewRow selectedRow = dataGridView11.Rows[index];
            string str = selectedRow.Cells[0].Value.ToString();
            string strqty = selectedRow.Cells[1].Value.ToString();

            Qty_recipe = Convert.ToInt32(strqty);
            int indexSub = str.IndexOf('_');
            recipe = str.Substring(0, indexSub);

            txtRecipe.Text = recipe;
            int plc = PLCRecipe[str];
            lblBaanSum.Text = plc.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Restart();
        }
    }
}
