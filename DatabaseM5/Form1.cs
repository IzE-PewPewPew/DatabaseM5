using Bunifu.UI.WinForms;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace DatabaseM5
{
    public partial class Form1 : Form
    {
        M5 d = new M5();
        SqlCommand cmd;
        SqlDataAdapter adapter;
        DataTable dt;

        private bool dragging = false;
        private Point dragCursorPoint;
        private Point dragFormPoint;

        public Form1()
        {
            d.Connect();
            InitializeComponent();
            LoadData();
        }

        //public void LoadData()
        //{
        //    dgc.DataSource = null;
        //    cmd = new SqlCommand("RSup",d.Connection);
        //    cmd.CommandType = CommandType.StoredProcedure;

        //    SqlDependency dep = new SqlDependency(cmd);
        //    dep.OnChange += new OnChangeEventHandler(OnChange);

        //    adapter = new SqlDataAdapter(cmd);
        //    dt = new DataTable();
        //    adapter.Fill(dt);

        //    dgc.DataSource = dt;
        //}

        public void LoadData()
        {
            if (d.idTab == 2)
            {
                gggg.DataSource = null;
                using (cmd = new SqlCommand("RPro", d.Connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        gggg.DataSource = dt;
                    }
                }
            }
            else if (d.idTab == 3)
            {
                gggg.DataSource = null;
                using (cmd = new SqlCommand("RSup", d.Connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        gggg.DataSource = dt;
                    }
                }
            }
            else if (d.idTab == 1)
            {
                gggg.DataSource = null;
                using (cmd = new SqlCommand("RCus", d.Connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        gggg.DataSource = dt;
                    }
                }
            }
            else if (d.idTab == 4)
            {
                gggg.DataSource = null;
                using (cmd = new SqlCommand("RSta", d.Connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        DataTable dt = new DataTable();
                        dt.Load(reader);
                        gggg.DataSource = dt;
                    }
                }
            }
        }


        public void OnChange(object caller, SqlNotificationEventArgs e)
        {
            if (this.InvokeRequired)
            {
                gggg.BeginInvoke(new MethodInvoker(LoadData));
            }
            else
            {
                LoadData();
            }
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            panel1.Location = new Point(this.Width - panel1.Width - 5, 3);
            lbprocname.Visible = false;
            labelqty.Visible = false;
            lbupis.Visible = false;
            lbsup.Visible = false;
            lbqty1.Visible = false;
            txtsup.Visible = false;
            bunifuLabel15.Visible = false;
            d.idTab = 1;
            tabFunction(1);

        }


        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }



        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (d.idTab == 3)
            {
                cmd = new SqlCommand("InsertSup", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@su", txtName.Text);
                cmd.Parameters.AddWithValue("@ad", txtaddress.Text);
                cmd.Parameters.AddWithValue("@con", txtCon.Text);

                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Added Supplier Data.";
            }
            if (d.idTab == 2)
            {
                cmd = new SqlCommand("InsertProc", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@pc", txtName.Text);
                cmd.Parameters.AddWithValue("@us", txtaddress.Text);
                cmd.Parameters.AddWithValue("@qy", txtCon.Text);
                cmd.Parameters.AddWithValue("@sp", txtsup.Text);
                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Added Product Data.";
            }
            if (d.idTab == 1)
            {
                cmd = new SqlCommand("InsertCus", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@cn", txtName.Text);
                cmd.Parameters.AddWithValue("@ct", txtCon.Text);
                //cmd.Parameters.AddWithValue("@sp", txtsup.Text);
                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Added Product Data.";
            }
            if (d.idTab == 4)
            {
                cmd = new SqlCommand("InsertSta", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@fn", txtName.Text);
                if (rbfmale.Checked == true)
                    cmd.Parameters.AddWithValue("@gd", "F");
                else
                    cmd.Parameters.AddWithValue("@gd", "M");
                cmd.Parameters.AddWithValue("@db", dobp.Value);
                cmd.Parameters.AddWithValue("@pos", txtsup.Text);
                cmd.Parameters.AddWithValue("@sr", txtCon.Text);
                if (cbswork.Checked == true)
                    cmd.Parameters.AddWithValue("@sw", 1);
                else
                    cmd.Parameters.AddWithValue("@sw", 0);
                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Added Staff Data.";
            }

            Clear();
            //MessageBox.Show("Data Stored.");
        }

        private void panel3_Click(object sender, EventArgs e)
        {
            if (d.idTab == 3)
            {
                if (rbID.Checked && !rbName.Checked)
                {
                    if (string.IsNullOrWhiteSpace(txtID.Text))
                    {
                        LoadData();
                    }
                    else if (int.TryParse(txtID.Text, out _))
                    {
                        (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("ID = '{0}'",
                        txtID.Text.Replace("'", "''"));
                        lbstat.Text = "Successfully search supplier by ID.";
                    }
                    else
                    {
                        lbstat.Text = "Error";
                    }
                }
                else if (!rbID.Checked && rbName.Checked)
                {
                    (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("Supplier LIKE '%{0}%'",
                    txtName.Text.Replace("'", "''"));
                    lbstat.Text = "Successfully search supplier by Name.";
                }
            }
            else if (d.idTab == 2)
            {
                if (rbID.Checked && !rbName.Checked)
                {
                    if (string.IsNullOrWhiteSpace(txtID.Text))
                    {
                        LoadData();
                    }
                    else if (int.TryParse(txtID.Text, out _))
                    {
                        (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("Code = '{0}'",
                        txtID.Text.Replace("'", "''"));
                        lbstat.Text = "Successfully search product by ID.";
                    }
                    else
                    {
                        lbstat.Text = "Error";
                    }
                }
                else if (!rbID.Checked && rbName.Checked)
                {
                    (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '%{0}%'",
                    txtName.Text.Replace("'", "''"));
                    lbstat.Text = "Successfully search product by Name.";
                }
            }
            else if (d.idTab == 1)
            {
                if (rbID.Checked && !rbName.Checked)
                {
                    if (string.IsNullOrWhiteSpace(txtID.Text))
                    {
                        LoadData();
                    }
                    else if (int.TryParse(txtID.Text, out _))
                    {
                        (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("ID = '{0}'",
                        txtID.Text.Replace("'", "''"));
                        lbstat.Text = "Successfully search Customer by ID.";
                    }
                    else
                    {
                        lbstat.Text = "Error";
                    }
                }
                else if (!rbID.Checked && rbName.Checked)
                {
                    (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '%{0}%'",
                    txtName.Text.Replace("'", "''"));
                    lbstat.Text = "Successfully search Customer by Name.";
                }
            }
            else if (d.idTab == 4)
            {

                if (rbID.Checked && !rbName.Checked)
                {
                    if (string.IsNullOrWhiteSpace(txtID.Text))
                    {
                        LoadData();
                    }
                    else if (int.TryParse(txtID.Text, out _))
                    {
                        (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("ID = '{0}'",
                        txtID.Text.Replace("'", "''"));
                        lbstat.Text = "Successfully search Staff by ID.";
                    }
                    else
                    {
                        lbstat.Text = "Error";
                    }
                }
                else if (!rbID.Checked && rbName.Checked)
                {
                    (gggg.DataSource as DataTable).DefaultView.RowFilter = string.Format("Name LIKE '%{0}%'",
                    txtName.Text.Replace("'", "''"));
                    lbstat.Text = "Successfully search Staff by Name.";
                }

            }
            Clear();
        }

        void Clear()
        {
            txtID.Text = "";
            txtName.Text = "";
            txtaddress.Text = "";
            txtCon.Text = "";
            txtsup.Text = "";
            rbfmale.Checked = false;
            rbmale.Checked = false;
            cbswork.Checked = false;
        }

        private void panel4_Click(object sender, EventArgs e)
        {
            LoadData();
            lbstat.Text = "Table Reset";
            Clear();
        }

        private void bunifuGradientPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            dragging = true;
            dragCursorPoint = Cursor.Position;
            dragFormPoint = this.Location;
        }

        private void bunifuGradientPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (dragging)
            {
                Point dif = Point.Subtract(Cursor.Position, new Size(dragCursorPoint));
                this.Location = Point.Add(dragFormPoint, new Size(dif));
            }
        }

        private void bunifuGradientPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            dragging = false;
        }

        private void bunifuLabel10_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel6_Click(object sender, EventArgs e)
        {
            if (d.idTab == 3)
            {
                if(rbID.Checked == true)
                {
                    cmd = new SqlCommand("DeleteSup", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", txtID.Text);
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Supplier by ID.";
                }
                else
                {
                    cmd = new SqlCommand("DeleteSupByName", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@su", Convert.ToString(txtName.Text));
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Supplier by name.";
                }
            }
            else if (d.idTab == 2)
            {
                if (rbID.Checked == true)
                {
                    cmd = new SqlCommand("DeletePro", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", txtID.Text);
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Product by ID.";
                }
                else
                {
                    cmd = new SqlCommand("DeleteProByName", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@pc", Convert.ToString(txtName.Text));
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Product by name.";
                }
            }
            else if (d.idTab == 1)
            {
                if (rbID.Checked == true)
                {
                    cmd = new SqlCommand("DeleteCus", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", txtID.Text);
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Customer by ID.";
                }
                else
                {
                    cmd = new SqlCommand("DeleteCusByName", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@cn", Convert.ToString(txtName.Text));
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Customer by name.";
                }
            }
            else if (d.idTab == 4)
            {
                if(rbID.Checked == true)
                {
                    cmd = new SqlCommand("DeleteSta", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@id", txtID.Text);
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Staff by ID.";
                }
                else
                {
                    cmd = new SqlCommand("DeleteStaByName", d.Connection);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.AddWithValue("@fn", Convert.ToString(txtName.Text));
                    cmd.ExecuteNonQuery();
                    LoadData();
                    lbstat.Text = "Deleted Staff by name.";
                }
            }
            Clear();
        }
        

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (d.idTab == 3)
            {
                cmd = new SqlCommand("UpdateSup", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@su", txtName.Text);
                cmd.Parameters.AddWithValue("@ad", txtaddress.Text);
                cmd.Parameters.AddWithValue("@con", txtCon.Text);

                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Updated Supplier by ID.";
            }
            else if (d.idTab == 2)
                {
                cmd = new SqlCommand("UpdatePro", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@pc", txtName.Text);
                cmd.Parameters.AddWithValue("@us", txtaddress.Text);
                cmd.Parameters.AddWithValue("@qy", txtCon.Text);
                cmd.Parameters.AddWithValue("@sp", txtsup.Text);
                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Updated Product by ID.";
            }else if (d.idTab == 1)
            {
                cmd = new SqlCommand("UpdateCus", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@cn", txtName.Text);
                cmd.Parameters.AddWithValue("@ct", txtCon.Text);

                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Updated Customer by ID.";
            }
            else if (d.idTab == 4)
            {
                cmd = new SqlCommand("UpdateSta", d.Connection);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.AddWithValue("@id", txtID.Text);
                cmd.Parameters.AddWithValue("@fn", txtName.Text);
                if (rbfmale.Checked == true)
                    cmd.Parameters.AddWithValue("@gd", "F");
                else
                    cmd.Parameters.AddWithValue("@gd", "M");
                cmd.Parameters.AddWithValue("@db", dobp.Value);
                cmd.Parameters.AddWithValue("@pos", txtsup.Text);
                cmd.Parameters.AddWithValue("@sr", txtCon.Text);
                if (cbswork.Checked == true)
                    cmd.Parameters.AddWithValue("@sw", 1);
                else
                    cmd.Parameters.AddWithValue("@sw", 0);
                cmd.ExecuteNonQuery();
                LoadData();
                lbstat.Text = "Updated Staff Data by ID.";
            }
            Clear();
        }

 
        void tabFunction(int id)
        {
            if (d.idTab == 2)
            {
                lbdob.Visible = false;
                dobp.Visible = false;
                lbgender.Visible = false;
                rbfmale.Visible = false;
                rbmale.Visible = false;
                cbswork.Visible = false;
                lbswork.Visible = false;

                supActive.Visible = false;
                cusActive.Visible = false;
                proActive.Visible = true;
                staActive.Visible = false;
                lbprocname.Visible = true;
                labelqty.Visible = true;
                labelqty.Text = "Quantity";
                lbupis.Visible = true;
                lbsup.Visible = true;
                lbqty1.Visible = true;
                bunifuLabel15.Visible = true;
                txtsup.Visible = true;
                bunifuLabel11.Visible = false;
                lbcon.Visible = false;
                lbname.Visible = false;
                lbaddress.Visible = false;
                txtaddress.Visible = true;
                lbpos.Visible = false;
                bunifuGroupBox3.Size = new System.Drawing.Size(799, 173);
                lbqty1.Location = new Point(653, 97);
                LoadData();
                lbstat.Text = "Changed to Product Table.";
            }
            else if (d.idTab == 3)
            {
                lbdob.Visible = false;
                dobp.Visible = false;
                lbgender.Visible = false;
                rbfmale.Visible = false;
                rbmale.Visible = false;
                cbswork.Visible = false;
                lbswork.Visible = false;

                supActive.Visible = true;
                cusActive.Visible = false;
                proActive.Visible = false;
                staActive.Visible = false;
                lbprocname.Visible = false;
                labelqty.Visible = false;
                lbupis.Visible = false;
                lbsup.Visible = false;
                lbqty1.Visible = false;
                txtsup.Visible = false;
                bunifuLabel15.Visible = false;
                lbcon.Visible = true;
                lbname.Visible = true;
                lbname.Text = "Name";
                lbaddress.Visible = true;
                bunifuLabel11.Visible = true;
                txtaddress.Visible = true;
                lbpos.Visible = false;
                bunifuGroupBox3.Size = new System.Drawing.Size(799, 173);
                lbqty1.Location = new Point(653, 97);
                LoadData();
                lbstat.Text = "Changed to Supplier Table.";
            }else if (d.idTab == 1)
            {
                lbdob.Visible = false;
                dobp.Visible = false;
                lbgender.Visible = false;
                rbfmale.Visible = false;
                rbmale.Visible = false;
                cbswork.Visible = false;
                lbswork.Visible = false;

                supActive.Visible = false;
                cusActive.Visible = true;
                proActive.Visible = false;
                staActive.Visible = false;
                lbprocname.Visible = false;
                labelqty.Visible = false;
                lbupis.Visible = false;
                lbsup.Visible = false;
                lbqty1.Visible = false;
                txtsup.Visible = false;
                bunifuLabel15.Visible = false;
                lbcon.Visible = true;
                lbname.Visible = true;
                lbname.Text = "Name";
                lbaddress.Visible = false;
                bunifuLabel11.Visible = false;
                txtaddress.Visible = false;
                lbpos.Visible = false;
                bunifuGroupBox3.Size = new System.Drawing.Size(799, 173);
                lbqty1.Location = new Point(653, 97);
                LoadData();
                lbstat.Text = "Changed to Customer Table.";
            }else if(d.idTab == 4)
            {
                lbdob.Visible = true;
                dobp.Visible = true;
                lbgender.Visible = true;
                rbfmale.Visible = true;
                rbmale.Visible = true;
                cbswork.Visible = true;
                lbswork.Visible = true;

                supActive.Visible = false;
                cusActive.Visible = false;
                proActive.Visible = false;
                staActive.Visible = true;
                lbsup.Visible= false;
                txtsup.Visible= true;
                lbpos.Visible = true;
                lbprocname.Visible = false;
                lbname.Text = "FullName";
                lbname.Visible= true;
                labelqty.Visible = true;
                labelqty.Text = "Salary";
                lbcon.Visible = false;
                lbupis.Visible = false;
                lbaddress.Visible= false;
                bunifuLabel15.Visible= false;
                bunifuLabel11.Visible  = false;
                txtaddress.Visible= false;
                bunifuGroupBox3.Size = new System.Drawing.Size(799, 221);
                lbqty1.Location = new Point(582, 97);
                LoadData();
                lbstat.Text = "Changed to Staff Table.";
            }
            Clear();
        }

        private void cusTab_Click(object sender, EventArgs e)
        {
            d.idTab = 1;
            tabFunction(1);
        }

        private void cusSup_Click(object sender, EventArgs e)
        {
            d.idTab = 3;
            tabFunction(3);
        }

        private void cusPro_Click(object sender, EventArgs e)
        {
            d.idTab = 2;
            tabFunction(2);
        }

        private void cusStaff_Click(object sender, EventArgs e)
        {
            d.idTab = 4;
            tabFunction(4);
        }

        private void lbtab1_Click(object sender, EventArgs e)
        {
            d.idTab = 1;
            tabFunction(1);
        }

        private void lbtab2_Click(object sender, EventArgs e)
        {
            d.idTab = 3;
            tabFunction(3);
        }

        private void lbtab3_Click(object sender, EventArgs e)
        {
            d.idTab = 2;
            tabFunction(2);
        }

        private void lbtab4_Click(object sender, EventArgs e)
        {
            d.idTab = 4;
            tabFunction(4);
        }

        private void gggg_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.RowIndex < gggg.Rows.Count)
            {
                DataGridViewRow row = gggg.Rows[e.RowIndex];
                if (d.idTab == 3)
                {
                    txtID.Text = row.Cells[0].Value.ToString();
                    txtName.Text = row.Cells[1].Value.ToString();
                    txtaddress.Text = row.Cells[2].Value.ToString();
                    txtCon.Text = row.Cells[3].Value.ToString();
                }
                else if (d.idTab == 2)
                {
                    txtID.Text = row.Cells[0].Value.ToString();
                    txtName.Text = row.Cells[1].Value.ToString();
                    txtCon.Text = row.Cells[2].Value.ToString();
                    txtaddress.Text = row.Cells[3].Value.ToString();
                    txtsup.Text = row.Cells[4].Value.ToString();
                }
                else if (d.idTab == 1)
                {
                    txtID.Text = row.Cells[0].Value.ToString();
                    txtName.Text = row.Cells[1].Value.ToString();
                    txtCon.Text = row.Cells[2].Value.ToString();
                }
                else if (d.idTab == 4)
                {
                    txtID.Text = row.Cells[0].Value.ToString();
                    txtName.Text = row.Cells[1].Value.ToString();
                    if (row.Cells[2].Value.ToString() == "M")
                    {
                        rbmale.Checked = true;
                        rbfmale.Checked = false;
                    }
                    else if (row.Cells[2].Value.ToString() == "F")
                    {
                        rbmale.Checked = false;
                        rbfmale.Checked = true;
                    }
                    DateTime dateOfBirth;
                    if (DateTime.TryParse(row.Cells[3].Value.ToString(), out dateOfBirth))
                    {
                        dobp.Value = dateOfBirth;
                    }
                    txtsup.Text = row.Cells[4].Value.ToString();
                    txtCon.Text = row.Cells[5].Value.ToString();
                    bool isWorking;
                    if (bool.TryParse(row.Cells[6].Value.ToString(), out isWorking))
                    {
                        cbswork.Checked = isWorking;
                    }
                    else
                    {
                        if (row.Cells[6].Value.ToString() == "1")
                        {
                            cbswork.Checked = true;
                        }
                        else if (row.Cells[6].Value.ToString() == "0")
                        {
                            cbswork.Checked = false;
                        }
                        else
                        {
                            lbstat.Text = "Invalid BIT format in the selected cell.";
                        }
                    }
                }
            }
        }

        private void lbswork_Click(object sender, EventArgs e)
        {
            if (!cbswork.Checked) cbswork.Checked = true;
            else if (cbswork.Checked) cbswork.Checked = false;
        }
    }
}
