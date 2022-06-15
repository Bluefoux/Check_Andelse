using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Test20;

namespace Check_Andelse
{
    public partial class Form1 : Form
    {
        public DataTable missing_tbl = new DataTable();

        public Form1()
        {
            InitializeComponent();
            missing_tbl.Columns.Add("NAME");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog choose_file = new OpenFileDialog();

            if (choose_file.ShowDialog() == DialogResult.OK)
            {
                this.textBox1.Text = choose_file.FileName;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void compare_table(DataTable var_tab, DataTable obj_tab, string file_name, string var_file_name)
        {
            ClassExcelFile mah_class = new ClassExcelFile();
            DataTable tab = new DataTable();
            tab.Columns.Add("Name");

            //check if name(obj_tab) in var_tab
            foreach (DataRow dr_obj in obj_tab.Rows)
            {
                if(!exist_in_tb(dr_obj[1].ToString(), var_tab, "NAME"))
                {
                    DataRow rowObj = tab.NewRow();
                    rowObj["Name"] = dr_obj[1].ToString();
                    tab.Rows.Add(rowObj);
                }
                
            }
            mah_class.ExportToExcelFile(tab, file_name + "_missing_in_obj.xlsx");
            tab = null;
            tab = new DataTable();
            tab.Columns.Add("Name");
            //check if name(var_tab) in obj_tab
            DataTable sec_var_tab = new DataTable();
            foreach (DataRow dr_Variable in var_tab.Rows)
            {
                if (!exist_in_tb(dr_Variable[0].ToString(), obj_tab, "Objekt"))
                {
                    DataRow rowObj = tab.NewRow();
                    rowObj["Name"] = dr_Variable[0].ToString();
                    tab.Rows.Add(rowObj);
                }
            }
            mah_class.ExportToExcelFile(tab, var_file_name + "_missing_in_var.xlsx");
        }

        private string find_andelse(DataRow obj_row)
        {
            ClassExcelFile mah_class = new ClassExcelFile();
            DataTable and_table = new DataTable();
            string path = Environment.CurrentDirectory;
            string hjalp_fil = path + "\\Hjalpfiler\\Variable_Change_Signal.xlsx";
            and_table = mah_class.ConvertExcelToDataTable(hjalp_fil);
            string var = "";
            string obj = obj_row[0].ToString();
            int len_obj = obj.Length;
            foreach (DataRow dr_Variable in and_table.Rows)
            { 
                var = dr_Variable[1].ToString();
                if(obj.EndsWith(var) == true && var != "")
                {
                    return var;
                }
            }
            DataRow rowObj = missing_tbl.NewRow();

            rowObj["Name"] = obj;
            missing_tbl.Rows.Add(rowObj);
            return var;
        }

        private bool exist_in_tb(string name, DataTable obj_tab, string tab_name)
        {
            DataRow[] drExist = obj_tab.Select(tab_name + "= '" + name + "'");
            if (drExist.Length != 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private DataTable remove_andelse(DataTable obj_tab) //ta bort allt efter kolumn 2 i hjälpfil
        {
            string andelse;
            string cl;
            //foreach (DataRow dr_Variable in obj_tab.Rows)
            for (int i = obj_tab.Rows.Count - 1; i >= 0; i--)
            {
                DataRow dr_Variable = obj_tab.Rows[i];
                andelse = find_andelse(dr_Variable);
                cl = dr_Variable[0].ToString();
                if (cl.Length >= andelse.Length && cl != "")
                {
                    cl = cl.Remove(cl.Length - andelse.Length, andelse.Length);
                }
                //check if already in datatable
                //dr_Variable.SetField("NAME", cl);
                if (!exist_in_tb(cl, obj_tab, "NAME"))
                {
                    dr_Variable[0] = cl;
                }
                else
                {
                    obj_tab.Rows[i].Delete();
                }
            }
            obj_tab.AcceptChanges();
            return obj_tab;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string file_name = this.textBox1.Text;
            string var_name = this.textBox2.Text;
            if (file_name == "" || var_name == "")
            {
                return;
            }

            string[] splt_obj = file_name.Split('.');
            string[] splt_var = var_name.Split('.');
            DataTable obj_table = new DataTable();
            DataTable var_table = new DataTable();
            ClassExcelFile mah_class = new ClassExcelFile();
            switch (splt_obj[splt_obj.Length-1])
            {
                case "DBF":
                    switch (splt_var[splt_var.Length - 1])
                    {
                        case "DBF":
                            var_table = mah_class.GetYourData_DBF(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.GetYourData_DBF(file_name);
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsx":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.GetYourData_DBF(file_name);
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsm":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.GetYourData_DBF(file_name);
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        default:
                            break;
                    }
                    break;
                case "xlsx":
                    switch (splt_var[splt_var.Length - 1])
                    {
                        case "DBF":
                            var_table = mah_class.GetYourData_DBF(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsx":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsm":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        default:
                            break;
                    }
                    break;
                case "xlsm":
                    switch (splt_var[splt_var.Length - 1])
                    {
                        case "DBF":
                            var_table = mah_class.GetYourData_DBF(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsx":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        case "xlsm":
                            var_table = mah_class.ConvertExcelToDataTable(var_name);
                            var_table = remove_andelse(var_table);
                            mah_class.ExportToExcelFile(var_table, splt_var[0] + "_full.xlsx");
                            obj_table = mah_class.ConvertExcelToDataTable(file_name, "ObjektIBild");
                            compare_table(var_table, obj_table, splt_obj[0], splt_var[0]);
                            break;
                        default:
                            break;
                    }
                    break;
                default:
                    break;
            }
            mah_class.ExportToExcelFile(missing_tbl, splt_var[0] + "_missing.xlsx");
            obj_table = null;
            var_table = null;
            mah_class = null;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog choose_file = new OpenFileDialog();

            if (choose_file.ShowDialog() == DialogResult.OK)
            {
                this.textBox2.Text = choose_file.FileName;
            }
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            
        }
    }
}
