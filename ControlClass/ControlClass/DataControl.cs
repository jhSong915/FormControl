using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ControlClass
{
    public class DataControl
    {
        public DataSet GetFormDataByName(Form _frm)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("MAIN");
            dt.Rows.Add(dt.NewRow());
            foreach (Control _ctrl in _frm.Controls)
            {
                if (IsControlCheck(_ctrl) == true)
                {
                    dt.Columns.Add(new DataColumn(_ctrl.Name));
                    if (GetControlType(_ctrl) == "TextBox")
                        dt.Rows[0][_ctrl.Name] = _ctrl.Text;
                    else if (GetControlType(_ctrl) == "ComboBox")
                        dt.Rows[0][_ctrl.Name] = string.IsNullOrEmpty(((ComboBox)_ctrl).SelectedValue.ToString()) ? "" : ((ComboBox)_ctrl).SelectedValue.ToString();
                    else if (GetControlType(_ctrl) == "CheckBox")
                        dt.Rows[0][_ctrl.Name] = ((CheckBox)_ctrl).Checked == true ? "Y" : "";
                    else if (GetControlType(_ctrl) == "RadioButton")
                        dt.Rows[0][_ctrl.Name] = ((RadioButton)_ctrl).Checked == true ? "Y" : "";


                }
            }
            ds.Tables.Add(dt);
            return ds;
        }

        public void SetFormDataByName(Form _frm, DataSet _ds)
        {
            foreach (Control _ctrl in _frm.Controls)
            {
                if (IsControlCheck(_ctrl) == true)
                {
                    foreach (DataTable dt in _ds.Tables)
                    {
                        if (dt.Columns.Contains(_ctrl.Name))
                        {
                            if (GetControlType(_ctrl) == "TextBox")
                                _ctrl.Text = dt.Rows[0][_ctrl.Name.ToString()].ToString();
                            else if (GetControlType(_ctrl) == "ComboBox")
                                ((ComboBox)_ctrl).SelectedValue = dt.Rows[0][_ctrl.Name.ToString()].ToString();
                            else if (GetControlType(_ctrl) == "CheckBox")
                                ((CheckBox)_ctrl).Checked = dt.Rows[0][_ctrl.Name.ToString()].ToString() == "Y" ? true : false;
                            else if (GetControlType(_ctrl) == "RadioButton")
                                ((RadioButton)_ctrl).Checked = dt.Rows[0][_ctrl.Name.ToString()].ToString() == "Y" ? true : false;
                        }
                    }
                }
            }
        }

        public DataSet GetFormDataByTag(Form _frm)
        {
            DataSet ds = new DataSet();
            DataTable dt = new DataTable("MAIN");
            dt.NewRow();
            foreach (Control _ctrl in _frm.Controls)
            {
                if (IsControlCheck(_ctrl) == true)
                {
                    dt.Columns.Add(new DataColumn(_ctrl.Tag.ToString()));
                    if (GetControlType(_ctrl) == "TextBox")
                        dt.Rows[0][_ctrl.Tag.ToString()] = _ctrl.Text;
                    else if (GetControlType(_ctrl) == "ComboBox")
                        dt.Rows[0][_ctrl.Tag.ToString()] = string.IsNullOrEmpty(((ComboBox)_ctrl).SelectedValue.ToString()) ? "" : ((ComboBox)_ctrl).SelectedValue.ToString();
                    else if (GetControlType(_ctrl) == "CheckBox")
                        dt.Rows[0][_ctrl.Tag.ToString()] = ((CheckBox)_ctrl).Checked == true ? "Y" : "";
                    else if (GetControlType(_ctrl) == "RadioButton")
                        dt.Rows[0][_ctrl.Tag.ToString()] = ((RadioButton)_ctrl).Checked == true ? "Y" : "";
                }
            }
            ds.Tables.Add(dt);
            return ds;
        }

        public void SetFormDataByTag(Form _frm, DataSet _ds)
        {
            foreach (Control _ctrl in _frm.Controls)
            {
                if (IsControlCheck(_ctrl) == true)
                {
                    foreach (DataTable dt in _ds.Tables)
                    {
                        if (dt.Columns.Contains(_ctrl.Tag.ToString()))
                        {
                            if (GetControlType(_ctrl) == "TextBox")
                                _ctrl.Text = dt.Rows[dt.Rows.Count -1][_ctrl.Tag.ToString()].ToString();
                            else if (GetControlType(_ctrl) == "ComboBox")
                                ((ComboBox)_ctrl).SelectedValue = dt.Rows[dt.Rows.Count - 1][_ctrl.Tag.ToString()].ToString();
                            else if (GetControlType(_ctrl) == "CheckBox")
                                ((CheckBox)_ctrl).Checked = dt.Rows[dt.Rows.Count - 1][_ctrl.Tag.ToString()].ToString() == "Y" ? true : false;
                            else if (GetControlType(_ctrl) == "RadioButton")
                                ((RadioButton)_ctrl).Checked = dt.Rows[0][_ctrl.Tag.ToString()].ToString() == "Y" ? true : false;
                        }
                    }
                }
            }
        }

        public bool IsControlCheck(Control _ctrl)
        {
            if (_ctrl.GetType().ToString() == "System.Windows.Forms.TextBox" || _ctrl.GetType().ToString() == "System.Windows.Forms.ComboBox" ||
                _ctrl.GetType().ToString() == "System.Windows.Forms.CheckBox" || _ctrl.GetType().ToString() == "System.Windows.Forms.RadioButton")
                return true;
            else
                return false;
        }

        public string GetControlType(Control _ctrl)
        {
            string return_value = null;
            if (_ctrl.GetType().ToString().Replace("System.Windows.Forms.", "") == "TextBox" || 
                _ctrl.GetType().ToString().Replace("System.Windows.Forms.", "") == "ComboBox" ||
                _ctrl.GetType().ToString().Replace("System.Windows.Forms.", "") == "CheckBox" || 
                _ctrl.GetType().ToString().Replace("System.Windows.Forms.", "") == "RadioButton")
                return_value = _ctrl.GetType().ToString().Replace("System.Windows.Forms.", "");
                return return_value;
        }
    }
}
