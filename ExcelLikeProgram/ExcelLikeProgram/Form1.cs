using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExcelLikeProgram
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.InitDataGRid();
        }

        private Dictionary<int, string> Cols = new Dictionary<int, string>();//eliminar esto

        private Dictionary<string, string> Cells = new Dictionary<string, CellData>();//key = rowIndex+ColIndex , value = valor en formato A3 por ejemplo
        private Dictionary<string, string> CellFormulaContainer = new Dictionary<string,string>();
        private Dictionary<string, object> CellValueContainer = new Dictionary<string, object>();

        private void InitDataGRid()
        {
            //inicializar celdas
            char ColId = (char)65;
            int charS = 65;
            for (int i = 0; i < 26; i++)
            {
                ColId = (char)(charS + i);
                this.dataGridView1.Columns.Add("col"+ColId.ToString(),ColId.ToString());
                this.dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                this.Cols.Add(i,ColId.ToString());
            }

            //inicializar columnas
            for (int j = 0; j < 100; j++)
            {
                this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[j].HeaderCell.Value = (j+1).ToString();
                
            }

            //inicializar diccionario de celdas
            string record = "#";
            int startngChar = 65;
            for (int k = 0; k < this.dataGridView1.RowCount; k++)
            {
                for (int l = 0; l < this.dataGridView1.ColumnCount; l++)
                {
                    string key = (k.ToString()+";"+l.ToString());
                    string val = ((char)(startngChar + l)).ToString() + (k + 1);//inicializar valores de formato A-Z
                    record += ","+key;

                    try
                    {
                        this.Cells.Add(key, val);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Error en =>Key:"+key+";Val:"+val);
                    }
                    

                }
            }

            //MessageBox.Show(record);
            this.dataGridView1.RowHeadersWidth = 60;

            //
            dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView1_EditingControlShowing);
        }

        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
        }

        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == '=')
            {

                string cellValue = Char.ToString(e.KeyChar);
                //Get the column and row position of the selected cell
                int column = dataGridView1.CurrentCellAddress.X;
                int row = dataGridView1.CurrentCellAddress.Y;

                //this.dataGridView1.Rows[row].Cells[column].Style.BackColor = Color.Red;
                this.textBox1.Text = "[Formula]";
            }
        }
        
        

        //obtiene el id de la celda dada en formato A1 , b2 , etc
        private string GetCellId(int _rowIndex, int _colIndex)
        {
            string key = _rowIndex.ToString()+";"+_colIndex;
            return this.Cells[_rowIndex.ToString() + ";" + _colIndex];
        }

        //este metodo analiza la cadena apara averiguar si es una formula
        private bool CanParseFormula(string _formula)
        {
            bool flag = true;
            //analizar si la cadena enviada es una fórmula
            
            //validar si empieza con "="
            if (_formula.Substring(0, 1) != "=")
                return false;

            //validar si contiene operadores

            //validar si contiene operando válidos

            return flag;
        }

        //analiza y obtiene un resultado de la formula/operacion recibida
        private object ParseFormula( string _formula)
        {
            object val = null;

            return val;
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //int rowIndex =  e.RowIndex; //utilizar + 1 para localizar el dato
           // int colindex = e.ColumnIndex;
            
            //guardar el valor actual de la celda
            string currCellId = this.GetCellId(e.RowIndex, e.ColumnIndex);
            object currCellValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            bool esFormula = false;

            if(currCellValue!= null)
                if (this.CanParseFormula(currCellValue.ToString())) // valida si el valor actual de la celda es una formula
                    esFormula = true;

            if (this.CellFormulaContainer.ContainsKey(currCellId) && esFormula)//validar si es una formula para guardarla
            {
                this.CellFormulaContainer[currCellId] = currCellValue.ToString();
                //this.CellValueContainer[currCell] = currCellValue;
            }
            else if (!this.CellFormulaContainer.ContainsKey(currCellId) && esFormula)
            {
                this.CellFormulaContainer.Add(currCellId, currCellValue.ToString());
            }

            

            //procesar si la formula puede ser analizada para obetener un resultado

            //reemplazar ese resultado por el valor actual de la celda
            if (esFormula)
            {
                //this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Blue;
            }

            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = currCellId;
        }

        //metodo para ejecucion de operacion en rango de columnas

        //metodo para ejecucion de operacion en rango de filas


        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

            string currCellId = this.GetCellId(e.RowIndex, e.ColumnIndex);//esta parte utiliza un indice aumentado para eliminar el indice 0
            //reemplazar contenido por el vlaor de la formula almacenada
            if (this.CellFormulaContainer.ContainsKey(currCellId))
            {
                this.CellValueContainer[currCellId] = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (object)this.CellFormulaContainer[currCellId];
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("ta cambiando loco");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
            {
                object val = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                string currCell = this.GetCellId(e.RowIndex,e.ColumnIndex);

                if (this.CellFormulaContainer.ContainsKey(currCell))
                {
                    this.textBox1.Text = this.CellFormulaContainer[currCell];
                }
                else
                { 
                    //mostrar valor de la celda
                    if(val != null)
                        this.textBox1.Text = val.ToString();
                }
                
            }
            
        }
    }
}
