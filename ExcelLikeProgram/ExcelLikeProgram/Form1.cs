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

        private Dictionary<int, string> Cols = new Dictionary<int, string>();
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

            this.dataGridView1.RowHeadersWidth = 60;
        }

        
        

        //obtiene el vlaor de la celda dada en formato A1 , b2 , etc
        private object GetCellValue(string _cell)
        {
            string Col = _cell.Substring(0, 1);
            string row = _cell.Substring(1, 1);

            int asciiVal = Convert.ToChar(Col.ToUpper());
            int colIndex = asciiVal;
            int rowIndex = int.Parse(row);

            object val = null;

            if (colIndex <= this.dataGridView1.ColumnCount && rowIndex <= this.dataGridView1.RowCount)
            {
                val = this.dataGridView1.Rows[rowIndex].Cells[colIndex].Value; 
            }
            return val;
        }

        //este metodo analiza la cadena apara averiguar si es una formula
        private bool CanParseFormula(string _formula)
        {
            bool flag = false;
            //analizar si la cadena enviada es una fórmula

            //validar si empieza con "="

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
            int rowIndex =  e.RowIndex+1; //utilizar + 1 para localizar el dato
            int colindex = e.ColumnIndex;
            
            //guardar el valor actual de la celda
            string currCell = this.Cols[colindex] + rowIndex.ToString();//esta parti utiliza un indice aumentado para eliminar el indice 0
            object currCellValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            bool esFormula = false;

            if (this.CanParseFormula(currCell)) // valida si el valor actual de la celda es una formula
                esFormula = true;

            if (this.CellFormulaContainer.ContainsKey(this.Cols[colindex]) && esFormula)//validar si es una formula para guardarla
            {
                this.CellFormulaContainer[currCell] = currCellValue.ToString();
                //this.CellValueContainer[currCell] = currCellValue;
            }
            else if (!this.CellFormulaContainer.ContainsKey(this.Cols[colindex]) && esFormula)
            {
                this.CellFormulaContainer.Add(currCell, currCellValue.ToString());
            }

            

            //procesar si la formula puede ser analizada para obetener un resultado

            //reemplazar ese resultado por el valor actual de la celda
            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = currCell;
        }

        //metodo para ejecucion de operacion en rango de columnas

        //metodo para ejecucion de operacion en rango de filas


        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            int rowIndex = e.RowIndex + 1; //utilizar + 1 para localizar el dato
            int colindex = e.ColumnIndex;

            //guardar el valor actual de la celda
            string currCell = this.Cols[colindex] + rowIndex.ToString();//esta parti utiliza un indice aumentado para eliminar el indice 0
            object currCellValue;
            //reemplazar contenido por el vlaor de la formula almacenada
            if (this.CellFormulaContainer.ContainsKey(this.Cols[colindex]))
            {
                this.CellValueContainer[currCell] = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (object)this.CellFormulaContainer[currCell];
            }
        }
    }
}
