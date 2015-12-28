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

        private Dictionary<string, CellData> Cells = new Dictionary<string, CellData>();//key = rowIndex+ColIndex , value = CellData definida en otra clase para guardar varios datos de la celda
        private bool editingFormula;
        private void InitDataGRid()
        {
            //inicializar Columnas
            char ColId = (char)65;
            int charS = 65;
            for (int i = 0; i < 26; i++)
            {
                ColId = (char)(charS + i);
                this.dataGridView1.Columns.Add("col"+ColId.ToString(),ColId.ToString());
                this.dataGridView1.Columns[i].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
               // this.Cols.Add(i,ColId.ToString());
            }

            //inicializar Filas
            for (int j = 0; j < 100; j++)
            {
                this.dataGridView1.Rows.Add();
                this.dataGridView1.Rows[j].HeaderCell.Value = (j+1).ToString();
                
            }

            //inicializar diccionario de celdas e inicializar las celdas
            string record = "#";
            int startngChar = 65;
            for (int k = 0; k < this.dataGridView1.RowCount; k++)
            {
                for (int l = 0; l < this.dataGridView1.ColumnCount; l++)
                {
                    string id = (k.ToString()+";"+l.ToString()); //coordenadas
                    string cellName = ((char)(startngChar + l)).ToString() + (k + 1);//inicializar valores de formato A-Z//nombre en cell Data
                    record += ","+id;

                    try
                    {
                        CellData cell = new CellData(id,cellName);
                        cell.X = k;
                        cell.Y = l;
                        this.Cells.Add(id, cell);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Error en =>Key:"+id+";Val:"+cellName);
                    }
                    

                }
            }

            //MessageBox.Show(record);
            this.dataGridView1.RowHeadersWidth = 60;

            //controlador de evento de edicion de control
            dataGridView1.EditingControlShowing += new DataGridViewEditingControlShowingEventHandler(dataGridView1_EditingControlShowing);
        }

        //metodo delegado de vento en control
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
        }

        //metodo delegado de evento de presion de tecla se ejecuta al pulsar una tecla al tener activada una celda en modo edicion
        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            int column = dataGridView1.CurrentCellAddress.X;
            int row = dataGridView1.CurrentCellAddress.Y;

            if (e.KeyChar == '=' && !this.editingFormula)
            {
                this.editingFormula = true;
            }


            if (this.editingFormula && !char.IsControl(e.KeyChar)/*&& (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '+' || e.KeyChar == '-' || e.KeyChar == '/' || e.KeyChar == '*' || e.KeyChar == '^' || e.KeyChar == '(' || e.KeyChar == ')')*/)
            {

                string cellValue = Char.ToString(e.KeyChar);
                //Get the column and row position of the selected cell
               

                //this.dataGridView1.Rows[row].Cells[column].Style.BackColor = Color.Red;
                this.textBox1.Text += char.ToString(e.KeyChar); // agregamos el vlaor del char al textbox
            }
            else if(e.KeyChar == 157)
            {
                if (this.dataGridView1.Rows[row].Cells[column].Value != null)
                    this.textBox1.Text = this.dataGridView1.Rows[row].Cells[column].Value.ToString();

            }
        }
        
        

        //obtiene el id de la celda dada en formato A1 , b2 , etc
        private string GetCellKey(int _rowIndex, int _colIndex)
        {
            //string key = _rowIndex.ToString()+";"+_colIndex;
            return _rowIndex.ToString() + ";" + _colIndex.ToString();
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
            TextParser parser = new TextParser(this.Cells);

            parser.Parse(this.textBox1.Text);
            this.editingFormula = false;
            return null;
        }

        //se ejecuta al terminar de editar una celda o salir de modo edicion
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //int rowIndex =  e.RowIndex; //utilizar + 1 para localizar el dato
           // int colindex = e.ColumnIndex;
            
            //guardar el valor actual de la celda
            string currCellId = this.GetCellKey(e.RowIndex, e.ColumnIndex);
            object currCellValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            bool esFormula = false;

            CellData myCell = this.Cells[currCellId];

            if(currCellValue!= null)
                if (this.CanParseFormula(currCellValue.ToString())) // valida si el valor actual de la celda es una formula
                    esFormula = true;

            if (!string.IsNullOrEmpty(myCell.CurrentFormula)  && esFormula)//validar si es una formula para guardarla
            {
                myCell.CurrentFormula = currCellValue.ToString();
            }

            if (currCellValue != null)
            {
                myCell.CurrentFormula = currCellValue.ToString();
                myCell.CurrentValue = currCellValue;
               // MessageBox.Show(myCell.CurrentValue.ToString());
                //procesar si la formula puede ser analizada para obetener un resultado
            }
            

            //reemplazar ese resultado por el valor actual de la celda
            if (esFormula)
            {
                //this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Blue;
            }

            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = myCell.Id ;
        }

        //metodo para ejecucion de operacion en rango de columnas

        //metodo para ejecucion de operacion en rango de filas


        //este metodo se ejecuta cuando se inicia la edicion de una celda 
        //guarda el valur actual de la celda y lo reemplaza por la formula si fuera necesario
        private void dataGridView1_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {

            string currCellId = this.GetCellKey(e.RowIndex, e.ColumnIndex);//esta parte utiliza un indice aumentado para eliminar el indice 0
            CellData myCell = this.Cells[currCellId];
            //reemplazar contenido por el vlaor de la formula almacenada
            
            if (!string.IsNullOrEmpty(myCell.CurrentFormula))
            {
                myCell.CurrentValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
                this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = (object)myCell.CurrentFormula;
                this.textBox1.Text = myCell.CurrentFormula;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show(":D");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex > 0 && e.ColumnIndex > 0)
            {
                if (this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                {
                    object val = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;

                    string currCell = this.GetCellKey(e.RowIndex, e.ColumnIndex);
                    CellData myCell = this.Cells[currCell];
                    if (!string.IsNullOrEmpty(myCell.CurrentFormula))
                    {
                        this.textBox1.Text = myCell.CurrentFormula;
                    }
                    else
                    {
                        //mostrar valor de la celda
                        if (val != null)
                            this.textBox1.Text = val.ToString();
                    }

                }
            }
            
            
        }

        //copia el vlaor del textbox a la celda actual
        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            int col = this.dataGridView1.CurrentCellAddress.X;
            int row = this.dataGridView1.CurrentCellAddress.Y;
            string cellId = this.GetCellKey(row, col);
            CellData myCell = this.Cells[cellId];

            

            if (myCell.CurrentValue == null)
            {
                myCell.CurrentValue = e.KeyChar;
            }
            else
            {
                string currVal = myCell.CurrentValue.ToString() + char.ToString(e.KeyChar);
                myCell.CurrentValue = currVal;
                this.dataGridView1.Rows[row].Cells[col].Value = myCell.CurrentValue;
            }

            
           
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.ParseFormula(this.textBox1.Text);
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.ParseFormula(this.textBox1.Text);
            }
        }
    }
}
