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

        private enum FONT_PROP { Bold , Italic , Underlined , changeSize};
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.InitDataGRid();
            this.initComboBox();
        }

        private void initComboBox()
        {
            foreach (FontFamily font in System.Drawing.FontFamily.Families)
            {
                this.comboBox1.Items.Add(font.Name);
            }

            this.comboBox1.SelectedItem = "Microsoft Sans Serif";
            
        }

        private Dictionary<string, CellData> Cells = new Dictionary<string, CellData>();//key = rowIndex+ColIndex , value = CellData definida en otra clase para guardar varios datos de la celda
        private bool editingFormula;
        private bool editingCell;
       // private bool autoUpdatingCellValue;
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
        private bool runOnce;//solo reistrara el evento una vez y no uno cada vez q se seleccione un control
        private void dataGridView1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            if (!this.runOnce)
            {
                e.Control.KeyPress += new KeyPressEventHandler(Control_KeyPress);
                this.runOnce = true;
            }
               
        }

        //metodo delegado de evento de presion de tecla se ejecuta al pulsar una tecla al tener activada una celda en modo edicion
        private void Control_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (/*!this.autoUpdatingCellValue &&*/ this.editingCell)
            {
                int column = dataGridView1.CurrentCellAddress.X;
                int row = dataGridView1.CurrentCellAddress.Y;

                if (e.KeyChar == '=' && !this.editingFormula)
                {
                    this.editingFormula = true;
                }

                if (!char.IsControl(e.KeyChar)/*&& (char.IsLetterOrDigit(e.KeyChar) || e.KeyChar == '+' || e.KeyChar == '-' || e.KeyChar == '/' || e.KeyChar == '*' || e.KeyChar == '^' || e.KeyChar == '(' || e.KeyChar == ')')*/)
                {
                //    //Get the column and row position of the selected cell
                    this.textBox1.Text += char.ToString(e.KeyChar); // agregamos el vlaor del char al textbox

                //    //this.dataGridView1.Rows[row].Cells[column].Style.BackColor = Color.Red;
                //    this.textBox1.Text += char.ToString(e.KeyChar); // agregamos el vlaor del char al textbox
                }
                else if (e.KeyChar == 8)
                {

                    if(!String.IsNullOrEmpty(this.textBox1.Text))
                        this.textBox1.Text = this.textBox1.Text.Substring(0, this.textBox1.Text.Length - 1);

                }
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
        private TextParser ParseFormula(string _formula)
        {
            TextParser parser=null;

            if (!string.IsNullOrEmpty(_formula))
            {
                if (_formula.Substring(0, 1) == "=")
                {
                    parser = new TextParser(this.Cells);

                    parser.Parse(this.textBox1.Text);
                }
            }
           
            return parser;
        }

        //se ejecuta al terminar de editar una celda o salir de modo edicion
        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            //guardar el valor actual de la celda
            string currCellId = this.GetCellKey(e.RowIndex, e.ColumnIndex);
            object currCellValue = this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;


            CellData myCell = this.Cells[currCellId];


            if (currCellValue != null)
            {
                myCell.CurrentFormula = currCellValue.ToString();
                myCell.CurrentValue = currCellValue;
               // MessageBox.Show(myCell.CurrentValue.ToString());
                //procesar si la formula puede ser analizada para obetener un resultado
            }
            

            //reemplazar ese resultado por el valor actual de la celda
            TextParser parsed = this.ParseFormula(myCell.CurrentFormula);

            if (parsed != null)
            {
                if (parsed.EstadoOperacion == TextParser.OperationState.Correcta)
                {
                    double res = parsed.Result;
                    myCell.CurrentValue = res;
                }
                else if (parsed.EstadoOperacion == TextParser.OperationState.ErrorOperador)
                    myCell.CurrentValue = "#Operador Invalido!";
                else if (parsed.EstadoOperacion == TextParser.OperationState.ErrorParametro)
                    myCell.CurrentValue = "#Parametros Incorrectos";
                else if (parsed.EstadoOperacion == TextParser.OperationState.ErrorSintaxis)
                    myCell.CurrentValue = "#Error de Sintaxis";
                else if (parsed.EstadoOperacion == TextParser.OperationState.ErrorDesconocido)
                    myCell.CurrentValue = "#Error Desconocido";

                //myCell.CurrentValue = "#error";
            }
                
                
            this.dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = myCell.CurrentValue ;

            //this.autoUpdatingCellValue = false;
            this.editingCell = false;
            this.editingFormula = false;
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
              
            this.editingCell = true;
           // this.autoUpdatingCellValue = false;

        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show(":D");
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            if (e.RowIndex > 0 && e.ColumnIndex > 0)
            {
                this.textBox1.Clear();
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
            
            if (!this.editingCell)
            {
                int col = this.dataGridView1.CurrentCellAddress.X;
                int row = this.dataGridView1.CurrentCellAddress.Y;
                string cellId = this.GetCellKey(row, col);
                CellData myCell = this.Cells[cellId];
                if (myCell.CurrentValue == null)
                {
                    if(e.KeyChar!= 8)
                        myCell.CurrentValue = e.KeyChar;
                }
                else
                {
                    if (e.KeyChar != 8)
                    {
                        string currVal = myCell.CurrentValue.ToString() + char.ToString(e.KeyChar);
                        myCell.CurrentValue = currVal;
                        this.dataGridView1.Rows[row].Cells[col].Value = myCell.CurrentValue;
                    }
                    
                }
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

        #region apariencia de texto en celda

        private void CambiarAparienciadeCeldasSeleccionadas(FONT_PROP _prop , bool _state)
        {
            if (this.dataGridView1.SelectedCells.Count > 0)
            {
                foreach (DataGridViewCell cell in this.dataGridView1.SelectedCells)
                {
                    int col = this.dataGridView1.CurrentCellAddress.X;//obtenemos coordenadas de celda seleccionada
                    int row = this.dataGridView1.CurrentCellAddress.Y;

                    //guardar el vlaor en celldata
                    //falta validar el estado actual de la fuente
                    //

                    //cambiar el valor de la fuente de la celda
                    //cell.Style.Font = _newFont;
                    switch (_prop)
	                {
                        case FONT_PROP.Bold:
                            {
                                Font nf;

                                if (_state)
                                {
                                    

                                    if (this.chkCursiva.Checked)
                                        if (this.chkSubrayado.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold);

                                }
                                else
                                {
                                    if (this.chkCursiva.Checked)
                                        if (this.chkSubrayado.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular | FontStyle.Italic | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular | FontStyle.Italic);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular);
                                }

                                cell.Style.Font = nf;
                            }
                            break;
                        case FONT_PROP.Italic:
                            {
                                Font nf;

                                if (_state)
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkSubrayado.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Italic);

                                }
                                else
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkSubrayado.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Regular);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular);
                                }

                                cell.Style.Font = nf;
                            }
                            break;
                        case FONT_PROP.Underlined:
                            {
                                Font nf;

                                if (_state)
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkCursiva.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Underline);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Underline);

                                }
                                else
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkCursiva.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Regular | FontStyle.Italic);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Regular);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular);
                                }

                                cell.Style.Font = nf;
                            }
                            break;
                        case FONT_PROP.changeSize:
                            {
                                Font nf;
                                _state = this.chkSubrayado.Checked;
                                if (_state)
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkCursiva.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Italic | FontStyle.Underline);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Underline);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Underline);

                                }
                                else
                                {
                                    if (this.chkNegrita.Checked)
                                        if (this.chkCursiva.Checked)
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Regular | FontStyle.Italic);
                                        else
                                            nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Bold | FontStyle.Regular);
                                    else
                                        nf = new System.Drawing.Font(this.comboBox1.SelectedItem.ToString(), (int)this.numericUpDown1.Value, FontStyle.Regular);
                                }


                                cell.Style.Font = nf;
                            }
                            break;
                        default:
                         break;
	                }
                    
                    
                }
            }
        }

        //pone el contenido de la celda en negrita
        private void btnNegrita_Click(object sender, EventArgs e)
        {
            //permite cambiar el tipo de fuente a negrita a toda la seleccion del grid
            CheckBox chk = sender as CheckBox;
            if (chk.Checked)
            {
                this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.Bold, true);
            }
            else
            {
                this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.Bold, false);
            }
           

        }

        private void btnCursiva_Click(object sender, EventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            if (chk.Checked)
            {
                this.CambiarAparienciadeCeldasSeleccionadas( FONT_PROP.Italic, true);
            }
            else
            {
                this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.Italic, false);
            }
            
        }

        private void btnSubrayada_Click(object sender, EventArgs e)
        {
            CheckBox chk = sender as CheckBox;
            if (chk.Checked)
            {
                this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.Underlined, true);
            }
            else
            {
                this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.Underlined, false);
            }
        }

        private void numericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            this.CambiarAparienciadeCeldasSeleccionadas(FONT_PROP.changeSize ,true);
            //this.CambiarAparienciadeCeldasSeleccionadas(new Font("Microsoft Sans Serif", (int)this.numericUpDown1.Value, FontStyle.Regular));
        }

        #endregion

        private void button2_Click(object sender, EventArgs e)
        {
            if (this.colorDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Color color = this.colorDialog1.Color;

                foreach (DataGridViewCell cell in this.dataGridView1.SelectedCells)
                {
                    cell.Style.BackColor = color;
                }
            }
        }

        private void btnCentrar_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in this.dataGridView1.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in this.dataGridView1.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewCell cell in this.dataGridView1.SelectedCells)
            {
                cell.Style.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
        }

        //actualizar info de celda
        private void dataGridView1_CurrentCellChanged(object sender, EventArgs e)
        {
            int ColumnIndex = this.dataGridView1.CurrentCellAddress.X;
            int RowIndex = this.dataGridView1.CurrentCellAddress.Y;

            if (RowIndex > 0 && ColumnIndex > 0)
            {
                this.textBox1.Clear();
                if (this.dataGridView1.Rows[RowIndex].Cells[ColumnIndex].Value != null)
                {
                    object val = this.dataGridView1.Rows[RowIndex].Cells[ColumnIndex].Value;

                    string currCell = this.GetCellKey(RowIndex, ColumnIndex);
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



    }
}
