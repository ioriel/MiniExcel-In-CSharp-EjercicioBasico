﻿using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelLikeProgram
{

   
    //esta clase debe analizar y dividir la entrada en pedazos pequeños
    //
    //

    /*
     Expr ::= 
      Expr + Expr
    | Expr - Expr   
    | Expr / Expr       // real division
    | Expr div Expr     // integer division
    | Expr % Expr       // real modulo
    | Expr mod Expr     // integer modulo
    | Expr and Expr     
    | Expr or Expr
    | Expr < Expr 
    | Expr <= Expr 
    | Expr > Expr 
    | Expr >= Expr 
    | Expr == Expr 
    | Expr != Expr 
    | - Expr            
    | not Expr 		
    | IDENTIFIER 				// variable access
    | IDENTIFIER(Expr, Expr, ...) 		// function call 	
    | Expr . IDENTIFIER 			// member variable		
    | Expr . IDENTIFIER(Expr, Expr, ...) 	// member function	
    | Expr .. IDENTIFIER(Expr, Expr, ...) 	// member function, same as single dot
                                            // but returns the receiver type
    | new IDENTIFIER(Expr, Expr, ...) 		// constructor call
    | destroy Expr                  // destroy object
    | Expr castTo Type				// casting
    | Expr instanceof Type			// instance checking	
    | begin
        Statements
      end // statement expr
    | (param1, param2, ...) -> Expr  // anonymous function
    | (Expr)                        // parantheses
     */
    /// <summary>
    /// analizador de texto
    /// </summary>
    class TextParser
    {
        public delegate void EventHandlerOnTextChange( object _sender , string _input);
        private enum Operaciones{Suma , Resta ,Multiplica ,Divide ,Potencia , Undef};
        public enum OperationState { Correcta ,Pendiente, ErrorParametro , ErrorOperador ,ErrorSintaxis,ErrorDesconocido}
        public EventHandlerOnTextChange OnTextChange;
        
        public List<string> Functions;
        public List<string> Operators;
        public List<string> Operands;

        public Queue<string> FullOperation;

        //guarda el estado de la operacion
        private OperationState estadoOperacion;
        public OperationState EstadoOperacion 
        {
            get { return this.estadoOperacion; }
        }

        private Dictionary<string, CellData> inputCells;
        public Dictionary<string,CellData> InputCells 
        {
            set { this.inputCells = value; } 
        }

        private string input;
        public string Input 
        {
            get
            {
                return this.input;
            } 
        }

        private double result;
        public double Result { get { return this.result; } }

        //constructor
        public TextParser(Dictionary<string, CellData> _cells)
        {
            this.OnTextChange = null;
            this.inputCells = _cells;
            this.estadoOperacion = OperationState.Pendiente;
        }

        //:( convierte el nombre de la celda Formato A2..z100 etc en coordenadas de la grilla
        private string ConverCellNameToKey(string _name)
        {
            string ColId = _name.Substring(0, 1);
            string rowId = _name.Substring(1, _name.Length-1);

            int valColId = Convert.ToChar(ColId)-65;
            int valRowId = int.Parse(rowId)-1;
            return valRowId.ToString()+";"+valColId.ToString() ;
        }


        private double Eval(double _expA, double _expB, Operaciones _operacion)
        {
            double result = 0.0;

            switch (_operacion)
            {
                case Operaciones.Suma:
                    result = _expA + _expB;
                    break;
                case Operaciones.Resta:
                    result = _expA - _expB;
                    break;
                case Operaciones.Multiplica:
                    result = _expA * _expB;
                    break;
                case Operaciones.Divide:
                    if (_expB != 0)
                        result = _expA / _expB;
                    else
                        MessageBox.Show("No se puede dividir entre 0");
                    break;
                case Operaciones.Potencia:
                    result = Math.Pow(_expA , _expB);
                    break;
                case Operaciones.Undef:
                    break;
                default:
                    break;
            }

            return result;
        }


        private Operaciones GetOperation(string _operator)
        {
            Operaciones operacion = Operaciones.Undef;
            switch (_operator)
            {
                case "+":
                    operacion = Operaciones.Suma;
                    break;
                case "-":
                    operacion = Operaciones.Resta;
                    break;
                case "/":
                    operacion = Operaciones.Divide;
                    break;
                case "*":
                    operacion = Operaciones.Multiplica;
                    break;
                case "^":
                    operacion = Operaciones.Potencia;
                    break;
                default:
                    break;
            }
            return operacion;
        }


        private string PeekOperator(string _input)
        { 
            string operador = "null";

            foreach (char ch in _input)
            {
                if (ch == '+' || ch == '-' || ch == '*' || ch == '/' || ch == '^')
                {
                    operador = ch.ToString();
                }
            }

            return operador;
        }
        //analiza la formula
        public bool Parse(string _input)
        {
            this.input = _input;
            bool flag = false;
            bool isSingleReferenceOp;

            //analizar consistencia de la formula -- Si tiene  sentido o si tiene algun error antes de

            //verificar consistencia de simbolos de agrupacion

            //verificar consistencia de operadores binarios

            //verificar consistencia de operadores unarios

            //verificar consistencia de parametros de operacions (operandos)

            Regex regex = new Regex("[a-zA-Z]{1}[0-9]{1,3}"); //busca patrones de nombre de celda en la formula

            MatchCollection replaceables = regex.Matches(this.Input);
            string replacedWithValues="";

            foreach (Match item in replaceables)
            {
                //string replacement = item.Value;
                string cellKey = ConverCellNameToKey(item.Value);
                string replacement= "";

                if (this.inputCells[cellKey].CurrentValue != null)
                {
                    replacement = this.inputCells[cellKey].CurrentValue.ToString();
                }
                else
                    replacement = "0";
                //reemplazar nombres de celdas por valores de celdas en la formula
                replacedWithValues = regex.Replace(this.Input, replacement);
                //buscar todos los valores de la formula
            }

            //verificar si es operacion de redireccionamiento de celda y terminar parseo

            if (this.singleReferenceOperation())
            {
                this.input = replacedWithValues;

                this.estadoOperacion = OperationState.Correcta;
                this.result = Double.Parse(this.Input.Substring(1,this.Input.Length-1));
                return true;
            }

            if(replaceables.Count > 0)
                this.input = replacedWithValues;//cadena con vlaores de celdas

            

            //regex 1 operacion binaria
            //^[=]+[(]?[0-9.0-9]+[\+]?[\-]?[\*]?[\/]?[\^]?[0-9.0-9]+[)]?
            //verificar si tiene forma de operacion
            Regex regOperacionBinaria = new Regex(@"^[=]+[(]?[0-9.0-9]+[\+]?[\-]?[\*]?[\/]?[\^]?[0-9.0-9]+[)]?");
            //->>> ojo no reconoce numeros negativos ... falta implementar
            if (regOperacionBinaria.IsMatch(this.Input))
            {
                //analiza 2 expresiones y un operador
                MatchCollection operandos = Regex.Matches(this.Input, @"[0-9.0-9]+");
                //Match operador = Regex.Match(this.Input, @"[\+]?[\-]?[\*]?[\/]?[\^]");//[\+]{0,1}[\-]{0,1}[\*]{0,1}[\/]{0,1}
                //Match operador = Regex.Match(this.Input, @"[+]*[-]*[*]*[/]*");
                string operador = this.PeekOperator(this.input);
                double operandoA;
                double operandoB;

                if (operandos.Count == 2)
                {
                    operandoA = Double.Parse(operandos[0].Value);
                    operandoB = Double.Parse(operandos[1].Value);

                    //MessageBox.Show(operador);
                    if (operador.Length > 0)
                    {
                        this.result = this.Eval(operandoA, operandoB, GetOperation(operador));
                        this.estadoOperacion = OperationState.Correcta;
                        flag = true;
                    }
                    else
                        this.estadoOperacion = OperationState.ErrorOperador;
                }
                else
                    this.estadoOperacion = OperationState.ErrorParametro;
            }
            else
                this.estadoOperacion = OperationState.ErrorSintaxis;

            //^[=]?[(]*[0-9.0-9]+[\+]?[\-]?[\*]?[\/]?[\^]?[0-9.0-9]+[)]*
            return flag;
        }

        //verifica si es una operacion de vlaidacion simpe
        private bool singleReferenceOperation()
        {
            Match res = Regex.Match(this.input, "^[=]+[a-zA-Z]{1}[0-9]{1,3}$"); //regex de valor de celda
            if (res.Length ==3)
            {
                return true;
            }
            return false;
        }

        

        private int GetOperatorPrecedence(String token)
        {
            if (token == "+" || token == "-")
            {
                return 1;
            }
            else if (token == "*" || token == "/")
            {
                return 2;
            }
            else if (token == "^")
            {
                return 3;
            }
            else
            {
                throw new ArgumentException("Invalid token");
            }
        }


    }
}
