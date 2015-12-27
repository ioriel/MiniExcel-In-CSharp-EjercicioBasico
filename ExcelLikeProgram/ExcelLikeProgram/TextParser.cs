using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace ExcelLikeProgram
{

   
    //esta clase debe analizar y dividir la entrada en pedazos pequeños
    //
    /// <summary>
    /// analizador de texto
    /// </summary>
    class TextParser
    {
        public delegate void EventHandlerOnTextChange( object _sender , string _input);

        public EventHandlerOnTextChange OnTextChange;
        
        public List<string> Functions;
        public List<string> Operators;
        public List<string> Operands;

        public Queue<string> FullOperation;


        private string input;
        public string Input 
        {
            get
            {
                return this.input;
            } 
            set
            {

            }
        }

        //constructor
        public TextParser()
        {
            this.OnTextChange = null;
        }

        //analiza la formula
        public bool Parse(string _input)
        {
            bool flag = false;
            // operando operador operando operador operando operador

            //hallar la primera operacion analizando los parentesis -tokens
            //analizar profundidad de la operacion((a+b) + e * (g/h) - 1 (((a+b)+(b-a))))) 
            //
            int levels=0;
            int closedOps=0;
            //analizar consistencia de la formula -- Si tiene  sentido o si tiene algun error antes de

            //verificar consistencia de simbolos de agrupacion

            //verificar consistencia de operadores binarios

            //verificar consistencia de operadores unarios

            //verificar consistencia de parametros de operacions (operandos)

            foreach (char val in _input)
            {
                //buscar niveles de separacion de operaciones con ()

                if (val == '(')
                {
                    levels++;
                }
                if (val == ')')
                {
                    closedOps++;
                }
            }

            //si no hay niveles de jerarquia analizar oepradores

            return flag;
        }

        
        //ejecuta la formula
        public void Execute()
        { 
        
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
