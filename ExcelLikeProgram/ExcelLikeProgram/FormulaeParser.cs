using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelLikeProgram
{
    class FormulaeParser
    {
        //symbols

        //operators

        //jerarquia

        //words
        private static List<string> OperationWord;

        private static void initOperationWords()
        {
            OperationWord = new List<string>();

            OperationWord.Add("SUMA");
            OperationWord.Add("RESTA");
            OperationWord.Add("MULT");
            OperationWord.Add("DIV");
            OperationWord.Add("MOD");
        }

    }
}
