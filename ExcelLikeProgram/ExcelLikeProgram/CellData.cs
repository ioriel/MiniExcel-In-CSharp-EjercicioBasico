using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelLikeProgram
{
    //guarda algunos datos relacionados a la celda 
    class CellData
    {
        private string id; //id de celda
        public string Id { get { return this.id; } }
        private string name;
        public string Name { get { return this.name; } }
        private string currentValue; // valor actual
        public string CurrentValue 
        {
            get { return this.currentValue; }
            set { this.currentValue = value; }
        }
        private bool bold;
        public bool IsBold { get { return this.bold; } set { this.bold = value; } }
        private bool sub;
        public bool IsSub { get { return this.sub; } set { this.sub = value; } }
        private bool curs;
        public bool IsCursive { get { return this.curs; } set { this.curs = value; } }

        //corrdenadas en la grilla
        private int x;
        public int X { get { return this.x; } set { this.x = value; } }
        private int y;
        public int Y { get { return this.y; } set { this.y = value; } }

        //constructor de la clase
        public CellData(string _cellId)
        {
            this.id = _cellId;
        }

    }
}
