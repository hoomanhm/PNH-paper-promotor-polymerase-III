using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Final2
{

        public class Amino
        {

            public Amino()
            {
                codon = new List<Codon>();
            }
            public string AminoName;
            public double N;
            public double X_Max;
            public double RSCU_Max;
            /// <summary>
            /// Denuminator of RSCU
            /// </summary>
            public double Frequence_Amino;
            public List<Codon> codon;


        }
        public class Codon
        {
            public string Name;
            public double RSCU;
            public double W;
            public Amino Parent;
            public double Frequence_Codon;
        }

    
}
