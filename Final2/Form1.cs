using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Microsoft.Office.Interop.Excel;

namespace Final2
{
   
    public partial class Form1 : Form
    {
       
        
        List<Codon> AllCodon = new List<Codon>();
        List<Amino> AllAmino = new List<Amino>();
        List<string> FileList = new List<string>();
        List<string> Result = new List<string>();
        List<string> address = new List<string>();
        StringBuilder strb = new StringBuilder();

        int length;
        
        void Show()
        {
            

            System.Data.DataTable DT = new System.Data.DataTable();
            DT.Columns.Add("Profile File1 And Genes");
            DT.Columns.Add("Answer");

            for (int i = 0; i < Result.Count; i++)
            {
                DataRow Row = DT.NewRow();
                Row[0] = address[i];
                Row[1] = Result[i];
                DT.Rows.Add(Row.ItemArray);
            }


            dataGridView1.DataSource = DT;

        }
        void Search(string selected)
        {
           
            string ch;

            try
            {
                length = selected.Length;
                
                for (int i = 0; i < length; i += 3)
                {
                    ch = selected.ToString().Substring(i, 3);
                    var selectedcodon = AllCodon.Find(s => s.Name == ch);
                    selectedcodon.Frequence_Codon += 1;
                    selectedcodon.Parent.Frequence_Amino += 1;
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
            }

        }
        void Calculation()
        {

            foreach (var amin in AllAmino)
            {

                // update Max
                //X-Max
                amin.X_Max = 0;
                foreach (var item in amin.codon)
                {
                    if (item.Frequence_Codon > item.Parent.X_Max)
                        item.Parent.X_Max = item.Frequence_Codon;
                }

                //RSCU & W
                foreach (var item in amin.codon)
                {

                    if (item.Parent.Frequence_Amino == 0)
                    {
                        item.RSCU = 0; item.W = 0.5;
                    }
                    else
                    {
                        item.RSCU = (item.Frequence_Codon / ((1 / item.Parent.N) * item.Parent.Frequence_Amino));

                        item.W = item.Frequence_Codon / item.Parent.X_Max;
                    }
                    if (item.W == 0)
                        item.W = 0.5;

                }

                //RSCU_Max & W_Max
                amin.RSCU_Max = 0;
                foreach (var codon in amin.codon)
                {

                    if (codon.RSCU > codon.Parent.RSCU_Max)
                        codon.Parent.RSCU_Max = codon.RSCU;

                }

                
            }
         
            double CAIobs = 1;
            double CAI_Max = 1;
            foreach (var item in AllCodon)
            {
                if (item.RSCU != 0)
                {
                    CAIobs *= item.RSCU;
                    CAI_Max *= item.Parent.RSCU_Max;
                }

            }

            double pow = 0;
            pow = 1.0 / Convert.ToDouble((length / 3));
            double div = 0;
            div = CAIobs / CAI_Max;
            double result = 0;
            result = Convert.ToDouble(Math.Pow(div, pow));
            Result.Add(result.ToString());
        }
        void Clean()
        {
            dataGridView1.DataSource = string.Empty;
            richTextBox1.Text = string.Empty;
            richTextBox2.Text = string.Empty;
            button3.Enabled = true;
            FileList.RemoveRange(0, FileList.Count);
            Result.RemoveRange(0, Result.Count);
            address.RemoveRange(0, address.Count);
            length = 0;
            radioButton3.Enabled = true;
            radioButton4.Enabled = true;
            radioButton4.Checked = true;
            radioButton3.Checked = false;
            button2.Enabled = true;
            foreach (var item in AllCodon)
            {
                item.W = 0;
                item.RSCU = 0;
                item.Frequence_Codon = 0;
            }
            foreach (var item in AllAmino)
            {
                item.Frequence_Amino = 0;
                item.RSCU_Max = 0;
                item.X_Max = 0;
            }
        }
        public Form1()
        {
            
            InitializeComponent();

            #region codon
            Codon C1 = new Codon();
            C1.Name = "ttt";
            C1.Frequence_Codon = 0;
            C1.RSCU = 0;
            C1.W = 0;

            Codon C2 = new Codon();
            C2.Name = "ttc";
            C2.Frequence_Codon = 0;
            C2.RSCU = 0;
            C2.W = 0;

            Codon C3 = new Codon();
            C3.Name = "tta";
            C3.Frequence_Codon = 0;
            C3.RSCU = 0;
            C3.W = 0;

            Codon C4 = new Codon();
            C4.Name = "ttg";
            C4.Frequence_Codon = 0;
            C4.RSCU = 0;
            C4.W = 0;

            Codon C5 = new Codon();
            C5.Name = "tct";
            C5.Frequence_Codon = 0;
            C5.RSCU = 0;
            C5.W = 0;

            Codon C6 = new Codon();
            C6.Name = "tcc";
            C6.Frequence_Codon = 0;
            C6.RSCU = 0;
            C6.W = 0;

            Codon C7 = new Codon();
            C7.Name = "tca";
            C7.Frequence_Codon = 0;
            C7.RSCU = 0;
            C7.W = 0;

            Codon C8 = new Codon();
            C8.Name = "tcg";
            C8.Frequence_Codon = 0;
            C8.RSCU = 0;
            C8.W = 0;

            Codon C9 = new Codon();
            C9.Name = "tat";
            C9.Frequence_Codon = 0;
            C9.RSCU = 0;
            C9.W = 0;

            Codon C10 = new Codon();
            C10.Name = "tac";
            C10.Frequence_Codon = 0;
            C10.RSCU = 0;
            C10.W = 0;

            Codon C11 = new Codon();
            C11.Name = "taa";
            C11.Frequence_Codon = 0;
            C11.RSCU = 0;
            C11.W = 0;

            Codon C12 = new Codon();
            C12.Name = "tag";
            C12.Frequence_Codon = 0;
            C12.RSCU = 0;
            C12.W = 0;

            Codon C13 = new Codon();
            C13.Name = "tgt";
            C13.Frequence_Codon = 0;
            C13.RSCU = 0;
            C13.W = 0;

            Codon C14 = new Codon();
            C14.Name = "tgc";
            C14.Frequence_Codon = 0;
            C14.RSCU = 0;
            C14.W = 0;

            Codon C15 = new Codon();
            C15.Name = "tga";
            C15.Frequence_Codon = 0;
            C15.RSCU = 0;
            C15.W = 0;

            Codon C16 = new Codon();
            C16.Name = "tgg";
            C16.Frequence_Codon = 0;
            C16.RSCU = 0;
            C16.W = 0;

            Codon C17 = new Codon();
            C17.Name = "ctt";
            C17.Frequence_Codon = 0;
            C17.RSCU = 0;
            C17.W = 0;

            Codon C18 = new Codon();
            C18.Name = "ctc";
            C18.Frequence_Codon = 0;
            C18.RSCU = 0;
            C18.W = 0;

            Codon C19 = new Codon();
            C19.Name = "cta";
            C19.Frequence_Codon = 0;
            C19.RSCU = 0;
            C19.W = 0;

            Codon C20 = new Codon();
            C20.Name = "ctg";
            C20.Frequence_Codon = 0;
            C20.RSCU = 0;
            C20.W = 0;

            Codon C21 = new Codon();
            C21.Name = "cct";
            C21.Frequence_Codon = 0;
            C21.RSCU = 0;
            C21.W = 0;

            Codon C22 = new Codon();
            C22.Name = "ccc";
            C22.Frequence_Codon = 0;
            C22.RSCU = 0;
            C22.W = 0;

            Codon C23 = new Codon();
            C23.Name = "cca";
            C23.Frequence_Codon = 0;
            C23.RSCU = 0;
            C23.W = 0;

            Codon C24 = new Codon();
            C24.Name = "ccg";
            C24.Frequence_Codon = 0;
            C24.RSCU = 0;
            C24.W = 0;

            Codon C25 = new Codon();
            C25.Name = "cat";
            C25.Frequence_Codon = 0;
            C25.RSCU = 0;
            C25.W = 0;

            Codon C26 = new Codon();
            C26.Name = "cac";
            C26.Frequence_Codon = 0;
            C26.RSCU = 0;
            C26.W = 0;

            Codon C27 = new Codon();
            C27.Name = "caa";
            C27.Frequence_Codon = 0;
            C27.RSCU = 0;
            C27.W = 0;

            Codon C28 = new Codon();
            C28.Name = "cag";
            C28.Frequence_Codon = 0;
            C28.RSCU = 0;
            C28.W = 0;

            Codon C29 = new Codon();
            C29.Name = "cgt";
            C29.Frequence_Codon = 0;
            C29.RSCU = 0;
            C29.W = 0;

            Codon C30 = new Codon();
            C30.Name = "cgc";
            C30.Frequence_Codon = 0;
            C30.RSCU = 0;
            C30.W = 0;

            Codon C31 = new Codon();
            C31.Name = "cga";
            C31.Frequence_Codon = 0;
            C31.RSCU = 0;
            C31.W = 0;

            Codon C32 = new Codon();
            C32.Name = "cgg";
            C32.Frequence_Codon = 0;
            C32.RSCU = 0;
            C32.W = 0;
            Codon C33 = new Codon();
            C33.Name = "att";
            C33.Frequence_Codon = 0;
            C33.RSCU = 0;
            C33.W = 0;

            Codon C34 = new Codon();
            C34.Name = "atc";
            C34.Frequence_Codon = 0;
            C34.RSCU = 0;
            C34.W = 0;

            Codon C35 = new Codon();
            C35.Name = "ata";
            C35.Frequence_Codon = 0;
            C35.RSCU = 0;
            C35.W = 0;

            Codon C36 = new Codon();
            C36.Name = "atg";
            C36.Frequence_Codon = 0;
            C36.RSCU = 0;
            C36.W = 0; 

            Codon C37 = new Codon();
            C37.Name = "act";
            C37.Frequence_Codon = 0;
            C37.RSCU = 0;
            C37.W = 0;

            Codon C38 = new Codon();
            C38.Name = "acc";
            C38.Frequence_Codon = 0;
            C38.RSCU = 0;
            C38.W = 0;

            Codon C39 = new Codon();
            C39.Name = "aca";
            C39.Frequence_Codon = 0;
            C39.RSCU = 0;
            C39.W = 0;

            Codon C40 = new Codon();
            C40.Name = "acg";
            C40.Frequence_Codon = 0;
            C40.RSCU = 0;
            C40.W = 0;
            Codon C41 = new Codon();
            C41.Name = "aat";
            C41.Frequence_Codon = 0;
            C41.RSCU = 0;
            C41.W = 0;

            Codon C42 = new Codon();
            C42.Name = "aac";
            C42.Frequence_Codon = 0;
            C42.RSCU = 0;
            C42.W = 0;

            Codon C43 = new Codon();
            C43.Name = "aaa";
            C43.Frequence_Codon = 0;
            C43.RSCU = 0;
            C43.W = 0;

            Codon C44 = new Codon();
            C44.Name = "aag";
            C44.Frequence_Codon = 0;
            C44.RSCU = 0;
            C44.W = 0;

            Codon C45 = new Codon();
            C45.Name = "agt";
            C45.Frequence_Codon = 0;
            C45.RSCU = 0;
            C45.W = 0;

            Codon C46 = new Codon();
            C46.Name = "agc";
            C46.Frequence_Codon = 0;
            C46.RSCU = 0;
            C46.W = 0;

            Codon C47 = new Codon();
            C47.Name = "aga";
            C47.Frequence_Codon = 0;
            C47.RSCU = 0;
            C47.W = 0;

            Codon C48 = new Codon();
            C48.Name = "agg";
            C48.Frequence_Codon = 0;
            C48.RSCU = 0;
            C48.RSCU = 0;
            C48.W = 0;

            Codon C49 = new Codon();
            C49.Name = "gtt";
            C49.Frequence_Codon = 0;
            C49.RSCU = 0;
            C49.W = 0;

            Codon C50 = new Codon();
            C50.Name = "gtc";
            C50.Frequence_Codon = 0;
            C50.RSCU = 0;
            C50.W = 0;

            Codon C51 = new Codon();
            C51.Name = "gta";
            C51.Frequence_Codon = 0;
            C51.RSCU = 0;
            C51.W = 0;

            Codon C52 = new Codon();
            C52.Name = "gtg";
            C52.Frequence_Codon = 0;
            C52.RSCU = 0;
            C52.W = 0;

            Codon C53 = new Codon();
            C53.Name = "gct";
            C53.Frequence_Codon = 0;
            C53.RSCU = 0;
            C53.W = 0;

            Codon C54 = new Codon();
            C54.Name = "gcc";
            C54.Frequence_Codon = 0;
            C54.RSCU = 0;
            C54.W = 0;

            Codon C55 = new Codon();
            C55.Name = "gca";
            C55.Frequence_Codon = 0;
            C55.RSCU = 0;
            C55.W = 0;

            Codon C56 = new Codon();
            C56.Name = "gcg";
            C56.Frequence_Codon = 0;
            C56.RSCU = 0;
            C56.W = 0;

            Codon C57 = new Codon();
            C57.Name = "gat";
            C57.Frequence_Codon = 0;
            C57.RSCU = 0;
            C57.W = 0;

            Codon C58 = new Codon();
            C58.Name = "gac";
            C58.Frequence_Codon = 0;
            C58.RSCU = 0;
            C58.W = 0;

            Codon C59 = new Codon();
            C59.Name = "gaa";
            C59.Frequence_Codon = 0;
            C59.RSCU = 0;
            C59.W = 0;

            Codon C60 = new Codon();
            C60.Name = "gag";
            C60.Frequence_Codon = 0;
            C60.RSCU = 0;
            C60.W = 0;

            Codon C61 = new Codon();
            C61.Name = "ggt";
            C61.Frequence_Codon = 0;
            C61.RSCU = 0;
            C61.W = 0;

            Codon C62 = new Codon();
            C62.Name = "ggc";
            C62.Frequence_Codon = 0;
            C62.RSCU = 0;
            C62.W = 0;

            Codon C63 = new Codon();
            C63.Name = "gga";
            C63.Frequence_Codon = 0;
            C63.RSCU = 0;
            C63.W = 0;

            Codon C64 = new Codon();
            C64.Name = "ggg";
            C64.Frequence_Codon = 0;
            C64.RSCU = 0;
            C64.W = 0;
            #endregion

            #region Phe
            Amino A1 = new Amino();
            A1.AminoName = "Phe";
            A1.Frequence_Amino = 0;
            A1.N = 2;
            C1.Parent = A1;
            C2.Parent = A1;
            A1.codon.Add(C1);
            A1.codon.Add(C2);
            AllCodon.Add(C1);
            AllCodon.Add(C2);
            AllAmino.Add(A1);
            #endregion
            #region Leu
            Amino A2 = new Amino();
            A2.AminoName = "Leu";
            A2.Frequence_Amino = 0;
            A2.N = 6;
            A2.codon.Add(C3);
            A2.codon.Add(C4);
            A2.codon.Add(C17);
            A2.codon.Add(C18);
            A2.codon.Add(C19);
            A2.codon.Add(C20);
            C3.Parent = A2;
            C4.Parent = A2;
            C17.Parent = A2;
            C18.Parent = A2;
            C19.Parent = A2;
            C20.Parent = A2;
            AllCodon.Add(C3);
            AllCodon.Add(C4);
            AllCodon.Add(C17);
            AllCodon.Add(C18);
            AllCodon.Add(C19);
            AllCodon.Add(C20);
            AllAmino.Add(A2);
            #endregion
            #region Ser
            Amino A3 = new Amino();
            A3.AminoName = "Ser";
            A3.Frequence_Amino = 0;
            A3.N = 6;
            A3.codon.Add(C5);
            A3.codon.Add(C6);
            A3.codon.Add(C7);
            A3.codon.Add(C8);
            A3.codon.Add(C45);
            A3.codon.Add(C46);
            C7.Parent = A3;
            C8.Parent = A3;
            C5.Parent = A3;
            C6.Parent = A3;
            C45.Parent = A3;
            C46.Parent = A3;
            AllCodon.Add(C5);
            AllCodon.Add(C6);
            AllCodon.Add(C7);
            AllCodon.Add(C8);
            AllCodon.Add(C45);
            AllCodon.Add(C46);
            AllAmino.Add(A3);
            #endregion
            #region Tyr
            Amino A4 = new Amino();
            A4.AminoName = "Tyr";
            A4.Frequence_Amino = 0;
            A4.N = 2;
            A4.codon.Add(C9);
            A4.codon.Add(C10);
            C9.Parent = A4;
            C10.Parent = A4;
            AllCodon.Add(C9);
            AllCodon.Add(C10);
            AllAmino.Add(A4);
            #endregion
            #region Cys
            Amino A5 = new Amino();
            A5.AminoName = "Cys";
            A5.Frequence_Amino = 0;
            A5.N = 2;
            A5.codon.Add(C13);
            A5.codon.Add(C14);
            C13.Parent = A5;
            C14.Parent = A5;
            AllCodon.Add(C13);
            AllCodon.Add(C14);
            AllAmino.Add(A5);
            #endregion
            #region Trp
            Amino A6 = new Amino();
            A6.AminoName = "Trp";
            A6.Frequence_Amino = 0;
            A6.N = 1;
            A6.codon.Add(C16);
            C16.Parent = A6;
            AllCodon.Add(C16);
            AllAmino.Add(A6);
            #endregion
            #region Pro
            Amino A7 = new Amino();
            A7.AminoName = "Pro";
            A7.Frequence_Amino = 0;
            A7.N = 4;
            A7.codon.Add(C21);
            A7.codon.Add(C22);
            A7.codon.Add(C23);
            A7.codon.Add(C24);
            C21.Parent = A7;
            C22.Parent = A7;
            C23.Parent = A7;
            C24.Parent = A7;
            AllCodon.Add(C21);
            AllCodon.Add(C22);
            AllCodon.Add(C23);
            AllCodon.Add(C24);
            AllAmino.Add(A7);
            #endregion
            #region His
            Amino A8 = new Amino();
            A8.AminoName = "His";
            A8.Frequence_Amino = 0;
            A8.N = 2;
            A8.codon.Add(C25);
            A8.codon.Add(C26);
            C25.Parent = A8;
            C26.Parent = A8;
            AllCodon.Add(C25);
            AllCodon.Add(C26);
            AllAmino.Add(A8);
            #endregion
            #region Gln
            Amino A9 = new Amino();
            A9.Frequence_Amino = 0;
            A9.AminoName = "Gln";
            A9.N = 2;
            A9.codon.Add(C27);
            A9.codon.Add(C28);
            C27.Parent = A9;
            C28.Parent = A9;
            AllCodon.Add(C27);
            AllCodon.Add(C28);
            AllAmino.Add(A9);
            #endregion
            #region Arg
            Amino A10 = new Amino();
            A10.Frequence_Amino = 0;
            A10.AminoName = "Arg";
            A10.N = 6;
            A10.codon.Add(C29);
            A10.codon.Add(C30);
            A10.codon.Add(C31);
            A10.codon.Add(C32);
            A10.codon.Add(C47);
            A10.codon.Add(C48);
            C29.Parent = A10;
            C30.Parent = A10;
            C31.Parent = A10;
            C32.Parent = A10;
            C47.Parent = A10;
            C48.Parent = A10;
            AllCodon.Add(C29);
            AllCodon.Add(C30);
            AllCodon.Add(C31);
            AllCodon.Add(C32);
            AllAmino.Add(A10);
            AllCodon.Add(C47);
            AllCodon.Add(C48);
            #endregion
            #region Ile
            Amino A11 = new Amino();
            A11.AminoName = "Ile";
            A11.Frequence_Amino = 0;
            A11.N = 3;
            A11.codon.Add(C33);
            A11.codon.Add(C34);
            A11.codon.Add(C35);
            C33.Parent = A11;
            C34.Parent = A11;
            C35.Parent = A11;
            AllCodon.Add(C33);
            AllCodon.Add(C34);
            AllCodon.Add(C35);
            AllAmino.Add(A11);
            #endregion
            #region Met
            Amino A12 = new Amino();
            A12.Frequence_Amino = 0;
            A12.AminoName = "Met";
            A12.N = 1;
            A12.codon.Add(C36);
            C36.Parent = A12;
            AllCodon.Add(C36);
            AllAmino.Add(A12);
            #endregion
            #region Thr
            Amino A13 = new Amino();
            A13.Frequence_Amino = 0;
            A13.AminoName = "Thr";
            A13.N = 4;
            A13.codon.Add(C37);
            A13.codon.Add(C38);
            A13.codon.Add(C39);
            A13.codon.Add(C40);
            C37.Parent = A13;
            C38.Parent = A13;
            C39.Parent = A13;
            C40.Parent = A13;
            AllCodon.Add(C37);
            AllCodon.Add(C38);
            AllCodon.Add(C39);
            AllCodon.Add(C40);
            AllAmino.Add(A13);
            #endregion
            #region Asn
            Amino A14 = new Amino();
            A14.AminoName = "Asn";
            A14.Frequence_Amino = 0;
            A14.N = 2;
            A14.codon.Add(C41);
            A14.codon.Add(C42);
            C41.Parent = A14;
            C42.Parent = A14;
            AllCodon.Add(C41);
            AllCodon.Add(C42);
            AllAmino.Add(A14);
            #endregion
            #region Lys
            Amino A15 = new Amino();
            A15.AminoName = "Lys";
            A15.Frequence_Amino = 0;
            A15.N = 2;
            A15.codon.Add(C43);
            A15.codon.Add(C44);
            C43.Parent = A15;
            C44.Parent = A15;
            AllCodon.Add(C43);
            AllCodon.Add(C44);
            AllAmino.Add(A15);
            #endregion
            #region Val
            Amino A16 = new Amino();
            A16.AminoName = "Val";
            A16.Frequence_Amino = 0;
            A16.N = 4;
            A16.codon.Add(C49);
            A16.codon.Add(C50);
            A16.codon.Add(C51);
            A16.codon.Add(C52);
            C49.Parent = A16;
            C50.Parent = A16;
            C51.Parent = A16;
            C52.Parent = A16;
            AllCodon.Add(C49);
            AllCodon.Add(C50);
            AllCodon.Add(C51);
            AllCodon.Add(C52);
            AllAmino.Add(A16);
            #endregion
            #region Ala
            Amino A17 = new Amino();
            A17.AminoName = "Ala";
            A17.Frequence_Amino = 0;
            A17.N = 4;
            A17.codon.Add(C53);
            A17.codon.Add(C54);
            A17.codon.Add(C55);
            A17.codon.Add(C56);
            C53.Parent = A17;
            C54.Parent = A17;
            C55.Parent = A17;
            C56.Parent = A17;
            AllCodon.Add(C53);
            AllCodon.Add(C54);
            AllCodon.Add(C55);
            AllCodon.Add(C56);
            AllAmino.Add(A17);
            #endregion
            #region Asp
            Amino A18 = new Amino();
            A18.AminoName = "Asp";
            A18.Frequence_Amino = 0;
            A18.N = 2;
            A18.codon.Add(C57);
            A18.codon.Add(C58);
            C57.Parent = A18;
            C58.Parent = A18;
            AllCodon.Add(C57);
            AllCodon.Add(C58);
            AllAmino.Add(A18);
            #endregion
            #region Glu
            Amino A19 = new Amino();
            A19.AminoName = "Glu";
            A19.Frequence_Amino = 0;
            A19.N = 2;
            A19.codon.Add(C59);
            A19.codon.Add(C60);
            C59.Parent = A19;
            C60.Parent = A19;
            AllCodon.Add(C59);
            AllCodon.Add(C60);
            AllAmino.Add(A19);
            #endregion
            #region Gly
            Amino A20 = new Amino();
            A20.AminoName = "Gly";
            A20.Frequence_Amino = 0;
            A20.N = 4;
            A20.codon.Add(C61);
            A20.codon.Add(C62);
            A20.codon.Add(C63);
            A20.codon.Add(C64);
            C61.Parent = A20;
            C62.Parent = A20;
            C63.Parent = A20;
            C64.Parent = A20;
            AllCodon.Add(C61);
            AllCodon.Add(C62);
            AllCodon.Add(C63);
            AllCodon.Add(C64);
            AllAmino.Add(A20);
            #endregion
            #region Stop 
            Amino A21 = new Amino();
            A21.AminoName = "Stop";
            A21.Frequence_Amino = 0;
            A21.N = 3;
            A21.codon.Add(C11);
            A21.codon.Add(C12);
            A21.codon.Add(C15);
            C11.Parent = A21;
            C12.Parent = A21;
            C15.Parent = A21;
            AllCodon.Add(C11);
            AllCodon.Add(C12);
            AllCodon.Add(C15);
            AllAmino.Add(A21);
            #endregion
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string line;
            string linefirst = "";
            //Open File
            bool firstline = true;// firstline is true and other line is false
            OpenFileDialog op = new OpenFileDialog();

            op.InitialDirectory = "e:\\";
            op.Filter = "fasta files (*.fasta)|*.fasta|All files (*.*)|*.*";
            op.FilterIndex = 2;
            op.RestoreDirectory = true;
            op.Multiselect = true;
            op.Title = "Please Select Source File(s) ";
            if (op.ShowDialog() == DialogResult.OK)
            {
                foreach (String file in op.FileNames)
                {
                    try
                    {

                        TextReader Tr = new StreamReader(file);

                        do
                        {
                            line = (Tr.ReadLine()).Trim();
                            if (firstline)
                            {
                                firstline = false;
                                linefirst = line;
                                continue;
                            }
                            else
                            {
                                strb.Append(line);
                            }

                        } while (!(string.IsNullOrEmpty(line)));

                        strb = strb.Replace(" ", string.Empty);
                        FileList.Add(strb.ToString());
                        address.Add(linefirst);
                        strb = strb.Remove(0, strb.Length);
                        firstline = true;
                        Tr.Close();
                        Tr = null;

                        button2.Enabled = false;
                        button1.Enabled = false;
                        button3.Enabled = true;
                        radioButton3.Enabled = false;
                        radioButton4.Enabled = false;


                    }

                    catch (Exception ex)
                    {

                        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);

                    }

                }
                if (FileList.Count != 0)
                    MessageBox.Show("Upload was successful.\nPlease go to the Output tab");


            }

            foreach (var item in address)
            {
                richTextBox2.Text += item.ToString();
                richTextBox2.Text += ("\n").ToString();
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            groupBox1.Enabled = false;
            groupBox2.Enabled = false;
            button5.Enabled = false;
            button3.Enabled = false;


        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton3.Checked)
            {
                groupBox1.Enabled = true;
                groupBox2.Enabled = false;
                button1.Enabled = true;

            }

        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton4.Checked)
            {
                groupBox2.Enabled = true;
                groupBox1.Enabled = false;
                button2.Enabled = true;
            }


        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (richTextBox1.Text == "")
            { MessageBox.Show("Textbox Is Empty.Please Enter The Text"); }
            else
            {
                if (checkBox2.Checked)
                {
                    try
                    {
                        string text = richTextBox1.Text;
                        string[] lines = text.Split('\n');
                        bool firstline = true;
                        foreach (string line in lines)
                        {
                            if (firstline)
                            {
                                firstline = false;
                                address.Add(line.ToString());
                                continue;
                            }
                            else
                            {
                                strb.Append(line.Trim());
                            }
                        }
                        strb = strb.Replace(" ", string.Empty);
                        FileList.Add(strb.ToString());
                        strb = strb.Remove(0, strb.Length);
                        firstline = true;
                        text = "";
                        button2.Enabled = false;
                        button1.Enabled = false;
                        button3.Enabled = true;
                        radioButton3.Enabled = false;
                        radioButton4.Enabled = false;

                        MessageBox.Show("Upload was successful.\nPlease go to the Output tab");
                        for (int i = 0; i < lines.Length; i++)
                            lines[i] = string.Empty;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not Upload. Original error: " + ex.Message);
                    }
                }


                else
                {
                    try
                    {
                        string text = richTextBox1.Text;
                        string[] lines = text.Split('\n');
                        foreach (string line in lines)
                        {
                            strb.Append(line.Trim());

                        }
                        strb = strb.Replace(" ", string.Empty);
                        FileList.Add(strb.ToString());
                        strb = strb.Remove(0, strb.Length);
                        text = "";
                        address.Add("Genes uploaded");
                        button2.Enabled = false;
                        button1.Enabled = false;
                        button3.Enabled = true;
                        radioButton3.Enabled = false;
                        radioButton4.Enabled = false;
                        MessageBox.Show("Upload was successful.\nPlease go to the Output tab");

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: Could not Upload. Original error: " + ex.Message);
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Stopwatch sw = new System.Diagnostics.Stopwatch();
            sw.Start();
            foreach (var item in FileList)
            {
                sw.Start();
                Search(item);
                Calculation();
                Show();
                foreach (var Case in AllCodon)
                {
                    Case.W = 0;
                    Case.RSCU = 0;
                    Case.Frequence_Codon = 0;
                }
                foreach (var Case in AllAmino)
                {
                    Case.Frequence_Amino = 0;
                    Case.RSCU_Max = 0;
                    Case.X_Max = 0;
                }
            }
            sw.Stop();

            //MessageBox.Show(sw.ElapsedMilliseconds.ToString())  ;
            button3.Enabled = false;
            button5.Enabled = true;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Clean();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            object mis = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Microsoft.Office.Interop.Excel._Worksheet worksheet = null;
            app.Visible = false;
            worksheet = (Worksheet)workbook.Sheets["Sheet1"];
            worksheet = (Worksheet)workbook.ActiveSheet;
            worksheet.Name = "Export";
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                }
            }

            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel Document(*.xlsx)|*.xlsx";
            sfd.FileName = "Export";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                workbook.SaveAs(sfd.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            }
            app.Quit();
            MessageBox.Show("Export successfully completed");
        }

        void CutAction(object sender, EventArgs e)
        {
            richTextBox1.Cut();
        }
        void CopyAction(object sender, EventArgs e)
        {
            Clipboard.SetText(richTextBox1.SelectedText);
        }

        void PasteAction(object sender, EventArgs e)
        {
            if (Clipboard.ContainsText())
            {
                richTextBox1.Text
                    += Clipboard.GetText(TextDataFormat.Text).ToString();
            }
        }

        private void richTextBox1_MouseUp(object sender, MouseEventArgs e)
        {

            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {   //click event

                ContextMenu contextMenu = new System.Windows.Forms.ContextMenu();
                System.Windows.Forms.MenuItem menuItem = new System.Windows.Forms.MenuItem("Cut");
                menuItem.Click += new EventHandler(CutAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new System.Windows.Forms.MenuItem("Copy");
                menuItem.Click += new EventHandler(CopyAction);
                contextMenu.MenuItems.Add(menuItem);
                menuItem = new System.Windows.Forms.MenuItem("Paste");
                menuItem.Click += new EventHandler(PasteAction);
                contextMenu.MenuItems.Add(menuItem);

                richTextBox1.ContextMenu = contextMenu;
            }

        }
        
    }
}
