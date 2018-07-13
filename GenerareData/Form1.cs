using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
namespace GenerareData
{
    public partial class Form1 : Form
    {
        #region Variabile globale
        Workbook workbook;
        Microsoft.Office.Interop.Excel.Application app;
        Worksheet worksheet;
        GeneratorDate generatorDate;
        DateTime dataStart;
        DateTime dataSfarsit;
        int i = 2; int j = 0; int k = 2; int count = 0;
        string info = string.Empty;
        int procent = 0;
        #endregion

        public Form1()
        {
            InitializeComponent();
        }

        #region Cod Ajustare
        private void DateZilnice()
        {
            button1.Enabled = false;

            int rcount = worksheet.UsedRange.Rows.Count;
            DateTime col1, col2;
            col1 = Convert.ToDateTime(worksheet.Cells[rcount, 1].Value);
            col2 = Convert.ToDateTime(worksheet.Cells[rcount, 4].Value);
            int result = DateTime.Compare(col1, col2);

            DeterminaDataStartSfarsit(rcount, result);

            int marimeVector = Convert.ToInt32((dataSfarsit - dataStart).TotalDays);
            string[] vector = new string[marimeVector];

            worksheet.Cells[k - 1, 5].Value = "Ajustare";
            worksheet.Cells[k - 1, 4].Value = "Date zilnice";

            progressBar1.Minimum = 0;
            progressBar1.Maximum = rcount + 1;
            progressBar1.Visible = true;

            ScrieAjustarileZilniceInFisier(vector, rcount);

            label2.Text = "Ajustarea zilelor s-a efectuat, te rog asteapta!";
            count = 0;
            workbook.Save();
            i = 2; j = 0; k = 2; count = 0;
        }
        private void DeterminaDataStartSfarsit(int rcount, int result)
        {
            try
            {
                dataStart = DateTime.Parse(worksheet.Cells[2, 1].Value.ToShortDateString());
                if (result == 0)
                    dataSfarsit = DateTime.Parse(worksheet.Cells[rcount, 4].Value);
                else if (result < 0)
                    dataSfarsit = DateTime.Parse(worksheet.Cells[rcount, 4].Value);
                else
                    dataSfarsit = DateTime.Parse(worksheet.Cells[rcount, 1].Value.ToShortDateString());
            }
            catch (Exception)
            {
                MessageBox.Show("Formatul datei de pe coloana [A] este gresit sau ai spatii goale!",
                    "Motivul erorii", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                InchideExcel();
                System.Windows.Forms.Application.Restart();
            }
        }
        private void ScrieAjustarileZilniceInFisier(string[] vector, int rcount)
        {
            foreach (var date in generatorDate.GenereazaData(dataStart, dataSfarsit))
            {
                try
                {
                    vector[j] = date.ToShortDateString();
                    if (worksheet.Cells[i, 1].Value.ToShortDateString() != vector[j])
                    {
                        worksheet.Cells[k, 5].Value = worksheet.Cells[i - 1, 2].Value;
                        worksheet.Cells[k, 5].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        worksheet.Cells[k, 4].Value = date.ToShortDateString();
                        worksheet.Cells[k, 4].Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                        i--; count++;
                    }
                    else
                    {
                        worksheet.Cells[k, 5].Value = worksheet.Cells[i, 2].Value;
                        worksheet.Cells[k, 4].Value = date.ToShortDateString();
                    }
                    i++; j++; k++;
                    progressBar1.Value = i;
                    procent = ((i * 100) / rcount);
                    label4.Text = "Progres: " + procent + "%";
                    if (count > 0)
                        label2.Text = "Ajustez valorile zilnice: am ajustat " + count.ToString() + " valori pana acum...";
                    else label2.Text = "Verific valorile...";
                }
                catch (Exception)
                {
                    label2.Text = "Am gasit o eroare!";
                    MessageBox.Show("Cauzele erorii:\n" +
                        "Ai valori lipsa pe coloana [A] si [B]", "Eroare",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    InchideExcel();
                    System.Windows.Forms.Application.Restart();
                }
            }
        }
        private void DateLunare()
        {
            label2.Text = "Ajustez datele lunare...";
            int rcount = worksheet.UsedRange.Rows.Count;
            int w = 2, x = 2;
            var primaZiLuna = Convert.ToDateTime(worksheet.Cells[2, 1].Value);
            var dataFinal = Convert.ToDateTime(worksheet.Cells[rcount, 4].Value);

            worksheet.Cells[x - 1, 7].Value = "Date lunare";
            worksheet.Cells[x - 1, 8].Value = "Ajustare";

            foreach (DateTime data in generatorDate.GenereazaLuna(primaZiLuna, dataFinal))
            {
                try
                {
                    worksheet.Cells[x, 7].Value = data.ToShortDateString();
                    x++;
                }
                catch (Exception)
                {
                    MessageBox.Show("Exista valori pe coloana G", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    InchideExcel();
                    System.Windows.Forms.Application.Exit();
                }
            }
            worksheet.Cells[2, 7].Delete(XlDeleteShiftDirection.xlShiftUp);
            for (int z = 2; z < rcount; z++)
            {
                var data1 = Convert.ToDateTime(worksheet.Cells[z, 4].Value);
                var data2 = Convert.ToDateTime(worksheet.Cells[w, 7].Value);
                int rezultat = DateTime.Compare(data1, data2);
                if (rezultat == 0)
                {
                    try
                    {
                        worksheet.Cells[w, 8].Value = worksheet.Cells[z, 5].Value;
                        w++;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Exista valori pe coloana H", "Eroare", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        InchideExcel();
                        System.Windows.Forms.Application.Exit();
                    }
                }
                procent = (((z + 1) * 100) / rcount);
                label4.Text = "Progres: " + procent + "%";
            }
            workbook.Save();
        }
        private void VariableDummy(bool progress, string text, string text1, int pozitie, DayOfWeek dow)
        {
            int q = 1;
            label2.Text = text;
            int rcount = worksheet.UsedRange.Rows.Count;
            dataStart = Convert.ToDateTime(worksheet.Cells[2, 4].Value);
            dataSfarsit = Convert.ToDateTime(worksheet.Cells[rcount, 4].Value);
            worksheet.Cells[q, pozitie].Value = text1;
            foreach (DateTime data in generatorDate.GenereazaData(dataStart, dataSfarsit))
            {
                if (data.DayOfWeek == dow)
                {
                    worksheet.Cells[q + 1, pozitie].Value = "1";
                }
                else worksheet.Cells[q + 1, pozitie].Value = "0";
                q++;
                procent = (((q * 100) / rcount));
                label4.Text = "Progres: " + procent + "%";
            }
            workbook.Save();
        }
        private void PrimeleZile()
        {
            button1.Enabled = false;
            label2.Text = "Adaug primele 5 zile din an";
            int rcount = worksheet.UsedRange.Rows.Count;
            int w = 2;
            var inceputAn = Convert.ToDateTime(worksheet.Cells[2, 4].Value);
            var sfarsitAn = Convert.ToDateTime(worksheet.Cells[rcount, 4].Value);
            int x = 2;
            worksheet.Cells[x - 1, 15].Value = "Dummy 5 zile";
            string[] primeleZile = new string[rcount];
            int q = 0;
            foreach (DateTime data in generatorDate.GenereazaPrimeleZile(inceputAn, sfarsitAn))
            {
                primeleZile[q] = data.ToShortDateString();
                q++;
            }
            int b = 0;
            for (int v = 2; v <= rcount; v++)
            {
                var dat1 = Convert.ToDateTime(worksheet.Cells[v, 4].Value);
                var dat2 = Convert.ToDateTime(primeleZile[b]);
                int rezultat = DateTime.Compare(dat1, dat2);
                if (rezultat == 0)
                {
                    worksheet.Cells[v, 15].Value = 1;
                    b++;
                }
                else
                {
                    worksheet.Cells[v, 15].Value = 0;
                }
                procent = (((v) * 100) / rcount);
                label4.Text = "Progres: " + procent + "%";
            }
            label2.Text = "Gata, inchide programul, sau ia-o de la capat";
            workbook.Save();
            InchideExcel();
        }
        #endregion
        #region Main 
        public void MainExecution(object sender, EventArgs e)
        {
            label2.Text = "Pornesc conexiunea...";
            var cts = new CancellationTokenSource(); //don't bother

            Task.Factory.StartNew(() => DateZilnice(), cts.Token)
                .ContinueWith((task) =>
                {
                    DateLunare();
                }, cts.Token)
                .ContinueWith((task) =>
                {
                    VariableDummy(true, "Creez variabila dummy (1/5)", "D1-Luni", 10, DayOfWeek.Monday);
                    VariableDummy(true, "Creez variabila dummy (2/5)", "D2-Marti", 11, DayOfWeek.Tuesday);
                    VariableDummy(true, "Creez variabila dummy (3/5)", "D3-Miercuri", 12, DayOfWeek.Wednesday);
                    VariableDummy(true, "Creez variabila dummy (4/5)", "D4-Joi", 13, DayOfWeek.Thursday);
                    VariableDummy(true, "Creez variabila dummy (5/5)", "D5-Vineri", 14, DayOfWeek.Friday);
                }, cts.Token)
                .ContinueWith((task) =>
                {
                    PrimeleZile();
                }, cts.Token).ContinueWith((task) =>
                 {
                     label2.Text = "Toate ajustarile au fost efectuate cu succes.\nFisierul Excel este actualizat.";
                     button2.Enabled = true;
                     button1.Enabled = false;
                 });
        }
        #endregion
        #region Cauta fisier
        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Cauta fisier";
            ofd.Filter = ("Fisier .xlsx|*.xlsx");
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                app = new Microsoft.Office.Interop.Excel.Application();
                workbook = app.Workbooks.Open(ofd.FileName);
                worksheet = workbook.ActiveSheet;
                generatorDate = new GeneratorDate();
                button2.Enabled = false;
                button1.Enabled = true;
            }
        }
        #endregion
        #region Informatii
        private void label1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Program creat de Chiritoiu Culai\n\t\t\t\t\t-©SPE2016", "Copyright © CC", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            info = "Asigura-te de urmatoarele:\n" +
                    "-Pe prima coloana din Excel [A] ai doar valorile datei [Date]\n" +
                    " Formatul datei trebuie sa fie dd.MM.yyyy [zi.luna.an] (cu punct intre ele)\n" +
                    "-Asigura-te ca ai datele sortate in mod crescator:\n" +
                    "\tExemplu: din 25.07.2000 pana in 20.04.2016\n\n" +
                    "-Pe a doua coloana din Excel [B] ai valorile variabilei [Close]\n" +
                    "-Asigura-te ca pe coloana [A1] si B[1] ai [header](in caz contrar pierzi 2 valori)\n" +
                    "\n-Pe randurile coloanelor [A] sau [B] nu trebuie sa valori lipsa\n\n";
        }

        private void label3_Click(object sender, EventArgs e)
        {
            MessageBox.Show(info, "Cum functioneaza", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }
        #endregion
        #region Dispose object
        void InchideExcel()
        {
            app.Quit();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.FinalReleaseComObject(workbook);
            Marshal.FinalReleaseComObject(worksheet);
            Marshal.FinalReleaseComObject(app);
        }
        #endregion
    }
}