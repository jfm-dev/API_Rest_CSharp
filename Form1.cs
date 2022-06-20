using RestCountries;
using System.Data;
using System.Threading;
using System.IO;
using excel = Microsoft.Office.Interop.Excel;


namespace CREST
{
    public partial class Form1 : Form
    {


        public Form1()
        {

            var c = new Country();
            InitializeComponent();

        }

        private async void Form1_Load(object sender, EventArgs e)
        {
           await GetcounriesAsync();

        }



        async Task GetcounriesAsync()
        {
             List<Country> list = new List<Country>();
            var countries = new RestCountriesClient(server: "https://restcountries.com/", apiRoute: "/v2");
            var country = await countries.GetAllCountriesAsync();

     

            foreach (var item in country)
            {

               

                var c = new Country();
                c.Nome = item.Name;
                c.Nomenativo = item.NativeName;
                c.Populacao = item.Population;
                c.Capital = item.CapitalCity;
                c.Area = item.Area.ToString();
                c.Bandeira = item.FlagUriString;
                c.Fusohorario = "TimeZone";
                c.Subregiao = "Subregiao";
                c.Regiao = "Region";
                list.Add(c);

        
            }


            tablecountries.DataSource = list;
        }

        private void bunifuThinButton21_Click(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            


            Thread t = new Thread((ThreadStart)(() => {


               tablecountries.SelectAll();
               DataObject data = tablecountries.GetClipboardContent();
               if(data != null)Clipboard.SetDataObject(data);
               Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
               excel.Visible = true;
               Microsoft.Office.Interop.Excel.Workbook excelbook;
               Microsoft.Office.Interop.Excel.Worksheet excelsheet;
               
               object misseddata = System.Reflection.Missing.Value;
               excelbook = excel.Workbooks.Add(misseddata);
                
               excelsheet = (Microsoft.Office.Interop.Excel.Worksheet)excelbook.Worksheets.get_Item(1);
               Microsoft.Office.Interop.Excel.Range excelrang = (Microsoft.Office.Interop.Excel.Range)excelsheet.Cells[2, 1];
               excelrang.Select();
               excelsheet.Columns.AutoFit();
               excelsheet.PasteSpecial(excelrang, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, true);
          
            }));
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            var dt = new DataTable();
            foreach(DataGridViewColumn column in tablecountries.Columns)
            {
                if (column.Visible)
                {
                    dt.Columns.Add();
                }
            }
            object[] cellValues = new  object[tablecountries.Columns.Count];
            foreach(DataGridViewRow row in tablecountries.Rows)
            {
                for(int i = 0; i < row.Cells.Count; i++)
                {
                    cellValues[i] = row.Cells[i].Value;
                }
                dt.Rows.Add(cellValues);
            }
            string dir = @"C:\Countries";
          
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            ds.WriteXml(File.OpenWrite(@"C:\Countries\countries.xml"));
            var path = @"C:\Countries\countries.xml";
            MessageBox.Show("O ficheiro foi exportado com sucesso. Encontre o ficheiro neste caminho:"+path,"Sucesso", MessageBoxButtons.OK);
            
            
        }
    }
}