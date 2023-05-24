using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Data;
using System.Net;
using System.Reflection;

namespace ReadExcelDemo
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private Dictionary<string, List<string>> getPropNames()
        {
            var keyValuePairs = new Dictionary<string, List<string>>();

            keyValuePairs.Add("AccountNumber"
          , new List<string>() { "شماره حساب", "ش حساب" });

            keyValuePairs.Add("IBAN"
          , new List<string>() { "شماره شبا" });

            keyValuePairs.Add("NationalCode"
          , new List<string>() { "کد ملی", "شناسه ملی" ,"ش ملی","شماره ملی", "ش.م", "ش م" });

            keyValuePairs.Add("FirstName"
          , new List<string>() { "نام", "اسم" });

            keyValuePairs.Add("LastName"
          , new List<string>() { "نام خانوادگی", "فامیلی" });

            keyValuePairs.Add("DocumentDescription"
          , new List<string>() { "شرح سند" });

            keyValuePairs.Add("Debtor"
          , new List<string>() { "بدهکار" });

            keyValuePairs.Add("Creditor"
          , new List<string>() { "بستانکار" });

            keyValuePairs.Add("AccountBalance"
          , new List<string>() { "مانده حساب" });

            keyValuePairs.Add("DocumentNumber"
          , new List<string>() { "شماره سند", "ش سند" });

            keyValuePairs.Add("TerminalNumber"
          , new List<string>() { "شماره ترمینال", "ش ترمینال" });

            keyValuePairs.Add("TrackingCode"
          , new List<string>() { "کد پیگیری" });

            keyValuePairs.Add("BankName"
          , new List<string>() { "نام بانک" });

            keyValuePairs.Add("BankCode"
          , new List<string>() { "کد بانک" });

            keyValuePairs.Add("BranchName"
          , new List<string>() { "نام شعبه" });

            keyValuePairs.Add("BranchCode"
          , new List<string>() { "کد شعبه" });

            return keyValuePairs;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        var propNames = getPropNames();

                        List<DataModel> models = new List<DataModel>();

                        var workbook = new XLWorkbook(file.FileName);
                        var ws1 = workbook.Worksheet(1);

                        var rowCount = ws1.RowCount();
                        var firstRow = ws1.FirstRow().RowUsed();
                        var headerTitles = firstRow.CellsUsed().ToList();

                        IDictionary<string, string> celsName = headerTitles
                            .Where(cell => propNames.Any(a => a.Value.Contains(cell.GetValue<string>())))
                            .ToDictionary(pair => pair.Address.ColumnLetter, pair => propNames.FirstOrDefault(a => a.Value.Contains(pair.GetValue<string>())).Key);

                        var rows = ws1.RowsUsed();
                        foreach (var row in rows)
                        {
                            var newModel = new DataModel();
                            var type = newModel.GetType();
                            var props = new List<PropertyInfo>(type.GetProperties().ToList()).ToArray();

                            foreach (var cell in row.CellsUsed())
                            {
                                var celValue = cell.GetValue<string>();
                                if (celsName.Any(a => a.Key == cell.Address.ColumnLetter))
                                {
                                    var celName = celsName[cell.Address.ColumnLetter];
                                    var prop = props.First(a => a.Name == celName);
                                    prop.SetValue(newModel, cell.GetValue<string>());
                                }
                            }
                            models.Add(newModel);
                        }

                        dataGridView1.DataSource = ToDataTable(models);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("لطفا فایل با فرمت .xls  یا .xlsx انتخاب کنید کنید", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }


        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                dataTable.Columns.Add(prop.Name);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            return dataTable;
        }
    }
}