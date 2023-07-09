using ClosedXML.Excel;
using System.Data;
using System.Reflection;

namespace ReadExcelDemo
{
    public partial class Form1 : Form
    {

        public List<DataModel> dataModels = new List<DataModel>();


        public Form1()
        {
            InitializeComponent();
        }

        private Dictionary<string, List<string>> getPropNames()
        {
            // درصورتی که در هدر اکسل از نام های فارسی استفاده شده باشد ،با تعریف مقادیر زیر 
            // اطلاعات ردیف مربوطه را در فیلد انتخابی قرار می دهد


            var keyValuePairs = new Dictionary<string, List<string>>();
            keyValuePairs.Add("FirstName"
          , new List<string>() { "نام", "اسم" });

            keyValuePairs.Add("LastName"
          , new List<string>() { "شهرت", "نام خانوادگی", "فامیلی" });

            keyValuePairs.Add("NationalCode"
        , new List<string>() { "کد ملی", "شماره ملی", "شناسه ملی", "ش ملی" });

            keyValuePairs.Add("PhoneNumber"
        , new List<string>() { "شماره موبایل", "موبایل", "تلفن همراه", "شماره تلفن همراه", "شماره تلفن" });

            keyValuePairs.Add("PhoneNumbeaaar"
     , new List<string>() { "شماره موبایل", "موبایل", "تلفن همراه", "شماره تلفن همراه", "شماره تلفن" });

            return keyValuePairs;
        }

        private void ImportExcel(object sender, EventArgs e)
        {

            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {

                    try
                    {
                        dataModels = convertExcelToDataModel<DataModel>(file).ToList();

                        dataGridView1.DataSource = ToDataTable(dataModels);


                        if (dataModels.Count() > 0)
                        {
                            ExportExcel.Visible = true;
                        }
                        else
                        {
                            ExportExcel.Visible = false;
                        }

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


        public System.Data.DataTable ToDataTable<T>(List<T> items)
        {
            System.Data.DataTable dataTable = new System.Data.DataTable(typeof(T).Name);
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                string propName = prop.CustomAttributes.Any(a => a.AttributeType.Name == "DisplayAttribute") ? prop.CustomAttributes.First(f => f.AttributeType.Name == "DisplayAttribute").NamedArguments.First(f => f.MemberName == "Name").TypedValue.Value.ToString() : prop.Name;
                dataTable.Columns.Add(propName);
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

        public byte[] convertDataModelToExcel<T>(IEnumerable<T> objs) where T : class
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("data");
                var currentRow = 1;
                worksheet.Cell(currentRow, 1).Value = "ردیف";
                worksheet.SetRightToLeft();
                worksheet.SetShowRowColHeaders();
                worksheet.SetAutoFilter();
                worksheet.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                worksheet.ColumnWidth = 30;
                worksheet.Row(1).Height = 30;

                Type myType = (objs.First()).GetType();
                var props = new List<PropertyInfo>(myType.GetProperties().ToList()).ToArray();
                worksheet.FirstRow().Cells(1, props.Length + 1).Style.Fill.SetBackgroundColor(XLColor.Tomato);
                worksheet.FirstRow().Cells(1, props.Length + 1).Style.Font.SetFontColor(XLColor.White);
                worksheet.FirstRow().Cells(1, props.Length + 1).Style.Font.SetBold();
                worksheet.FirstRow().Cells(1, props.Length + 1).Style.Font.SetFontSize(13.0);
                for (int i = 0; i < props.Length; i++)
                {
                    var prop = props[i];
                    var propname = prop.CustomAttributes.Any(a => a.AttributeType.Name == "DisplayAttribute") ? prop.CustomAttributes.First(f => f.AttributeType.Name == "DisplayAttribute").NamedArguments.First(f => f.MemberName == "Name").TypedValue.Value.ToString() : prop.Name;
                    worksheet.Cell(currentRow, i + 2).Value = propname;
                }
                foreach (var obj in objs)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = currentRow - 1;

                    for (int i = 0; i < props.Length; i++)
                    {
                        var value = props[i].GetValue(obj)?.ToString();
                        worksheet.Cell(currentRow, i + 2).Value = value;
                    }
                }
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();
                    return content;
                }

            }
        }


        private void ExportExcel_Click(object sender, EventArgs e)
        {
            var excel = convertDataModelToExcel(dataModels);
            string fileName = $"DataToExcel-{DateTime.Now.Ticks}.xls";
            using (FileStream fsNew = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                fsNew.Write(excel, 0, excel.Length);
            }
        }

        private IEnumerable<T> convertExcelToDataModel<T>(OpenFileDialog file) where T : class
        {
            List<T> data = new List<T>();
            var propNames = getPropNames();

            var type = typeof(T);

            var props = new List<PropertyInfo>(type.GetProperties().ToList()).ToArray();


            var workbook = new XLWorkbook(file.FileName);
            var ws1 = workbook.Worksheet(1);

            var rowCount = ws1.RowCount();
            var rows = ws1.RowsUsed().Skip(1);
            var firstRow = ws1.RowsUsed().First();
            var headerTitles = firstRow.CellsUsed().ToList();

            IDictionary<string, string> celsName = headerTitles
                .ToDictionary(
                pair => pair.Address.ColumnLetter,
                pair =>
                propNames.Any(a => a.Value.Contains(pair.GetValue<string>())) ? propNames.FirstOrDefault(a => a.Value.Contains(pair.GetValue<string>())).Key  // اگه نام فارسیش وجود داشت کلیدش رو بزار
                : (propNames.Any(a => a.Key.Contains(pair.GetValue<string>())) ? pair.GetValue<string>() // اگه هدر برابر با نام پراپرتی بود خودشو بزار
                : (props.Any(a => a.Name == pair.GetValue<string>()) ? pair.GetValue<string>() : "NotSet"))); // اگه تو لیست نام ها وجود نداشت ولی توی پراپرتی های کلاس وجود داشت ، اسم پراپرتی رو بزار

            foreach (var row in rows)
            {
                var newModel = (T)Activator.CreateInstance(type);
                foreach (var cell in row.CellsUsed().Where(w => celsName.Select(s => s.Key).Contains(w.Address.ColumnLetter)))
                {

                    var celValue = cell.GetValue<string>();
                    var celName = celsName[cell.Address.ColumnLetter];
                    var prop = props.FirstOrDefault(a => a.Name == celName);
                    if (prop != null)
                    {
                        prop.SetValue(newModel, celValue);
                    }

                }
                data.Add(newModel);
            }

            return data;

        }
    }
}
