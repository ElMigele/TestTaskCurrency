using System;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using HtmlAgilityPack;
using ClosedXML.Excel;
using System.IO;


namespace TestTaskCurrency
{
    /// <summary>
    /// Класс формы
    /// </summary>
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        CurrencyList currencyList;

        // Задание сегодняшнего числа в календаре
        private void Form1_Load(object sender, EventArgs e)
        {
            dateTimePicker1.Value = DateTime.Today;
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            textBox1.Text = "Начало работы - дата: " + dateTimePicker1.Value.ToString("dd.MM.yyyy");
            button1.Enabled = false;
            string site = "http://www.cbr.ru/scripts/XML_daily.asp?date_req="
                    + dateTimePicker1.Value.ToString("dd/mm/yyyy");
            
            int i = 0;
            while (currencyList == null && i < 6)
            {
                i++;
                textBox1.Text += "\r\nПопытка подключения к сайту #" + i;
                currencyList = await Work();
            }
            if (currencyList == null)
            {
                textBox1.Text += "\r\nНе удалось получить данные с сайта!";
                button1.Enabled = true;
                return;
            }

            textBox1.Text += "\r\nДанные с сайта получены!";

            // Добавление в перечень валют российского рубля
            Currency rub = new Currency("Российский рубль", "643", "RUB", "1");
            if (!currencyList.Contains(rub))
            {
                currencyList.List.Add(rub);
                currencyList.List = currencyList.List.OrderBy(o => o.CharCode).ToList();
            }
            ExcelGeneration();
            
            textBox1.Text += "\r\nФайл xlsx собран";
            button1.Enabled = true;
        }

        /// <summary>
        /// Извлечение данных с сайта ЦБ РФ
        /// </summary>
        /// <returns>Массив валют</returns>
        private async Task<CurrencyList> Work()
        {
            CurrencyList result = new CurrencyList();
            string site = "http://www.cbr.ru/scripts/XML_daily.asp?date_req="
                    + dateTimePicker1.Value.Day + "/"
                    + dateTimePicker1.Value.Month + "/"
                    + dateTimePicker1.Value.Year;

            using (HttpClient client = new HttpClient())
            {
                try
                {
                    string xmlContent = await client.GetStringAsync(site);
                    XDocument doc = XDocument.Parse(xmlContent);
                    string z = doc.Root.ToString();

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.LoadXml(z);

                    XmlNodeList xmlNodes = xmlDoc.GetElementsByTagName("Valute");

                    result.Create(xmlNodes);
                }
                catch
                {
                    return null;
                }
            }
            return result;
        }

        /// <summary>
        /// Создание xlsx-файла
        /// </summary>
        private void ExcelGeneration()
        {
            var wb = new XLWorkbook();

            for (int i = 0; i < currencyList.List.Count; i++)
            {
                var ws = wb.Worksheets.Add(currencyList.List[i].CharCode);
                float currentCurrency = currencyList.List[i].Value;

                ws.Cell(1, 1).Value = "Код";
                ws.Cell(1, 2).Value = "Название";
                ws.Cell(1, 3).Value = "Курс";

                int passFlag = 1;
                for (int j = 0; j < currencyList.List.Count; j++)
                {
                    if (i == j)
                    {
                        passFlag = 0;
                        continue;
                    }
                    ws.Cell(j + passFlag + 1, 1).Value = currencyList.List[j].CharCode;
                    ws.Cell(j + passFlag + 1, 2).Value = currencyList.List[j].Name;
                    ws.Cell(j + passFlag + 1, 3).Value = currencyList.List[j].Value / currentCurrency;
                }
                ws.Column(1).AdjustToContents();
                ws.Column(2).AdjustToContents();
                ws.Column(3).CellsUsed().Style.NumberFormat.Format = "# ##0.0000";
                ws.Column(3).AdjustToContents();
            }

            string fileName = dateTimePicker1.Value.ToString("yyyyMMdd") + ".xlsx";
            textBox1.Text += "\r\nСоздание файла " + fileName;
            if (Directory.Exists(Directory.GetCurrentDirectory() + "//" + fileName))
                Directory.Delete(Directory.GetCurrentDirectory() + "//" + fileName);
            wb.SaveAs(fileName);
        }

        // Защита от некорректного задания даты
        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePicker1.Value > DateTime.Today
            || dateTimePicker1.Value <= DateTime.Parse("01/01/1994"))
                dateTimePicker1.Value = DateTime.Today;
        }
    }
}