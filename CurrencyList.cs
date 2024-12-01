using System.Collections.Generic;
using System.Xml;


namespace TestTaskCurrency
{
    /// <summary>
    /// Класс массива валют
    /// </summary>
    public class CurrencyList
    {
        /// <summary>
        /// Массив валют
        /// </summary>
        public List<Currency> List = new List<Currency>();

        /// <summary>
        /// Создание массива валют
        /// </summary>
        /// <param name="xmlNodes">Массив xml-переменной, посвященный валютам</param>
        public void Create(XmlNodeList xmlNodes)
        {
            List = new List<Currency>();

            foreach (XmlNode node in xmlNodes)
            {
                List.Add(new Currency(node));
            }
        }

        /// <summary>
        /// Проверка на наличие в массиве валюты
        /// </summary>
        /// <param name="checkCurrency">Валюта</param>
        /// <returns>Истина, если в массиве есть валюта с идентичным буквенным кодом</returns>
        public bool Contains(Currency checkCurrency)
        {
            foreach (Currency c in List)
            {
                if (c.CharCode.Equals(checkCurrency.CharCode))
                    return true;
            }
            return false;
        }
    }
}