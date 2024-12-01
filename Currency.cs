using System;
using System.Xml;


namespace TestTaskCurrency
{
    /// <summary>
    /// Класс валюты
    /// </summary>
    public class Currency
    {
        /// <summary>
        /// Название
        /// </summary>
        public string Name;
        /// <summary>
        /// Цифровой код
        /// </summary>
        public ushort NumCode;
        /// <summary>
        /// Буквенный код
        /// </summary>
        public string CharCode;
        /// <summary>
        /// Стоимость в российских рублях
        /// </summary>
        public float Value;

        /// <summary>
        /// Конструктор валюты
        /// </summary>
        /// <param name="n">Название</param>
        /// <param name="nc">Цифровой код</param>
        /// <param name="cc">Буквенный код</param>
        /// <param name="v">Стоимость в российских рублях</param>
        public Currency(string n, string nc, string cc, string v)
        {
            Name = n;
            NumCode = UInt16.Parse(nc);
            CharCode = cc;
            Value = float.Parse(v);
        }

        /// <summary>
        /// Конструктор валюты
        /// </summary>
        /// <param name="xml">Часть xml-переменной, посвященного одной валюте</param>
        public Currency(XmlNode xml)
        {
            foreach (XmlNode node in xml.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Name":
                        Name = node.InnerText;
                        break;
                    case "NumCode":
                        NumCode = UInt16.Parse(node.InnerText);
                        break;
                    case "CharCode":
                        CharCode = node.InnerText;
                        break;
                    case "VunitRate":
                        Value = float.Parse(node.InnerText);
                        break;
                }
            }
        }
    }
}