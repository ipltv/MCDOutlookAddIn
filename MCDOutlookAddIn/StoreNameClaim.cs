using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MCDOutlookAddIn
{
    /// <summary>
    /// Класс предназначен для получения названия ПБО по передаваемому номеру. 
    /// Класс StoreNameClaim представляет из себя корневой объект для работы со справочником имен ПБО.
    /// Класс предполагает работу со справочником ПБО. Справочник может иметь структуру по умолчанию формата: Номер ПБО - Название ПБО.
    /// Структура справочника также может быть задана индивидуально с помощью регулярных выражений.
    /// Класс загружает информацию из справочника и инкапсулирует ее в словаре.
    /// </summary>
    public class StoreNameClaim
    {
        /// <summary>
        /// Словарь содержащий информацию о номере ПБО (ключ) и названии ПБО (значение).
        /// </summary>
        private readonly Dictionary<string, string> stores;
        /// <summary>
        /// Объект FileInfo представляющий файл-справочник, который содержит информацию о ПБО.
        /// </summary>
        private readonly FileInfo handbook;
        /// <summary>
        /// Шаблон по умолчанию для проверки информации в справочнике.
        /// </summary>
        private readonly string bookStringTemplate = @"^\d{5}\t\w*";
        /// <summary>
        /// Создает объект класс StoreNameClaim. Структура справочника задается по умолчанию.
        /// </summary>
        /// <param name="path">Путь к файлу справочника</param>
        public StoreNameClaim(string path)
        {
            if (File.Exists(path))
            {
                handbook = new FileInfo(path);
                stores = HandBookReader(handbook, bookStringTemplate);
            }
            else
            {
                throw new FileNotFoundException("Handbook file not found");
            }
        }
        /// <summary>
        /// Метод для чтения информации из файл-справочника.
        /// Каждая строка файла проверяется на соответсвие заданному шаблону.
        /// Если строка не соответствует шаблону, данные из нее отбрасываются. 
        /// </summary>
        /// <param name="file">Объект FileInfo, который представляет файл-справочник</param>
        /// <param name="pattern">Шаблон для сравнения строки на основе регулярных выражений</param>
        /// <returns></returns>
        private Dictionary<string, string> HandBookReader(FileInfo file, string pattern)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            Regex regex = new Regex(pattern);
            using (StreamReader reader = new StreamReader(file.FullName))
            {
                string line;
                while ((line = reader.ReadLine()) != null || !reader.EndOfStream)
                {
                    if (regex.IsMatch(line))
                    {
                        result.Add(line.Substring(0, 5).Trim(), line.Substring(5).Trim());
                    }
                }
            }
            return result;
        }
        /// <summary>
        /// Возвращает название ПБО из справочника по номеру.
        /// </summary>
        /// <param name="number">Строковое представление номера</param>
        /// <returns></returns>
        public string GetStoreNameByNumber(string number)
        {
            if (stores.ContainsKey(number))
            {
                return stores[number];
            }
            else
            {
                return null;
            }
        }
        /// <summary>
        /// Возвращает название ПБО из справочника по номеру.
        /// </summary>
        /// <param name="number">Численное представление номера</param>
        /// <returns></returns>
        public string GetStoreNameByNumber(int number)
        {
            return GetStoreNameByNumber(number.ToString());
        }
    }
}
