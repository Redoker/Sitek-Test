using System.Net;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.Json;

namespace Sitek
{   
    //объект API
    public class CusApi
    {
        public string region { get; set; }
        public string name { get; set; }
        public string code { get; set; }
        public string type { get; set; }
    }

    public class Root
    {
        public List<CusApi> cus { get; set; }
    }

    public class Program
    {
        static void Main(string[] args)
        {
            Root jsonResponse;

            //Объекст запроса
            WebRequest request = WebRequest.Create("https://api.kontur.ru/dc.contacts/v1/cus");

            request.Method = "GET";
            
            //Отправляем запрос
            using (WebResponse response = request.GetResponse())
            {
                //Читаем ответ в поток, чтобы прочитать всю информацию
                using (Stream stream = response.GetResponseStream())
                {
                    StreamReader reader = new StreamReader(stream);
                    //Получаем ответ в строку потока
                    var jsonStream = reader.ReadToEnd();

                    //десериализируем json
                    jsonResponse = JsonConvert.DeserializeObject<Root>(jsonStream);
                }
            }

            //получаем текущую дату на пк и выводим
            Console.WriteLine("Текущая дата и время: " + DateTime.Now);

            //получаем кол-во записей с регионом 18
            Console.WriteLine("Общее кол-во записей: " + jsonResponse.cus.Count(x => x.region == "18") + "\n");

            //через linq делаем сортировку и проходим по каждому элементу
            //если регион элемента равен 18, то выводим его данные
            foreach (var item in jsonResponse.cus.OrderBy(x => x.type).ThenBy(x => x.code))
            {
                if (item.region == "18")
                {
                    Console.WriteLine($"{item.type}, {item.code}, {item.name}, {item.region}");
                }
            }

            ExcelWorksheet excelWorksheet = new ExcelWorksheet();
            excelWorksheet.CreateFile(jsonResponse);
        }
    }
}