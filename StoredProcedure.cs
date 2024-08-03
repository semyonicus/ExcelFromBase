using Spire.Xls;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using Newtonsoft.Json.Linq;
using System.ComponentModel;

namespace ReadExcelExample
{
    class StoredProcedure2
    {
        public string Start(string resultFilePath, string jsonString, string server, string basename)
        {

            string resultcheckFilePath = "c:\\temp\\~$result.xlsx";
            //resultFilePath = "c:\\temp\\result.xlsx";
            // Проверка, открыт ли файл result.xlsx
            if (File.Exists(resultcheckFilePath))
            {
                return "Файл result.xlsx открыт. Пожалуйста, закройте его и повторите попытку.";

            }


            // Создание экземпляра Workbook
            Workbook workbook = new Workbook();

            // Проверка заполнения ячейки A1
            if (string.IsNullOrEmpty(server))
            {
                return "Нету сервера, обычно это MMISLAB, но у вас должен быть доступ к серверу, он есть только у сотрудников ДОКО," +
                    " для обычных пользователей другая утилита";
            }
            if (string.IsNullOrEmpty(basename))
            {
                return "Нету базы данных, обычно это Абитуриенты или Деканат, но у вас должен быть доступ к серверу SQL, если его нет нажмите галочку" +
                    " вход через системную учетную запись";
            }


            if (string.IsNullOrEmpty(jsonString))
            {
                return "Нету данных JSOn:" +
                   "[{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\"," +
                   "\"params\":[{\"param\":\"date\", \"value\":\"20240401\"," +
                   "\"type\":\"string\"},{\"param\":\"trouble\", \"value\":\"0\",\"type\":\"int\"}],\"sheetnum\":2}," +
                   "{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\",\"param\":\"date\", \"value\":\"20240101\"," +
                   "\"type\":\"string\",\"sheetnum\":3}]";

            }

            // Парсинг JSON строки
            JArray json;
            try
            {
                json = JArray.Parse(jsonString);
            }
            catch (Exception ex)
            {
                return "Ошибка парсинга JSON: " + ex.Message + "Пример правильного содержания:" +
                    "[{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\"," +
                    "\"params\":[{\"param\":\"date\", \"value\":\"20240401\"," +
                    "\"type\":\"string\"},{\"param\":\"trouble\", \"value\":\"0\",\"type\":\"int\"}],\"sheetnum\":2}," +
                    "{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\",\"param\":\"date\", \"value\":\"20240101\"," +
                    "\"type\":\"string\",\"sheetnum\":3}]";
            }

            // Подключение к MSSQL
            string connectionString = "Data Source=" + server + ";Initial Catalog=" + basename + ";Integrated Security=True;Connect Timeout=30";


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch
                {
                    return "ошибка подключения к серверу " + server + " либо он неправильно введен либо недоступен ";
                }
                // Обработка каждой хранимой процедуры в JSON
                foreach (var query in json)
                {
                    string procedureName = query["procedure"].ToString();
                    int sheetNum = (int)query["sheetnum"];




                    // Создание команды для вызова хранимой процедуры
                    using (SqlCommand command = new SqlCommand(procedureName, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;

                        foreach (var param in query["params"])
                        {
                            string paramName = param["param"].ToString();
                            string paramValue = param["value"].ToString();
                            string paramType = param["type"].ToString();
                            // Проверка наличия обязательных параметров
                            if (string.IsNullOrEmpty(procedureName) || string.IsNullOrEmpty(paramName) || string.IsNullOrEmpty(paramValue) || string.IsNullOrEmpty(paramType) || sheetNum <= 0)
                            {
                                return "есть пустые значения";
                                // код поставлю из другого проекта
                            }
                            SqlParameter sqlParam = new SqlParameter(paramName, GetSqlDbType(paramType))
                            {
                                Value = paramValue
                            };
                            command.Parameters.Add(sqlParam);
                        }


                        // Проверка существования листа и создание, если его нет
                        Worksheet resultSheet;
                        if (sheetNum > workbook.Worksheets.Count)
                        {
                            resultSheet = workbook.Worksheets.Add($"Sheet{sheetNum}");
                        }
                        else
                        {
                            resultSheet = workbook.Worksheets[sheetNum - 1];
                        }

                        try
                        {
                            // Выполнение команды и получение результатов
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                int row = 1;

                                // Запись наименований полей в первую строку
                                for (int col = 0; col < reader.FieldCount; col++)
                                {
                                    resultSheet.Range[row, col + 1].Text = reader.GetName(col);
                                }
                                row++;
                                int currentRow = 0;
                                while (reader.Read())
                                {
                                    for (int col = 0; col < reader.FieldCount; col++)
                                    {
                                        resultSheet.Range[row, col + 1].Text = reader[col].ToString();
                                    }
                                    row++;
                                    currentRow++;
                                    
                                }
                            }

                        }
                        catch (SqlException ex)
                        {
                            return $"Извините, хранимой процедуры {procedureName} нет или у вас нет доступа: {ex.Message}";
                        }
                    }
                }

            }

            // Сохранение изменений в файл Excel
            workbook.SaveToFile(resultFilePath);
            return "Файл сохранен";
        }

        // Метод для получения SqlDbType из строки
        static SqlDbType GetSqlDbType(string type)
        {
            switch (type.ToLower())
            {
                case "int":
                    return SqlDbType.Int;
                case "string":
                    return SqlDbType.NVarChar;
                default:
                    throw new ArgumentException("Unsupported type");
            }
        }
    }
}
