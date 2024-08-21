using Newtonsoft.Json.Linq;
using ReadExcelExample;
using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Spire.Xls;
using System.Data.SqlClient;
using System.Security.Principal;
using System.Net.Sockets;
using System.Net;

namespace ExcelFromBase
{
    public partial class Form1 : Form
    {
        private Timer timer;
        private bool isBlinking;
        public string username = "";
        public Form1()
        {
            
            InitializeComponent();
            InitializeBackgroundWorker();
            InitializeBlinkingButton();
            //menuStrip1.Items.Clear();
            StartWork.Enabled = false;
            listBoxConfigs.Enabled = false;
            LoadChoiceToArea.Enabled = false;
        }
        private JObject configJson;
        private BackgroundWorker backgroundWorker;
        private void InitializeBackgroundWorker()
        {
#pragma warning disable IDE0017
            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.DoWork += new DoWorkEventHandler(BackgroundWorker_DoWork);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(BackgroundWorker_ProgressChanged);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BackgroundWorker_RunWorkerCompleted);
        }
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
        public static bool IsServerAvailable(string ipAddress, int port)
        {
            try
            {
                using (TcpClient client = new TcpClient())
                {
                    // Устанавливаем тайм-аут на 1 секунды
                    client.ReceiveTimeout = 1000;
                    client.SendTimeout = 1000;

                    // Пытаемся подключиться к серверу
                    client.Connect(ipAddress, port);
                    return true;
                }
            }
            catch (SocketException)
            {
                return false;
            }
        }
        private void Start(string resultFilePath, string jsonString, string server, string basename)
        {

            if (!IsServerAvailable(server, 1433)) {
                MessageBox.Show($" {server} сервер недоступен, вероятны порты закрыты или вы находитесь вне доступной сети");
                return;
            }

            string resultcheckFilePath = "c:\\temp\\~$result.xlsx";
            //resultFilePath = "c:\\temp\\result.xlsx";
            // Проверка, открыт ли файл result.xlsx
            if (File.Exists(resultcheckFilePath))
            {
                MessageBox.Show("Файл result.xlsx открыт. Пожалуйста, закройте его и повторите попытку.");
                return;

            }
            string filename;
            Workbook workbook = new Workbook();
            if (UseTemplate.Checked)
            {
                filename = TemplateFileName.Text;
                if (!File.Exists(filename))
                {
                    MessageBox.Show("Ошибка шаблон не найден");
                    return;

                }
                else
                {
                    workbook.LoadFromFile(filename);
                }
                
            } 
            // Создание экземпляра Workbook
            

            // Проверка заполнения ячейки A1
            if (string.IsNullOrEmpty(server))
            {
                MessageBox.Show("Нету сервера, обычно это MMISLAB, но у вас должен быть доступ к серверу, он есть только у сотрудников ДОКО," +
                    " для обычных пользователей другая утилита");
                return;
            }
            if (string.IsNullOrEmpty(basename))
            {
                MessageBox.Show("Нету базы данных, обычно это Абитуриенты или Деканат, но у вас должен быть доступ к серверу SQL, если его нет нажмите галочку" +
                    " вход через системную учетную запись");
                return;
            }


            if (string.IsNullOrEmpty(jsonString))
            {
                MessageBox.Show("Нету данных JSOn:" +
                   "[{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\"," +
                   "\"params\":[{\"param\":\"date\", \"value\":\"20240401\"," +
                   "\"type\":\"string\"},{\"param\":\"trouble\", \"value\":\"0\",\"type\":\"int\"}],\"sheetnum\":2}," +
                   "{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\",\"param\":\"date\", \"value\":\"20240101\"," +
                   "\"type\":\"string\",\"sheetnum\":3}]");
                return;

            }

            // Парсинг JSON строки
            JArray json;
            try
            {
                json = JArray.Parse(jsonString);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка парсинга JSON: " + ex.Message + "Пример правильного содержания:" +
                    "[{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\"," +
                    "\"params\":[{\"param\":\"date\", \"value\":\"20240401\"," +
                    "\"type\":\"string\"},{\"param\":\"trouble\", \"value\":\"0\",\"type\":\"int\"}],\"sheetnum\":2}," +
                    "{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\",\"param\":\"date\", \"value\":\"20240101\"," +
                    "\"type\":\"string\",\"sheetnum\":3}]");
                return;
            }

            // Подключение к MSSQL
            string connectionString;
            if (string.IsNullOrEmpty(LoginName.Text) || string.IsNullOrEmpty(Login2.Text))
            {
                connectionString = "Data Source=" + server + ";Initial Catalog=" + basename + ";Integrated Security=True;Connect Timeout=30";
                MessageBox.Show("Пробуем подключиться встроенной учетной записью если не получится придется ввести логин и пароль");
            }
            else
            {
                connectionString = $"Server ={server}; Database ={basename};User Id={LoginName.Text};Password={Login2.Text};";
                MessageBox.Show("Пробуем подключиться введенными учетными данными");
            }
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    connection.Open();
                }
                catch
                {
                    MessageBox.Show("ошибка учетных данных " + server + " попробуйте другой вариант ");
                    return;
                }
                string paramName;
                string paramValue;
                string paramType;
                string errors;
                // Обработка каждой хранимой процедуры в JSON
                foreach (var query in json)
                {
                    string procedureName = query["procedure"].ToString();
                    int sheetNum = (int)query["sheetnum"];
                    // Создание команды для вызова хранимой процедуры
                    using (SqlCommand command = new SqlCommand(procedureName, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        paramName = "";
                        paramValue = "";
                        paramValue = "";
                        errors = "";
                        foreach (var param in query["params"])
                        {
                            paramName = param["param"].ToString();
                            paramValue = param["value"].ToString();
                            paramType = param["type"].ToString();
                            
                            // Проверка наличия обязательных параметров
                            if (string.IsNullOrEmpty(procedureName) || string.IsNullOrEmpty(paramName) || string.IsNullOrEmpty(paramValue) || string.IsNullOrEmpty(paramType) || sheetNum <= 0)
                            {
                                if (string.IsNullOrEmpty(procedureName))
                                {
                                    errors+="Отсутствует обязательный параметр имя Процедуры в процедуре. ";
                                }
                                if (string.IsNullOrEmpty(paramName))
                                {
                                    errors += "Отсутствует обязательный параметр имя параметра - param.  ";
                                }
                                if (string.IsNullOrEmpty(paramValue))
                                {
                                    errors += "Отсутствует обязательный параметр значение параметра - value. ";
                                }
                                if (string.IsNullOrEmpty(paramType))
                                {
                                    errors += "Отсутствует обязательный параметр тип параметра - type, должен быть либо string либо int ";
                                }
                                if (sheetNum <= 0)
                                {
                                    errors += "Отсутствует обязательный параметр номер листа ";
                                }

                                errors += "Пример правильного содержания ячейки A1:" +
                                "[{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\"," +
                                "\"params\":[{\"param\":\"date\", \"value\":\"20240401\"," +
                                "\"type\":\"string\"},{\"param\":\"trouble\", \"value\":\"0\",\"type\":\"int\"}],\"sheetnum\":2}," +
                                "{\"procedure\":\"publicbase.dbo.GetStudentsbyDate\",\"param\":\"date\", \"value\":\"20240101\"," +
                                "\"type\":\"string\",\"sheetnum\":3}]";

                                MessageBox.Show("есть пустые значения" +errors);
                                return;
                                // код поставлю из другого проекта
                            }
                            SqlParameter sqlParam = new SqlParameter(paramName, GetSqlDbType(paramType))
                            {
                                Value = paramValue
                            };
                            command.Parameters.Add(sqlParam);
                        }
                        paramName = "username";
                        paramType = "string";
                        SqlParameter sqlUserName = new SqlParameter(paramName, GetSqlDbType(paramType))
                        {
                            Value = username
                        };
                        command.Parameters.Add(sqlUserName);

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
                                    backgroundWorker.ReportProgress(0, currentRow);

                                }
                            }

                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show($"Извините, хранимой процедуры {procedureName} нет или у вас нет доступа: {ex.Message}");
                            return;
                        }
                    }
                }

            }

            // Сохранение изменений в файл Excel
            workbook.SaveToFile(resultFilePath);
            return;
        }
        private void BackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string resultFilePath = ResultFile.Text;
            string jsonString = QueryArea.Text;
            string server = ServerName.Text;
            string basename = DataBaseName.Text;
            Start(resultFilePath, jsonString, server, basename);
        }
        private void BackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            StatusText.Text = $"Обработано строк: {e.UserState}";
        }

        private void BackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            StatusText.Text = "Загрузка завершена.";
            StartWork.Enabled = true;
        }
        private void InitializeBlinkingButton()
        {
            timer = new Timer();
            {
                timer.Interval = 500; // Интервал в миллисекундах (500 мс = 0.5 сек)
                timer.Tick += Timer_Tick;
                timer.Start();
            }

            LoadSettings.Click += LoadSettings_Click;
        }
        private void Timer_Stop(object sender, EventArgs e)
        {
            if (!DisableMessages.Checked)
            {
                timer.Stop();
            }
        }
        private void Timer_Tick(object sender, EventArgs e)
        {
            if (isBlinking)
            {
                LoadSettings.BackColor = Color.White;
            }
            else
            {
                LoadSettings.BackColor = Color.Green;
            }
            isBlinking = !isBlinking;
        }

        private void StartWork_Click(object sender, EventArgs e)
        {
            
            
            if (!backgroundWorker.IsBusy)
            {
                StatusText.Text = "Подождите процесс начался он может быть длительным до 30 сек";
                StartWork.Enabled = false;
                backgroundWorker.RunWorkerAsync();
            }
            
            

        }

        private void LoadSettings_Click(object sender, EventArgs e)
        {
            string filePath = PresetData.Text;

            if (!File.Exists(filePath))
            {
                MessageBox.Show("Файл config.json не найден.");
                return;
            }

            try
            {
                string jsonContent = File.ReadAllText(filePath);
                configJson = JObject.Parse(jsonContent);

                if (configJson["version"] == null)
                {
                    MessageBox.Show("Неверный формат JSON файла: отсутствует ключ 'version'.");
                    return;
                }
                listBoxConfigs.Items.Clear();
                foreach (var config in configJson["configs"])
                {
                    listBoxConfigs.Items.Add(config["name"].ToString());
                }
                LoadChoiceToArea.Enabled = true;
                listBoxConfigs.Enabled = true;
                if (!DisableMessages.Checked)
                {
                    timer.Stop();
                }
                LoadSettings.BackColor = SystemColors.Control; // Возвращаем цвет по умолчанию
                if (!DisableMessages.Checked)
                {
                    MessageBox.Show("Конфигурация загружена");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при чтении JSON файла: {ex.Message}");
            }
        }

        private void LoadChoiceToArea_Click(object sender, EventArgs e)
        {
            if (listBoxConfigs.SelectedItem == null)
            {
                MessageBox.Show("Пожалуйста, выберите конфигурацию.");
                return;
            }

            string selectedConfigName = listBoxConfigs.SelectedItem.ToString();
            var selectedConfig = configJson["configs"].First(config => config["name"].ToString() == selectedConfigName);

            QueryArea.Text = selectedConfig["data"].ToString();
            if (!DisableMessages.Checked)
            {
                MessageBox.Show("Конфигурация выбрана описание зелеными буквами.Нажимайте кнопку Шаг 3");
            }
        }

        private void ListBoxConfigs_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listBoxConfigs.SelectedItem == null)
            {
                StatusText.Text = "";
                return;
            }

            string selectedConfigName = listBoxConfigs.SelectedItem.ToString();
            var selectedConfig = configJson["configs"].First(config => config["name"].ToString() == selectedConfigName);

            StatusText.Text = selectedConfig["desc"].ToString();
            StartWork.Enabled = true;
        }


        private void Form1_Load(object sender, EventArgs e)
        {
            string domainname="";
            
            string[] temp = Convert.ToString(WindowsIdentity.GetCurrent().Name).Split('\\');
            if (!string.IsNullOrEmpty(temp[1]))
            {
                username = temp[1];
            } else
            {
                MessageBox.Show("Ошибка определения имени");
            }
            if (!string.IsNullOrEmpty(temp[0])) {
                domainname = temp[0];
            } else
            {
                MessageBox.Show("Ошибка определения домена");
            }
            int start = 0;
            if (domainname == "DESKTOP-SUU5GSR") start++;
            if (domainname == "S-VFU") start++;
            if (domainname == "DESKTOP-UERLFKI") start++;
            if (domainname == "DESKTOP - UERLFKI") start++;
            if (start>0)
            {

                UserNameTxt.Text = "У вас есть доступ :" + username;
            } else
            {
                MessageBox.Show("Доступа нет все функции заблокированы");
                UserNameTxt.Text = domainname+username+" у вас нет доступа, так вы не находитесь в домене S-VFU, но при работе  с базой еще будет проверка на доступ к сценариям";
                LoadSettings.Enabled = false;
                timer.Stop();
            }

        }

        private void UseTemplate_CheckedChanged(object sender, EventArgs e)
        {
            string filename = TemplateFileName.Text;
            if (!File.Exists(filename))
            {
                MessageBox.Show("Ошибка шаблон не найден");
                return;

            }
        }

        private void UploadExcel_Click(object sender, EventArgs e)
        {
            string server = ServerName.Text;
            string basename = DataBaseName.Text;
            string login = LoginName.Text;
            string passw = Login2.Text;

            if (string.IsNullOrEmpty(login) || string.IsNullOrEmpty(passw) || string.IsNullOrEmpty(server) || string.IsNullOrEmpty(basename))
            {
                MessageBox.Show("Для работы с этим надо обязательно ввести данные ИИСУСС пользователя, сервер и базу");
                return;
            }
            if (!IsServerAvailable(server, 1433))
            {
                MessageBox.Show($" {server} сервер недоступен, вероятны порты закрыты или вы находитесь вне доступной сети");
                return;
            }

            string connectionString = $"Server ={server}; Database ={basename};User Id={login};Password={passw};";

                try
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                    // Ваш код для работы с базой данных здесь
                    //MessageBox.Show("Подключение успешно установлено.");
                    Form2 form2 = new Form2(connectionString,login);
                    // Показываем Form2
                    form2.Show();
                }
                }
                catch (SqlException ex)
                {
                MessageBox.Show("Ошибка при подключении к базе данных: " + ex.Message);
                return;
                }
                catch (InvalidOperationException ex)
                {
                MessageBox.Show("Ошибка: " + ex.Message);
                return;
            }
                catch (Exception ex)
                {
                MessageBox.Show("Произошла ошибка: " + ex.Message);
                return;
            }
                
        }
    }
}
