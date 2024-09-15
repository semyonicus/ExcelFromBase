using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using Microsoft.CSharp.RuntimeBinder;

namespace ExcelFromBase
{
    public partial class Form2 : Form
    {
        static string _connectionString;
        private static string _login;
        private string _selectedFile;
        public string _selectedScenarioTable;

        public Form2(string connectionString, string login)
        {
            InitializeComponent();

            Text = "Соединение успешно";
            _connectionString = connectionString;
            _login = login;
            //   Workbook workbook = new Workbook();
            //   Worksheet sheet = workbook.Worksheets[0];
            //   sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            //   workbook.Save("output.xlsx");
            /*
            string procedureName = "uploadbase.dbo.ЗагрузкаРасписания";
            string[] columns;
            string fields = "result";
            // Создание команды для вызова хранимой процедуры
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                try
                {
                    using (SqlCommand command = new SqlCommand(procedureName, connection))
                    {
                        command.CommandType = CommandType.StoredProcedure;
                        try
                        {
                            connection.Open();
                            using (SqlDataReader reader = command.ExecuteReader())
                            {
                                columns = fields.Split(',');
                                while (reader.Read())
                                {
                                    foreach (var column in columns)
                                    {
                                        Text = reader[column].ToString();
                                    }
                                }
                                connection.Close();
                            }
                       

                        }
                        catch
                        {
                            MessageBox.Show("ошибка учетных данных не трогайте форму после соединения");
                            return;
                        }
                    }
                }
                catch
                {
                    MessageBox.Show("ошибка учетных данных  попробуйте другой вариант ");
                    return;
                }
            }*/
            ToolStripMenuItem loadMenuItem = new ToolStripMenuItem("Загрузка");
            loadMenuItem.Click += LoadMenuItem_Click;
            menuStrip1.Items.Add(loadMenuItem);


        }

        private void LoadMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    _selectedFile = openFileDialog.FileName;
                    // Обновляем Label с названием выбранного файла
                    foreach (Control control in this.Controls)
                    {
                        if (control is System.Windows.Forms.Label label)
                        {
                            label.Text = $"Выбранный файл: {_selectedFile}";
                            // Загружаем сценарии для пользователя
                            System.Windows.Forms.Label fileLabel = new System.Windows.Forms.Label { Location = new System.Drawing.Point(10, 30), AutoSize = true };
                            this.Controls.Add(fileLabel);
                            fileLabel.Visible = false;
                            // Добавляем ListBox для выбора сценария
                            System.Windows.Forms.ListBox scenarioListBox = new System.Windows.Forms.ListBox { 
                                Location = new System.Drawing.Point(10, 50), Size = new System.Drawing.Size(600, 100) };
                            this.Controls.Add(scenarioListBox);
                            //scenarioListBox.Visible = false;

                            // Добавляем кнопку StartScenario
                            System.Windows.Forms.Button startButton = new System.Windows.Forms.Button { Text = "Начать выбранный Сценарий", Location = new System.Drawing.Point(10, 170) };
                            startButton.Click += StartButton_Click;
                            this.Controls.Add(startButton);
                            //startButton.Visible = false;
                            LoadScenarios(scenarioListBox);

                        }
                    }
                }
            }
        }

        private void LoadScenarios(System.Windows.Forms.ListBox listBox)
        {
            string query = "EXEC YSUMain.dbo.ПоказатьСценарииДляПользователя @login";
            try
            {
                using (SqlConnection cnn = new SqlConnection(_connectionString))
                {
                    using (SqlCommand cmd = new SqlCommand(query, cnn))
                    {
                        cmd.Parameters.AddWithValue("@login", _login);
                        cnn.Open();
                        using (SqlDataReader reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string scenarioName = reader["ИмяСценария"].ToString();
                                //MessageBox.Show(scenarioName);
                                string scenarioTable = reader["Таблица"].ToString();
                                listBox.Items.Add(new { ScenarioName = scenarioName, ScenarioTable = scenarioTable });
                            }
                        }
                        cnn.Close();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ошибка { ex.Message}");
            }

            listBox.DisplayMember = "ScenarioName";
            listBox.ValueMember = "ScenarioTable";
            listBox.SelectedIndexChanged += (s, e) =>
            {
                _selectedScenarioTable = ((dynamic)listBox.SelectedItem).ScenarioTable;
            };
            if (listBox.Items.Count > 0)
            {
                listBox.SelectedIndex = 0;
            } else {
                MessageBox.Show("К сожалению у вас нет доступа ни к каким операциям");
                listBox.Visible = false;
            }

        }

        private void StartButton_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_selectedFile) && !string.IsNullOrEmpty(_selectedScenarioTable))
            {
                LoadExcelToDatabase(_selectedFile, _selectedScenarioTable);
            }
            else
            {
                MessageBox.Show("Пожалуйста, выберите файл и сценарий.");
            }
        }

        private void LoadExcelToDatabase(string filename, string tableName)
        {
            using (SqlConnection cnn = new SqlConnection(_connectionString))
            {
                // Создаем временную таблицу
                string createTableQuery = GenerateCreateTableQuery(tableName+_login, cnn);
                try
                {
                    using (SqlCommand cmd = new SqlCommand(createTableQuery, cnn))
                    {
                        cnn.Open();
                        cmd.ExecuteNonQuery();
                        cnn.Close();
                    }
                    MessageBox.Show($"таблица {tableName + _login} готова для загрузки");
                } catch (Exception ex)
                {
                    MessageBox.Show($"ошибка {ex.Message}");
                    return;
                }


                // Загружаем данные из Excel файла
                Workbook wb = new Workbook(filename);
                Worksheet worksheet = wb.Worksheets[0];
                DataTable dt = worksheet.Cells.ExportDataTable(0, 0, worksheet.Cells.MaxDataRow + 1, worksheet.Cells.MaxDataColumn + 1, true);

                // Сохраняем данные в временную таблицу
                using (SqlBulkCopy objbulk = new SqlBulkCopy(cnn))
                {
                    objbulk.DestinationTableName = "publicbase.dbo."+tableName + _login;

                    // Устанавливаем сопоставление столбцов
                    string mappingQuery = $"SELECT ПолеТаблицы, ПолеExcel FROM YSUMain.dbo.ПоляДляТаблиц where Таблица='{_selectedScenarioTable}'";
                    try
                    {
                        using (SqlCommand mappingCmd = new SqlCommand(mappingQuery, cnn))
                        {
                            cnn.Open();
                            using (SqlDataReader reader = mappingCmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    string fieldTable = reader["ПолеТаблицы"].ToString();
                                    string fieldExcel = reader["ПолеExcel"].ToString();
                                    objbulk.ColumnMappings.Add(fieldTable, fieldExcel);
                                    //MessageBox.Show($" {fieldTable} {fieldExcel}");
                                }
                            }
                            cnn.Close();
                        }
                    }
                    catch(Exception ex) {
                        MessageBox.Show($"ошибка {ex.Message}");
                        return;
                    }

                    cnn.Open();
                    try
                    {
                        objbulk.WriteToServer(dt);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ошибка {ex.Message}");
                        return;
                    }
                    MessageBox.Show($"Данные перенесены из {filename} {tableName + _login}");
                    cnn.Close();
                    string query = "EXEC [10.2.9.73].publicbase.dbo.ЗапускСценария @login,@table";
                    try
                    {
                            using (SqlCommand cmd = new SqlCommand(query, cnn))
                            {
                                cmd.Parameters.AddWithValue("@login", _login);
                                cmd.Parameters.AddWithValue("@table", tableName + _login);
                                cnn.Open();
                                using (SqlDataReader reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        string scenarioName = reader["ИмяСценария"].ToString();
                                        string scenarioTable = reader["Таблица"].ToString();
                                    }
                                }
                                cnn.Close();
                            }
                        
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"ошибка {ex.Message}");
                        return;
                    }


                    
                }
            }
        }

        private string GenerateCreateTableQuery(string tableName, SqlConnection cnn)
        {

            // удаляем
            string deleteTableQuery = $"select * from publicbase.dbo.{tableName}";
            try
            {
                cnn.Open();
                using (SqlCommand deleteCmd = new SqlCommand(deleteTableQuery, cnn))
                {
                    using (var reader = deleteCmd.ExecuteReader())
                    {
                        cnn.Close();
                        return $"delete from publicbase.dbo.{tableName}";
                    }
                }
            }
            catch
            {
                string query = $"SELECT ПолеТаблицы, ТипДанных FROM YSUMain.dbo.ПоляДляТаблиц where Таблица='{_selectedScenarioTable}'";
                string createTableQuery = $"CREATE TABLE publicbase.dbo.{tableName} (";
                if (cnn.State == System.Data.ConnectionState.Closed)
                {
                    cnn.Open();
                }
                using (SqlCommand cmd = new SqlCommand(query, cnn))
                {
                    using (SqlDataReader reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string field = reader["ПолеТаблицы"].ToString();
                            string dataType = reader["ТипДанных"].ToString();
                            createTableQuery += $"[{field}] {dataType},";
                        }
                    }
                }
                if (cnn.State == System.Data.ConnectionState.Open)
                {
                    cnn.Close();
                }
                createTableQuery = createTableQuery.TrimEnd(',') + ")";
                return createTableQuery;
            }
        }

    }
}
