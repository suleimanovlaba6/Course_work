using Marriage_registry.Marriage_registryDataSetTableAdapters;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Windows.Forms;

namespace Marriage_registry
{
    public partial class Form1 : Form
    {
        // З'єднання з базою даних
        private OleDbConnection connection;
        private const string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=D:\AIC\Marriage_registry\Marriage_registry\Marriage registry.accdb";

        public Form1()
        {
            InitializeComponent();
            // Ініціалізуємо з'єднання з базою даних
            connection = new OleDbConnection(connectionString);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: данная строка кода позволяет загрузить данные в таблицу "marriage_registryDataSet.Статистика". При необходимости она может быть перемещена или удалена.
            this.статистикаTableAdapter.Fill(this.marriage_registryDataSet.Статистика);
            // Завантаження даних у DataGridView при завантаженні форми
            LoadData();
            // Ініціалізуємо з'єднання з базою даних
            connection = new OleDbConnection(connectionString);
        }

        private void LoadData()
        {
            try
            {
                using (OleDbDataAdapter lostItemsAdapter = new OleDbDataAdapter("SELECT * FROM [Акти цивільного стану]", connection))
                {
                    DataTable lostItemsTable = new DataTable();
                    lostItemsAdapter.Fill(lostItemsTable);
                    dataGridView1.DataSource = lostItemsTable;
                }

                using (OleDbDataAdapter documentsAdapter = new OleDbDataAdapter("SELECT * FROM [Документи]", connection))
                {
                    DataTable documentsTable = new DataTable();
                    documentsAdapter.Fill(documentsTable);
                    dataGridView2.DataSource = documentsTable;
                }

                using (OleDbDataAdapter clientsAdapter = new OleDbDataAdapter("SELECT * FROM [Клієнти]", connection))
                {
                    DataTable clientsTable = new DataTable();
                    clientsAdapter.Fill(clientsTable);
                    dataGridView3.DataSource = clientsTable;
                }

                using (OleDbDataAdapter clientsAdapter = new OleDbDataAdapter("SELECT * FROM [Статистика]", connection))
                {
                    DataTable clientsTable = new DataTable();
                    clientsAdapter.Fill(clientsTable);
                    dataGridView4.DataSource = clientsTable;
                }
                // Встановлення автоматичного розширення стовпців для всіх DataGridView
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message);
            }
        }
        //////////////////////////PAGE-1///////////////////////////////////////
        private void CreateAct_Click(object sender, EventArgs e)
        {
            try
            {
                // Відкриваємо з'єднання з базою даних
                connection.Open();

                // Отримуємо дані з полів форми
                string pib = pib_page1.Text;
                DateTime registrationDate = dateTimePicker1.Value.Date;
                string actType = TypeAct_Page1.Text;
                string parentPib = string.IsNullOrEmpty(pib_perence_page1.Text) ? "-" : pib_perence_page1.Text;
                DateTime birthDate = dateTimePicker5.Value.Date; // Зберігаємо тільки дату
                string sex = sex_page1.Text;
                string address = adres_page1.Text;
                string passportIdText = idpass_page1.Text;

                // Перевірка, що ID паспорта містить 9 цифр
                if (passportIdText.Length != 9 || !passportIdText.All(char.IsDigit))
                {
                    MessageBox.Show("ID паспорта повинен містити 9 цифр.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                int passportId = int.Parse(passportIdText);

                // Перевірка, чи такий ID паспорту вже існує в базі даних
                string checkPassportIdQuery = "SELECT COUNT(*) FROM [Клієнти] WHERE [ID паспорта] = ?";
                using (OleDbCommand checkPassportIdCommand = new OleDbCommand(checkPassportIdQuery, connection))
                {
                    checkPassportIdCommand.Parameters.AddWithValue("@passportId", passportId);
                    int existingCount = (int)checkPassportIdCommand.ExecuteScalar();
                    if (existingCount > 0)
                    {
                        MessageBox.Show("Данні про ID паспорту вже є у базі даних. Зміни не внесено.", "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // Визначаємо тип документа на основі типу акту
                string documentType;
                switch (actType)
                {
                    case "Укладення шлюбу":
                        documentType = "Свідоцтво про брак";
                        break;
                    case "Отримання свідоцтва про розлучення":
                        documentType = "Довідка про розлучення";
                        break;
                    case "Отримання свідоцтва про народження":
                        documentType = "Свідоцтво про народження";
                        break;
                    case "Отримання свідоцтва про смерть":
                        documentType = "Довідка про смерть";
                        break;
                    default:
                        documentType = "Невідомий тип документа";
                        break;
                }

                // Додаємо дані до таблиці "Акти цивільного стану"
                string insertActQuery = "INSERT INTO [Акти цивільного стану] ([ПІБ Осіб зазначених у акті], [Дата реєстрації], [Тип акту], [ПІБ батьків]) VALUES (?, ?, ?, ?)";
                using (OleDbCommand command = new OleDbCommand(insertActQuery, connection))
                {
                    command.Parameters.AddWithValue("@pib", pib);
                    command.Parameters.AddWithValue("@registrationDate", registrationDate);
                    command.Parameters.AddWithValue("@actType", actType);
                    command.Parameters.AddWithValue("@parentPib", parentPib);
                    command.ExecuteNonQuery();
                }

                // Генеруємо унікальне 9-значне число для ID документа
                Random rand = new Random();
                int docId;
                bool isUnique;
                do
                {
                    docId = rand.Next(100000000, 999999999);

                    // Перевіряємо, чи існує вже такий ID в базі даних
                    string checkIdQuery = "SELECT COUNT(*) FROM [Документи] WHERE [ІD документа] = ?";
                    using (OleDbCommand checkCommand = new OleDbCommand(checkIdQuery, connection))
                    {
                        checkCommand.Parameters.AddWithValue("@docId", docId);
                        isUnique = (int)checkCommand.ExecuteScalar() == 0;
                    }
                } while (!isUnique);

                // Додаємо дані до таблиці "Документи"
                string insertDocQuery = "INSERT INTO [Документи] ([ПІБ особи зазначеної у документі], [Дата реєстрації запиту], [Тип документа], [Статус документа], [ІD документа]) VALUES (?, ?, ?, ?, ?)";
                using (OleDbCommand command = new OleDbCommand(insertDocQuery, connection))
                {
                    command.Parameters.AddWithValue("@pib", pib);
                    command.Parameters.AddWithValue("@registrationDate", registrationDate);
                    command.Parameters.AddWithValue("@documentType", documentType);
                    command.Parameters.AddWithValue("@status", "Не видано");
                    command.Parameters.AddWithValue("@docId", docId);
                    command.ExecuteNonQuery();
                }

                // Додаємо дані до таблиці "Клієнти"
                string insertClientQuery = "INSERT INTO [Клієнти] ([ПІБ], [Дата народження], [Стать], [Адреса], [ID паспорта]) VALUES (?, ?, ?, ?, ?)";
                using (OleDbCommand command = new OleDbCommand(insertClientQuery, connection))
                {
                    command.Parameters.AddWithValue("@pib", pib);
                    command.Parameters.AddWithValue("@birthDate", birthDate);
                    command.Parameters.AddWithValue("@sex", sex);
                    command.Parameters.AddWithValue("@address", address);
                    command.Parameters.AddWithValue("@passportId", passportId);
                    command.ExecuteNonQuery();
                }

                // Оновлюємо інтерфейс DataGridView
                LoadData();

                // Повідомлення про успішне занесення даних
                MessageBox.Show("Дані успішно занесено в таблиці.", "Успішно", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних у випадку помилки
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }
        //////////////////////////PAGE-2///////////////////////////////////////
        private void Search_page2_Click(object sender, EventArgs e)
        {
            try
            {
                // Відкриваємо з'єднання з базою даних
                connection.Open();

                // Отримуємо дані з полів форми
                string pib = pib_page2.Text;
                string parentPib = pib_perence_page2.Text;
                string actType = TypeAct_Page2.Text;
                string registrationDate = DateReg_page2.Text;

                // Формуємо SQL-запит для пошуку співпадінь
                string searchQuery = "SELECT * FROM [Акти цивільного стану] WHERE 1=1";

                // Додаємо умови до запиту, якщо відповідні поля заповнені
                if (!string.IsNullOrEmpty(pib))
                {
                    searchQuery += " AND [ПІБ Осіб зазначених у акті] = @pib";
                }

                if (!string.IsNullOrEmpty(parentPib))
                {
                    searchQuery += " AND [ПІБ батьків] = @parentPib";
                }

                if (!string.IsNullOrEmpty(actType))
                {
                    searchQuery += " AND [Тип акту] = @actType";
                }

                // Додаємо умову для дати, тільки якщо введено значення
                if (!string.IsNullOrEmpty(registrationDate))
                {
                    searchQuery += " AND [Дата реєстрації] = @registrationDate";
                }

                using (OleDbCommand command = new OleDbCommand(searchQuery, connection))
                {
                    if (!string.IsNullOrEmpty(pib))
                    {
                        command.Parameters.AddWithValue("@pib", pib);
                    }

                    if (!string.IsNullOrEmpty(parentPib))
                    {
                        command.Parameters.AddWithValue("@parentPib", parentPib);
                    }

                    if (!string.IsNullOrEmpty(actType))
                    {
                        command.Parameters.AddWithValue("@actType", actType);
                    }

                    // Додаємо параметр дати, тільки якщо введено значення
                    if (!string.IsNullOrEmpty(registrationDate))
                    {
                        command.Parameters.AddWithValue("@registrationDate", DateTime.Parse(registrationDate));
                    }

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable resultsTable = new DataTable();
                        adapter.Fill(resultsTable);

                        // Відображаємо результати пошуку у DataGridView
                        dataGridView1.DataSource = resultsTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }

        private void CancelSearch_page2_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void Delete_page2_Click(object sender, EventArgs e)
        {
            // Перевіряємо, чи є вибрані рядки в DataGridView
            if (dataGridView1.SelectedRows.Count > 0)
            {
                // Запитуємо користувача підтвердження видалення
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити обраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Отримуємо перший вибраний рядок
                    DataGridViewRow selectedRow = dataGridView1.SelectedRows[0];

                    // Отримуємо значення ID (або іншого унікального ідентифікатора) вибраного рядка
                    int id = Convert.ToInt32(selectedRow.Cells["iDАктуDataGridViewTextBoxColumn"].Value); // Припустимо, що у вас є стовпець з назвою "ID"

                    // SQL-запит DELETE
                    string query = "DELETE FROM [Акти цивільного стану] WHERE [ID акту] = @ID";

                    // Виконуємо SQL-запит
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        using (OleDbCommand command = new OleDbCommand(query, connection))
                        {
                            // Додаємо параметр для захисту від SQL-ін'єкцій
                            command.Parameters.AddWithValue("@ID", id);

                            // Відкриваємо підключення
                            connection.Open();

                            // Виконуємо команду
                            command.ExecuteNonQuery();
                        }
                    }

                    // Видаляємо вибраний рядок з DataGridView
                    dataGridView1.Rows.Remove(selectedRow);
                }
            }
            else
            {
                MessageBox.Show("Будь ласка, виберіть рядок для видалення.");
            }
        }
        //////////////////////////PAGE-3///////////////////////////////////////
        private void Search_page3_Click(object sender, EventArgs e)
        {
            try
            {
                // Відкриваємо з'єднання з базою даних
                connection.Open();

                // Отримуємо дані з полів форми
                string pib = pib_page3.Text;
                string idRec = IdRec_page3.Text;
                string docType = TypeDoc_Page3.Text;
                string registrationDate = DateReg_page3.Text;
                string docStat = StatusDoc_page3.Text;

                // Формуємо SQL-запит для пошуку співпадінь
                string searchQuery = "SELECT * FROM [Документи] WHERE 1=1";

                // Додаємо умови до запиту, якщо відповідні поля заповнені
                if (!string.IsNullOrEmpty(pib))
                {
                    searchQuery += " AND [ПІБ особи зазначеної у документі] = @pib";
                }

                if (!string.IsNullOrEmpty(idRec))
                {
                    searchQuery += " AND [ID Запису] = @idRec";
                }

                if (!string.IsNullOrEmpty(docType))
                {
                    searchQuery += " AND [Тип документа] = @docType";
                }

                if (!string.IsNullOrEmpty(docStat))
                {
                    searchQuery += " AND [Статус документа] = @docStat";
                }

                // Додаємо умову для дати, тільки якщо введено значення
                if (!string.IsNullOrEmpty(registrationDate))
                {
                    searchQuery += " AND [Дата реєстрації запиту] = @registrationDate";
                }

                using (OleDbCommand command = new OleDbCommand(searchQuery, connection))
                {
                    if (!string.IsNullOrEmpty(pib))
                    {
                        command.Parameters.AddWithValue("@pib", pib);
                    }

                    if (!string.IsNullOrEmpty(idRec))
                    {
                        command.Parameters.AddWithValue("@idRec", idRec);
                    }

                    if (!string.IsNullOrEmpty(docType))
                    {
                        command.Parameters.AddWithValue("@docType", docType);
                    }

                    if (!string.IsNullOrEmpty(docStat))
                    {
                        command.Parameters.AddWithValue("@docStat", docStat);
                    }

                    // Додаємо параметр дати, тільки якщо введено значення
                    if (!string.IsNullOrEmpty(registrationDate))
                    {
                        command.Parameters.AddWithValue("@registrationDate", registrationDate);
                    }

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable resultsTable = new DataTable();
                        adapter.Fill(resultsTable);

                        // Відображаємо результати пошуку у DataGridView
                        dataGridView2.DataSource = resultsTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }
        private void Edit_page3_Click_1(object sender, EventArgs e)
        {
            try
            {
                // Перевірка, чи вибрано хоча б один рядок для редагування
                if (dataGridView2.SelectedRows.Count > 0)
                {
                    // Перевірка, чи заповнено поле зі статусом документу
                    if (!string.IsNullOrEmpty(StatusDoc_page3.Text))
                    {
                        // Отримання значення, яке потрібно оновити
                        string newStatus = StatusDoc_page3.Text;

                        // Отримання індексу вибраного рядка
                        int selectedRowIndex = dataGridView2.SelectedRows[0].Index;

                        // Отримання ID рядка, який потрібно оновити
                        int selectedItemId = Convert.ToInt32(dataGridView2.Rows[selectedRowIndex].Cells["iDЗаписуDataGridViewTextBoxColumn"].Value);

                        // Оновлення стовбця "Статус документа" в таблиці "Документи"
                        string updateQuery = "UPDATE [Документи] SET [Статус документа] = ? WHERE [ID Запису] = ?";
                        using (OleDbCommand updateCommand = new OleDbCommand(updateQuery, connection))
                        {
                            updateCommand.Parameters.AddWithValue("@status", newStatus);
                            updateCommand.Parameters.AddWithValue("@id", selectedItemId);

                            // Відкриваємо з'єднання з базою даних
                            connection.Open();

                            int rowsAffected = updateCommand.ExecuteNonQuery();
                            if (rowsAffected > 0)
                            {
                                MessageBox.Show("Статус документа успішно оновлено!");
                                // Оновлення відображення таблиці
                                LoadData(); // Оновлення вмісту таблиці
                            }
                            else
                            {
                                MessageBox.Show("Не вдалося оновити статус документа.");
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("Будь ласка, оберіть статус документу.");
                    }
                }
                else
                {
                    MessageBox.Show("Будь ласка, виберіть рядок для редагування.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }
        private void CancelSearch_page3_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void Delete_page3_Click(object sender, EventArgs e)
        {
            // Перевіряємо, чи є вибрані рядки в DataGridView
            if (dataGridView2.SelectedRows.Count > 0)
            {
                // Запитуємо користувача підтвердження видалення
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити обраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Отримуємо перший вибраний рядок
                    DataGridViewRow selectedRow = dataGridView2.SelectedRows[0];

                    // Отримуємо значення ID (або іншого унікального ідентифікатора) вибраного рядка
                    int id = Convert.ToInt32(selectedRow.Cells["iDЗаписуDataGridViewTextBoxColumn"].Value); // Припустимо, що у вас є стовпець з назвою "ID"

                    // SQL-запит DELETE
                    string query = "DELETE FROM [Документи] WHERE [ID Запису] = @ID";

                    // Виконуємо SQL-запит
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        using (OleDbCommand command = new OleDbCommand(query, connection))
                        {
                            // Додаємо параметр для захисту від SQL-ін'єкцій
                            command.Parameters.AddWithValue("@ID", id);

                            // Відкриваємо підключення
                            connection.Open();

                            // Виконуємо команду
                            command.ExecuteNonQuery();
                        }
                    }

                    // Видаляємо вибраний рядок з DataGridView
                    dataGridView2.Rows.Remove(selectedRow);
                }
            }
        }

        //////////////////////////PAGE-4///////////////////////////////////////
        private void Search_page4_Click(object sender, EventArgs e)
        {
            try
            {
                // Відкриваємо з'єднання з базою даних
                connection.Open();

                // Отримуємо дані з полів форми
                string pib = pib_page4.Text;
                string birthday = BirthdayDate_page4.Text;
                string sex = sex_page4.Text;
                string idPassport = idpass_page4.Text;
                string address = adres_page4.Text;

                // Формуємо SQL-запит для пошуку співпадінь
                string searchQuery = "SELECT * FROM [Клієнти] WHERE 1=1";

                // Додаємо умови до запиту, якщо відповідні поля заповнені
                if (!string.IsNullOrEmpty(pib))
                {
                    searchQuery += " AND [ПІБ] = @pib";
                }

                if (!string.IsNullOrEmpty(birthday))
                {
                    searchQuery += " AND [Дата народження] = @birthday";
                }

                if (!string.IsNullOrEmpty(sex))
                {
                    searchQuery += " AND [Стать] = @sex";
                }

                if (!string.IsNullOrEmpty(idPassport))
                {
                    searchQuery += " AND [ID паспорта] = @idPassport";
                }

                if (!string.IsNullOrEmpty(address))
                {
                    searchQuery += " AND [Адреса] = @address";
                }

                using (OleDbCommand command = new OleDbCommand(searchQuery, connection))
                {
                    if (!string.IsNullOrEmpty(pib))
                    {
                        command.Parameters.AddWithValue("@pib", pib);
                    }

                    if (!string.IsNullOrEmpty(birthday))
                    {
                        command.Parameters.AddWithValue("@birthday", birthday);
                    }

                    if (!string.IsNullOrEmpty(sex))
                    {
                        command.Parameters.AddWithValue("@sex", sex);
                    }

                    if (!string.IsNullOrEmpty(idPassport))
                    {
                        command.Parameters.AddWithValue("@idPassport", idPassport);
                    }

                    if (!string.IsNullOrEmpty(address))
                    {
                        command.Parameters.AddWithValue("@address", address);
                    }

                    using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                    {
                        DataTable resultsTable = new DataTable();
                        adapter.Fill(resultsTable);

                        // Відображаємо результати пошуку у DataGridView
                        dataGridView3.DataSource = resultsTable;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }

        private void CancelSearch_page4_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void Delete_page4_Click(object sender, EventArgs e)
        {
            // Перевіряємо, чи є вибрані рядки в DataGridView
            if (dataGridView3.SelectedRows.Count > 0)
            {
                // Запитуємо користувача підтвердження видалення
                DialogResult result = MessageBox.Show("Ви впевнені, що хочете видалити обраний рядок?", "Підтвердження видалення", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (result == DialogResult.Yes)
                {
                    // Отримуємо перший вибраний рядок
                    DataGridViewRow selectedRow = dataGridView3.SelectedRows[0];

                    // Отримуємо значення ID (або іншого унікального ідентифікатора) вибраного рядка
                    int id = Convert.ToInt32(selectedRow.Cells["iDКлієнтаDataGridViewTextBoxColumn"].Value); // Припустимо, що у вас є стовпець з назвою "ID"

                    // SQL-запит DELETE
                    string query = "DELETE FROM [Клієнти] WHERE [ID клієнта] = @ID";

                    // Виконуємо SQL-запит
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        using (OleDbCommand command = new OleDbCommand(query, connection))
                        {
                            // Додаємо параметр для захисту від SQL-ін'єкцій
                            command.Parameters.AddWithValue("@ID", id);

                            // Відкриваємо підключення
                            connection.Open();

                            // Виконуємо команду
                            command.ExecuteNonQuery();
                        }
                    }

                    // Видаляємо вибраний рядок з DataGridView
                    dataGridView3.Rows.Remove(selectedRow);
                }
            }
        }
        //////////////////////////PAGE-5///////////////////////////////////////

        private void update_page5_Click(object sender, EventArgs e)
        {
            try
            {
                // Відкриваємо з'єднання з базою даних
                connection.Open();

                // Мапування між типами подій і їх назвами у базі даних
                Dictionary<string, string> eventTypeToActType = new Dictionary<string, string>
        {
            { "Укладення шлюбу", "Укладення шлюбу" },
            { "Отримання свідоцтва про народження", "Отримання свідоцтва про народження" },
            { "Отримання свідоцтва про розлучення", "Отримання свідоцтва про розлучення" },
            { "Отримання свідоцтва про смерть", "Отримання свідоцтва про смерть" }
        };

                // Мапування між типами документів і їх назвами у базі даних
                Dictionary<string, string> eventTypeToDocType = new Dictionary<string, string>
        {
            { "Укладення шлюбу", "Свідоцтво про брак" },
            { "Отримання свідоцтва про народження", "Свідоцтво про народження" },
            { "Отримання свідоцтва про розлучення", "Довідка про розлучення" },
            { "Отримання свідоцтва про смерть", "Довідка про смерть" }
        };

                // Запит для отримання кількості актів по кожному типу події
                string countActsQuery = @"
            SELECT [Тип акту], COUNT(*) AS [Кількість]
            FROM [Акти цивільного стану]
            GROUP BY [Тип акту]";
                Dictionary<string, int> actsCountDict = new Dictionary<string, int>();

                using (OleDbCommand countActsCommand = new OleDbCommand(countActsQuery, connection))
                {
                    using (OleDbDataReader reader = countActsCommand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string actType = reader["Тип акту"].ToString();
                            int count = Convert.ToInt32(reader["Кількість"]);
                            actsCountDict[actType] = count;
                        }
                    }
                }

                // Запит для отримання кількості виданих, невиданих і скасованих актів по кожному типу документа
                string countStatusQuery = @"
            SELECT [Тип документа], 
                   SUM(IIF([Статус документа] = 'Видано', 1, 0)) AS [Актів видано],
                   SUM(IIF([Статус документа] = 'Не видано', 1, 0)) AS [Актів не видано],
                   SUM(IIF([Статус документа] = 'Скасовано', 1, 0)) AS [Актів скасовано]
            FROM [Документи]
            GROUP BY [Тип документа]";
                Dictionary<string, (int issued, int notIssued, int canceled)> statusCountDict = new Dictionary<string, (int issued, int notIssued, int canceled)>();

                using (OleDbCommand countStatusCommand = new OleDbCommand(countStatusQuery, connection))
                {
                    using (OleDbDataReader reader = countStatusCommand.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            string docType = reader["Тип документа"].ToString();
                            int issuedCount = Convert.ToInt32(reader["Актів видано"]);
                            int notIssuedCount = Convert.ToInt32(reader["Актів не видано"]);
                            int canceledCount = Convert.ToInt32(reader["Актів скасовано"]);

                            statusCountDict[docType] = (issuedCount, notIssuedCount, canceledCount);
                        }
                    }
                }

                // Оновлюємо відповідні рядки у таблиці "Статистика"
                foreach (DataGridViewRow row in dataGridView4.Rows)
                {
                    string eventType = row.Cells["типПодіїDataGridViewTextBoxColumn"]?.Value?.ToString();
                    if (eventType != null)
                    {
                        // Оновлюємо колонку "Кількість актів"
                        if (eventTypeToActType.ContainsKey(eventType) && actsCountDict.ContainsKey(eventTypeToActType[eventType]))
                        {
                            row.Cells["кількістьАктівDataGridViewTextBoxColumn"].Value = actsCountDict[eventTypeToActType[eventType]];
                        }

                        // Оновлюємо колонки "Актів видано", "Актів не видано", "Актів скасовано"
                        if (eventTypeToDocType.ContainsKey(eventType) && statusCountDict.ContainsKey(eventTypeToDocType[eventType]))
                        {
                            var statusCounts = statusCountDict[eventTypeToDocType[eventType]];
                            row.Cells["актівВиданоDataGridViewTextBoxColumn"].Value = statusCounts.issued;
                            row.Cells["актівНеВиданоDataGridViewTextBoxColumn"].Value = statusCounts.notIssued;
                            row.Cells["актівСкасованоDataGridViewTextBoxColumn"].Value = statusCounts.canceled;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Помилка: " + ex.Message, "Помилка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Закриваємо з'єднання з базою даних
                if (connection.State == ConnectionState.Open)
                    connection.Close();
            }
        }
    }
}