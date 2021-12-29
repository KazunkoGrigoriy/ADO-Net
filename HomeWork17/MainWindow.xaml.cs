using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Configuration;
using System.Data.OleDb;
using System.Data;

namespace HomeWork17
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        
        public MainWindow() { InitializeComponent(); Preparing(); }

        OleDbDataAdapter Oleda;
        SqlDataAdapter SQLda;
        DataTable SQLdt, Oledt;
        DataRowView row;

        private void Preparing()
        {
            SqlConnectionStringBuilder strConSQL = new SqlConnectionStringBuilder()
            {
                DataSource = @"(localdb)\MSSQLLocalDB",
                InitialCatalog = "SQLbase",
                IntegratedSecurity = false,
                UserID = "username",
                Password = "Ytpyf.24017"
            };
            SqlConnection sqlConnection = new SqlConnection() { ConnectionString = strConSQL.ConnectionString };
            try
            {
                sqlConnection.Open();
                SQLda = new SqlDataAdapter();
                SQLdt = new DataTable();

                #region select

                var sql = @"SELECT * FROM TableClient Order by TableCLient.Id";

                SQLda.SelectCommand = new SqlCommand(sql, sqlConnection);

                #endregion

                #region insert

                sql = @"INSERT INTO TableClient(LastName, FirstName, MiddleName, NumberTel, Email)
                                     VALUES(@LastName, @FirstName, @MiddleName, @NumberTel, @Email)
                   SET @Id=@@IDENTITY";

                SQLda.InsertCommand = new SqlCommand(sql, sqlConnection);

                SQLda.InsertCommand.Parameters.Add("@Id", SqlDbType.Int, 4, "Id");
                SQLda.InsertCommand.Parameters.Add("@LastName", SqlDbType.NVarChar, 50, "LastName");
                SQLda.InsertCommand.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50, "FirstName");
                SQLda.InsertCommand.Parameters.Add("@MiddleName", SqlDbType.NVarChar, 50, "MiddleName");
                SQLda.InsertCommand.Parameters.Add("@NumberTel", SqlDbType.NVarChar, 50, "NumberTel");
                SQLda.InsertCommand.Parameters.Add("@Email", SqlDbType.NVarChar, 50, "Email");

                #endregion

                #region update


                sql = @"UPDATE TableClient SET
                           LastName = @LastName,
                           FirstName = @FirstName,
                           MiddleName = @MiddleName,
                           NumberTel = @NumberTel,
                           Email = @Email
                   WHERE Id = @Id";

                SQLda.UpdateCommand = new SqlCommand(sql, sqlConnection);

                SQLda.UpdateCommand.Parameters.Add("@Id", SqlDbType.Int, 4, "Id");
                SQLda.UpdateCommand.Parameters.Add("@LastName", SqlDbType.NVarChar, 50, "LastName");
                SQLda.UpdateCommand.Parameters.Add("@FirstName", SqlDbType.NVarChar, 50, "FirstName");
                SQLda.UpdateCommand.Parameters.Add("@MiddleName", SqlDbType.NVarChar, 50, "MiddleName");
                SQLda.UpdateCommand.Parameters.Add("@NumberTel", SqlDbType.NVarChar, 50, "NumberTel");
                SQLda.UpdateCommand.Parameters.Add("@Email", SqlDbType.NVarChar, 50, "Email");


                #endregion

                #region delete

                sql = @"DELETE FROM TableClient WHERE Id = @Id";

                SQLda.DeleteCommand = new SqlCommand(sql, sqlConnection);

                SQLda.DeleteCommand.Parameters.Add("@Id", SqlDbType.Int, 4, "Id");

                #endregion

                SQLda.Fill(SQLdt);

                dataGrid.ItemsSource = SQLdt.DefaultView;

                sqlConnection.Close();

            }
            catch (Exception e)
            {
                MessageBox.Show("К базе данных SQL подключиться не удалось");
            }        
        }
        
        private void dataGrid_AddingNewItem(object sender, AddingNewItemEventArgs e)
        {
            DataRow r = SQLdt.NewRow();
            SQLdt.Rows.Add(r);
            SQLda.Update(SQLdt);
        }

        private void dataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            row = (DataRowView)dataGrid.SelectedItem;
            row.BeginEdit();          
        }

        private void dataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            if (row == null) return;
            row.EndEdit();
            SQLda.Update(SQLdt);
        }

        private void dataGrid_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                row = (DataRowView)dataGrid.SelectedItem;
                SQLda.Update(SQLdt);
                //MessageBox.Show("Удалено");
            }

            if (e.Key == Key.Right)
            {
                SqlConnectionStringBuilder strConSQL = new SqlConnectionStringBuilder()
                {
                    DataSource = @"(localdb)\MSSQLLocalDB",
                    InitialCatalog = "SQLbase",
                    IntegratedSecurity = false,
                    UserID = "username",
                    Password = "Ytpyf.24017"
                };
                SqlConnection sqlConnection = new SqlConnection() { ConnectionString = strConSQL.ConnectionString };

                sqlConnection.Open();

                string path = Environment.CurrentDirectory + "\\database.mdb;";
                OleDbConnection strConAccess = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path);
                OleDbConnection accessConnection = new OleDbConnection() { ConnectionString = strConAccess.ConnectionString };
                try
                {
                    accessConnection.Open();
                    Oleda = new OleDbDataAdapter();
                    Oledt = new DataTable();
                    DataRowView row = (DataRowView)dataGrid.SelectedItems[0];
                    var r = row["Email"];
                    var ole = "SELECT * FROM TablePurchase WHERE Email = '" + r + "'";

                    Oleda.SelectCommand = new OleDbCommand(ole, accessConnection);

                    Oleda.Fill(Oledt);
                    dataGrid1.ItemsSource = Oledt.DefaultView;
                    accessConnection.Close();
                }
                catch (Exception a)
                {
                    MessageBox.Show("К базе данных MS Access подключиться не удалось");
                }
                sqlConnection.Close();
            }
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            DataRow r = SQLdt.NewRow();
            FormAdd formAdd = new FormAdd(r);
            formAdd.ShowDialog();
            if(formAdd.DialogResult.Value)
            {
                SQLdt.Rows.Add(r);
                SQLda.Update(SQLdt);
            }
        }
    }
}
