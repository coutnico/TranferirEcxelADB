using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;

using System.Data.SqlClient;

namespace GestorExcelABasedeDatos
{
    public partial class Form1 : Form
    {
        int startRow = 6;
        int endRow = 10;
        int startColumn = 4;
        int endColumn = 4;
        static string ConexionString = "server= localhost, 1433 ; database= Auxiliar  ; user=sa; password=Verdeverde8";
        SqlConnection conexion = new SqlConnection(ConexionString);

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {


        }

        private void btnImportar_Click(object sender, EventArgs e)
        {


            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Archivos de Excel|*.xlsx|Archivos de Excel 97-2003|*.xls" })
            {
                try
                {
                    if (ofd.ShowDialog() == DialogResult.OK)
                    {
                        using (var stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read))
                        {
                            using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                            {
                                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration()
                                {
                                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                                    {


                                        UseHeaderRow = true
                                    }

                                });


                                var originalTable = dataSet.Tables[0];
                                var dataTable = new System.Data.DataTable();

                                // Agrega las columnas al nuevo DataTable
                                for (int j = 0; j < Math.Min(originalTable.Columns.Count, endColumn); j++)
                                {
                                    dataTable.Columns.Add(originalTable.Columns[j].ColumnName, originalTable.Columns[j].DataType);
                                }

                                // Agrega las filas omitiendo las primeras filas y columnas
                                for (int i = startRow - 1; i < Math.Min(originalTable.Rows.Count, endRow); i++)
                                {
                                    var newRow = dataTable.NewRow();

                                    for (int j = startColumn - 1; j < originalTable.Columns.Count; j++)
                                    {
                                        newRow[j - (startColumn - 1)] = originalTable.Rows[i][j];
                                    }

                                    dataTable.Rows.Add(newRow);
                                }

                                // Utilizar dataTable como desees...
                                dataGridView1.DataSource = dataTable;

                            }
                        }
                    }

                }
                catch (Exception)
                {

                    MessageBox.Show("Por favor cerrar el Excel");
                }
                dataGridView1.Columns[0].HeaderText = "Nombre";
                dataGridView1.Columns[1].HeaderText = "Apellido";
                dataGridView1.Columns[2].HeaderText = "Edad";
                dataGridView1.Columns[3].HeaderText = "Domicilio";



            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                conexion.Open();
                MessageBox.Show("Base de datos conectada con exito");
            }
            catch (Exception)
            {
                MessageBox.Show("La base de datos ya a sido conecatada");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                conexion.Close();
                MessageBox.Show("Base de datos desconectada con exito");
            }
            catch (Exception)
            {
                MessageBox.Show("La base de datos ya a sido desconecatada");
            }
        }

        private void btnExportar_Click(object sender, EventArgs e)
        {

            if (conexion.State == ConnectionState.Closed)
            {
                MessageBox.Show("No se realizo la conexion a base de datos");
                return;
            }

            try
            {
                string query = "INSERT INTO Datos ([Nombre], [Apellido], [Edad], [Domicilio]) VALUES (@valor1, @valor2, @valor3, @valor4)";
                SqlCommand comando = new SqlCommand(query, conexion);


                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    DataGridViewRow fila = dataGridView1.Rows[i];

                    // Limpiar parámetros de la iteración anterior
                    comando.Parameters.Clear();

                    // Asignar valores a los parámetros
                    comando.Parameters.AddWithValue("@valor1", fila.Cells[0].Value);
                    comando.Parameters.AddWithValue("@valor2", fila.Cells[1].Value);
                    comando.Parameters.AddWithValue("@valor3", fila.Cells[2].Value);
                    comando.Parameters.AddWithValue("@valor4", fila.Cells[3].Value);

                    // Ejecutar la inserción
                    comando.ExecuteNonQuery();
                }

                MessageBox.Show("Registros agregados con éxito");
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Hubo un error al agregar los registros: {ex.Message}");
            }
            finally
            {
                if (conexion.State == ConnectionState.Open)
                {
                    conexion.Close();
                }
            }
        }
    }
}
