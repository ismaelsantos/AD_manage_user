using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace ADmanageUser.Class
{
    class Csv
    {
        /// <summary>
        /// Gera arquivo csv apartir de um DataGridView.
        /// </summary>
        /// <param name="dgvDados">DataGridView</param>
        /// <returns>Caminho do arquivo</returns>
        public static string GenerateFormatCsv(DataGridView dgvDados)
        {
            string lineColumn = FormatCsv(dgvDados);
            return SaveCsvFile(lineColumn);
        }

        private static string SaveCsvFile(string lineColumn)
        {
            SaveFileDialog saveFile = new SaveFileDialog();
            saveFile.Filter = "Arquivo exportado | *.csv";
            saveFile.FileName = "Script exportação";

            if (saveFile.ShowDialog() == DialogResult.OK)
            {
                StreamWriter arquivo = new StreamWriter(saveFile.FileName);
                arquivo.WriteLine(lineColumn);
                arquivo.Close();
            }

            return saveFile.FileName;
        }

        private static string FormatCsv(DataGridView dgvDados)
        {
            string lineColumn = "";
            try
            {
                /*//Código para capturar somente o nome das colunas
                string linha = "";

                for (int i = 0; i < dgvDados.Columns.Count; i++)
                    linha += dgvDados.Columns[i].HeaderText + ";";

                arquivo.WriteLine(linha);//*/

                //Código para capturar os dados
                int semicolonCount = 0;
                for (int i = 0; i < dgvDados.Rows.Count - 1; i++)
                {
                    for (int j = 0; j < dgvDados.Columns.Count; j++)
                    {
                        if (semicolonCount < (dgvDados.Columns.Count - 1))
                            lineColumn += dgvDados.Rows[i].Cells[j].Value + ";";
                        else
                            lineColumn += dgvDados.Rows[i].Cells[j].Value;

                        semicolonCount++;
                    }

                    semicolonCount = 0;
                    lineColumn += "\n";
                }
            }
            catch (NullReferenceException error)
            {
                return error.Message;
            }

            return lineColumn;
        }
    }

}
