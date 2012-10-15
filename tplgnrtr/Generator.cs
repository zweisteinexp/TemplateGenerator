using System;
using System.Collections.Generic;

using System.Text;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using System.Data;
using System.Data.OleDb;

namespace tplgnrtr
{
    class Generator
    {
        private string sourcePath = null;
        private string sqlPath = null;
        private string asPath = null;

        private string historyPath = AppDomain.CurrentDomain.BaseDirectory + "\\history";

        private Dictionary<string, long> lastModifiedTime;

        public Generator(string pTplPath, string pSqlPath, string pAsPath)
        {
            sourcePath = pTplPath;
            sqlPath = pSqlPath;// +"sql\\";
            asPath = pAsPath;
        }
        private void loadHistory()
        {
            BinaryFormatter formatter = new BinaryFormatter();

            if (File.Exists(historyPath))
            {
                StreamReader sr = new StreamReader(historyPath);
                lastModifiedTime = (Dictionary<string, long>)formatter.Deserialize(sr.BaseStream);
                sr.Close();
            }
            else
            {
                lastModifiedTime = new Dictionary<string, long>();
            }
        }

        /// <summary>
        /// Save the LastWriteTime of all the excel files.
        /// </summary>
        private void saveHistory()
        {
            BinaryFormatter formatter = new BinaryFormatter();
            StreamWriter sw = new StreamWriter(historyPath);

            formatter.Serialize(sw.BaseStream, lastModifiedTime);
            sw.Close();
        }
		
		private void _traverseDirectory( string path )
		{
			// traverse sub directories
			string[] dir = System.IO.Directory.GetDirectories(path);

			for (int i = 0; i < dir.Length; i++)
			{
				_traverseDirectory( dir[i] );
			}
			
			// traverse files
			string[] files = System.IO.Directory.GetFiles(path, "*.xlsx");

			for (int i = 0; i < files.Length; i++)
			{
				int k = files[i].LastIndexOf("\\");
				string tableName = files[i].Substring(k + 1, files[i].Length - k - 6);
				
				// ignore temperary file
				if (tableName.IndexOf("~$") == 0)
				{
					continue;
				}

				long currentModifiedtime = File.GetLastWriteTimeUtc(files[i]).Ticks;

				if (!lastModifiedTime.ContainsKey(tableName) || currentModifiedtime != lastModifiedTime[tableName])
				{
					ImportOneTemplate(path + "\\", tableName);
				}

				lastModifiedTime[tableName] = currentModifiedtime;
			}
		}

        public void go()
        {
            // for avoiding crash of Excel
            System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US");

            // read the generation record of last time to see whether the Excel files have been updated.
            // if not, skip generation this time.
            loadHistory();

            try
            {
                logStatus("Starting...");

				_traverseDirectory( sourcePath );

                saveHistory();

                // Call import.bat to generate SQLite DB
                //generateSQLiteDB();

                logStatus("All Done.");
                //MessageBox.Show("All Done Successfully.");
            }
            catch (Exception ex)
            {
                ReportError(ex);
            }
            finally
            {

            }
        }

        /// <summary>
        /// Read an Excel file and generation corresponding template files.
        /// </summary>
        /// <param name="sourcePath">path the folder of the Excel file</param>
        /// <param name="tableName">File name without extension</param>
        private void ImportOneTemplate(string sourcePath, string tableName)
        {
            OleDbConnection m_XLSConn = null;
            OleDbDataAdapter m_XLSAdapter;
            DataTable m_dtXLS = null;

            try
            {
                logStatus("Exporting " + tableName + "...");

                string currentSourcePath = sourcePath + tableName + ".xlsx";

                string s_ConnString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source = " + currentSourcePath + ";Extended Properties ='Excel 12.0;IMEX=1;HDR=NO;'";
                string s_SQLSelect = "select * from [Main$]";
                m_XLSConn = new OleDbConnection(s_ConnString);
                m_XLSAdapter = new OleDbDataAdapter(s_SQLSelect, m_XLSConn);

                m_dtXLS = new DataTable();
                m_XLSConn.Open();
                m_XLSAdapter.Fill(m_dtXLS);


                GenerateSQL(m_dtXLS, tableName);
                GenerateActionScript(m_dtXLS, tableName);

                logStatus("   " + tableName + " exported.");
            }
            catch (Exception ex)
            {
                ReportError(ex);
            }
            finally
            {
                if (m_XLSConn != null)
                    m_XLSConn.Close();
            }
        }

        private void GenerateActionScript(DataTable templateData, string tableName)
        {
            StringBuilder asResults = new StringBuilder();

            logStatus("   Generating ActionScript statements for " + tableName + "...");

            // get info about which columns need be exported
            int flagRow = 2;

            IList<bool> needExport = new List<bool>();
            string value;
            string purifiedValue;
            bool hasColumns = false;

            for (int i = 0; i < templateData.Columns.Count; i++)
            {
                if (templateData.Rows[flagRow].IsNull(i))
                {
                    throw new Exception("Importation flag is not specified at column " + i + " in table " + tableName + ". Column Count: " + templateData.Columns.Count);
                }

                value = (string)templateData.Rows[flagRow][i];
                if ("Both".Equals(value) || "Client".Equals(value))
                {
                    needExport.Add(true);
                    hasColumns = true;
                }
                else
                {
                    needExport.Add(false);
                }
            }

            if (!hasColumns)
            {
                return;
            }


            // export column head
            StringBuilder valueString = new StringBuilder();

            for (int j = 0; j < templateData.Columns.Count; j++)
            {
                if (!needExport[j])
                {
                    continue;
                }

                if (templateData.Rows[3].IsNull(j))
                {
                    throw new Exception("Field name is not specified at column " + j + " in table " + tableName);
                }

                value = (string)templateData.Rows[3][j];
                valueString.Append(" \"" + value + "\" : \"" + value + "\",");
            }
            asResults.AppendLine("name[\"" + tableName + "\"] = {" + valueString.Remove(valueString.Length - 1, 1).ToString() + " };");
            asResults.AppendLine("value[\"" + tableName + "\"] = {");

            // export content
            for (int i = 4; i < templateData.Rows.Count; i++)
            {
                if (templateData.Rows[i].IsNull(0))
                {
                    throw new Exception("id is not specified at row " + (i + 1) + " in table " + tableName);
                }

                valueString = new StringBuilder();

                for (int j = 0; j < templateData.Columns.Count; j++)
                {
                    if (!needExport[j])
                    {
                        continue;
                    }

                    value = templateData.Rows[i].IsNull(j) ? "" : (string)templateData.Rows[i][j];

                    purifiedValue = value.Trim();
                    purifiedValue = purifiedValue.Replace("\n", "");

                    if (value != purifiedValue)
                    {
                        //logStatus("     WARNING: space or line break exists in row " + i + ", column " + j);
                    }

                    valueString.Append(" \"" + (string)templateData.Rows[3][j] + "\" : \"" + purifiedValue + "\",");
                }

                asResults.AppendLine("\t\"" + (string)templateData.Rows[i][0] + "\" : {" + valueString.Remove(valueString.Length - 1, 1).ToString() + " },");
                //txtLua.Text = i.ToString();
                //Application.DoEvents();
            }

            asResults.Remove(asResults.Length - 3, 3);
            asResults.AppendLine("");
            asResults.AppendLine("}");

            string currentAsPath = asPath + tableName + ".as";
            this.writefile(asResults, currentAsPath);
        }

        private void GenerateSQL(DataTable templateData, string tableName)
        {
            StringBuilder sqlResults = new StringBuilder();

            logStatus("   Generating SQL statements for " + tableName + "...");

            sqlResults.AppendLine("BEGIN TRANSACTION;");
            sqlResults.AppendLine("DROP TABLE IF EXISTS " + tableName + ";");

            // get info about which columns need be exported
            int flagRow = 2;

            IList<bool> needExport = new List<bool>();
            string value;
            string purifiedValue;
            bool hasColumns = false;

            for (int i = 0; i < templateData.Columns.Count; i++)
            {
                if (templateData.Rows[flagRow].IsNull(i))
                {
                    throw new Exception("Importation flag is not specified at column " + i + " in table " + tableName + ". Column Count: " + templateData.Columns.Count);
                }

                value = (string)templateData.Rows[flagRow][i];
                if ("Both".Equals(value) || "Server".Equals(value))
                {
                    needExport.Add(true);
                    hasColumns = true;
                }
                else
                {
                    needExport.Add(false);
                }
            }

            if (!hasColumns)
            {
                return;
            }

            // export column head
            StringBuilder valueString = new StringBuilder();

            for (int j = 0; j < templateData.Columns.Count; j++)
            {
                if (!needExport[j])
                {
                    continue;
                }

                if (templateData.Rows[3].IsNull(j))
                {
                    throw new Exception("Field name is not specified at column " + j + " in table " + tableName);
                }

                value = (string)templateData.Rows[3][j];

                if (j == 0)
                {
                    valueString.Append("\"" + value + "\" INTEGER PRIMARY KEY ASC,");
                }
                else
                {
                    valueString.Append("\"" + value + "\",");
                }
            }
            sqlResults.AppendLine("CREATE TABLE " + tableName + "(" + valueString.Remove(valueString.Length - 1, 1).ToString() + ");");

            // export content
            for (int i = 4; i < templateData.Rows.Count; i++)
            {
                if (templateData.Rows[i].IsNull(0))
                {
                    throw new Exception("id is not specified at row " + (i + 1) + " in table " + tableName);
                }

                valueString = new StringBuilder();

                for (int j = 0; j < templateData.Columns.Count; j++)
                {
                    if (!needExport[j])
                    {
                        continue;
                    }

                    value = templateData.Rows[i].IsNull(j) ? "" : (string)templateData.Rows[i][j];

                    purifiedValue = value.Trim();
                    purifiedValue = purifiedValue.Replace("\n", "");

                    if (value != purifiedValue)
                    {
                        //logStatus("     WARNING: space or line break exists in row " + i + ", column " + j);
                    }

                    valueString.Append("\"" + purifiedValue + "\",");
                }

                sqlResults.AppendLine("INSERT INTO " + tableName + " VALUES (" + valueString.Remove(valueString.Length - 1, 1).ToString() + ");");
                //txtLua.Text = i.ToString();
                //Application.DoEvents();
            }

            //sql文件结束
            sqlResults.AppendLine("COMMIT;");

            //导出sql文件
            string currentSqlPath = sqlPath + tableName + ".sql";
            this.writefile(sqlResults, currentSqlPath);
        }

        //#region 写文件
        public void writefile(StringBuilder sqlStr, string sqlPath)
        {
            // StreamWriter sw = new StreamWriter("d:\\TestFile.txt", true, System.Text.Encoding.GetEncoding("GB2312"));
            using (StreamWriter sw = new StreamWriter(sqlPath))
            {
                // Add some text to the file.\
                sw.Write(sqlStr);
            }
        }

        private void logStatus(string message)
        {
            Console.WriteLine(" - " + message);
        }

        private void ReportError(Exception ex)
        {
            Console.WriteLine("!!!ERROR: " + ex.Message);
            Environment.Exit(1);
        }

    }
}
