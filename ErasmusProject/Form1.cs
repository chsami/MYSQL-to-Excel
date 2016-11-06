using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using MySql.Data.MySqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Extensions;
using DocumentFormat.OpenXml;
using System.Threading;

namespace ErasmusProject
{
    public partial class ErasmusProject : Form
    {

        private static MySqlConnection connection;
        private static string server;
        private static string database;
        private static string uid;
        private static string password;
        static ListBox listbox;
        static ProgressBar progressBar;
        private static string fileLocation = "";


        public ErasmusProject()
        {
            InitializeComponent();

        }

        #region excel methods
        private void addExcelFieldValue(String[] text, char[] cellName)
        {
            using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Open(fileLocation, true))
            {
                // Get the SharedStringTablePart. If it does not exist, create a new one.
                SharedStringTablePart shareStringPart;
                if (spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                {
                    shareStringPart = spreadSheet.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                }
                else
                {
                    shareStringPart = spreadSheet.WorkbookPart.AddNewPart<SharedStringTablePart>();
                }
                int index = 0;
                Cell cell = null;
                // Insert a new worksheet.
                WorksheetPart worksheetPart = GetWorksheetPart(spreadSheet.WorkbookPart, "Blad1");
                uint fieldNumber = 3;
                DateTime startTime = DateTime.Now;
                for (uint i = 0; i < text.Length; i++)
                {
                    if (i % 9 == 0 && i != 0)
                    {
                        fieldNumber++;
                        Invoke(new Action(() =>
                        {
                            TimeSpan timeRemaining = TimeSpan.FromTicks(DateTime.Now.Subtract(startTime).Ticks * (text.Length - (i + 1)) / (i + 1));
                            lblTimer.Text = "Time left : " + timeRemaining.ToString("mm\\:ss");
                        }));
                    }
                    if (text[i] != null)
                    {
                        // Insert the text into the SharedStringTablePart.
                        index = InsertSharedStringItem(text[i], shareStringPart);
                        // Insert cell A1 into the new worksheet.
                        cell = InsertCellInWorksheet(cellName[i].ToString(), fieldNumber, worksheetPart);
                        // Set the value of cell
                        cell.CellValue = new CellValue(index.ToString());
                        cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                        handleListBoxItems((int)i, text[i]);
                        addValueToProgressbBar();
                        
                    }

                }
                // Save the new worksheet.
                worksheetPart.Worksheet.Save();
            }
        }

        public static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string sheetName)
        {
            string relId = workbookPart.Workbook.Descendants<Sheet>().First(s => sheetName.Equals(s.Name)).Id;
            return (WorksheetPart)workbookPart.GetPartById(relId);
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private static WorksheetPart InsertWorksheet(WorksheetPart worksheetPart, WorkbookPart workbookPart)
        {

            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            // Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            Sheet sheets = workbookPart.Workbook.Descendants<Sheet>().
            Where(s => s.Name == "mySheet").FirstOrDefault();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            /* uint sheetId = 1;
             if (sheets.Elements<Sheet>().Count() > 0)
             {
                 sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
             }

             string sheetName = "Sheet" + sheetId;*/

            // Append the new worksheet and associate it with the workbook.
            //Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            // sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }
                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }

        #endregion

        #region form actions
        private void button1_Click(object sender, EventArgs e)
        {
            if (fileLocation.Equals(""))
            {
                MessageBox.Show("You have to select a excel file first.");
                return;
            }
            else
            {
                Invoke(new Action(() =>
                {
                    progressBar = progressBar1;
                    listbox = listBox1;
                }));
                //multithreading - > to avoid the UI from 'freezing'
                ThreadPool.QueueUserWorkItem(doWork);
            }
        }

        /**
         * This method will make a connection with the sql db and execute all the queries,
         * placing the data in an excel file
         * */
        private void doWork(object state)
        {
            try
            {
                server = "localhost";
                database = "e4p";
                uid = "root";
                password = "root";
                string connectionString;
                //encrypt your connection string with : https://msdn.microsoft.com/en-us/library/ff647398.aspx
                connectionString = "SERVER=" + server + ";" + "DATABASE=" +
                database + ";" + "UID=" + uid + ";" + "PASSWORD=" + password + ";";
                connection = new MySqlConnection(connectionString);
                connection.Open();
                executeQuery(connection, "extractFile");
                MessageBox.Show("Excel file has been filled up correctly.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /**
         * This method will print out the record number and the text of the record in a listbox
         * */
        private void handleListBoxItems(int itemId, string text)
        {
            Invoke(new Action(() =>
            {
                listbox.Items.Add(itemId + ": Imported " + text + "!");
                listbox.Refresh();
                int visibleItems = listbox.ClientSize.Height / listbox.ItemHeight;
                listbox.TopIndex = Math.Max(listbox.Items.Count - visibleItems + 1, 0);
            }));
        }

        private void addValueToProgressbBar()
        {
            Invoke(new Action(() =>
            {
                progressBar.Value = progressBar.Value + 1;
                progressBar.Refresh();
            }));
        }

        #endregion

        #region MySql methods
        /**
         * This method will call a stored procedure and add the value to an excel field
         * */
        private void executeQuery(MySqlConnection connection, String storeProcedure)
        {
            var command = new MySqlCommand(storeProcedure, connection);
            command.CommandType = CommandType.StoredProcedure;
            int listboxCounter = 0;
            uint unicode = 65;//start at A
            int amtFields = calculateLoadingBar(command);
            string[] data = new string[amtFields];
            char[] fields = new char[amtFields];
            using (var dataReader = command.ExecuteReader())
            {
                while (dataReader.Read())
                {
                    for (int i = 0; i < dataReader.FieldCount; i++)
                    {
                        //cache the data into arrays
                        data[listboxCounter] = dataReader.GetString(i);
                        fields[listboxCounter++] = (char)unicode++;
                    }
                    unicode = 65;
                }
                dataReader.Close();
            }
            addExcelFieldValue(data, fields);
        }

        private int calculateLoadingBar(MySqlCommand command)
        {
            int amtFieldsCounter = 0;
            using (var dataReader = command.ExecuteReader())
            {
                while (dataReader.Read())
                {
                    for (int i = 0; i < dataReader.FieldCount; i++)
                    {
                        amtFieldsCounter++;
                    }
                }
                dataReader.Close();
            }
            Invoke(new Action(() =>
            {
                progressBar.Maximum = ++amtFieldsCounter; //set the loading bar maximum to the amount of fields
            }));
            return amtFieldsCounter;
        }
        #endregion

        private void btnLoadExcel_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                fileLocation = openFileDialog1.FileName;
            }
        }
    }
}