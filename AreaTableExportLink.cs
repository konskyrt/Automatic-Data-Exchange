//
// Konstantinos Kyrtsonis
//

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;
using System.Windows.Forms;

// --- EPPLUS
using OfficeOpenXml;
using OfficeOpenXml.Style;

using Autodesk.AutoCAD.Runtime;
using AcadApp=Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace AreaTableExcelLink
{


    public class AreaTableExportLink
    {
        private AcadApp.Document mDoc;
        private Editor mEd;
        private Database mDb;

        // --- All tables with style "AreaTables" in the drawing
        private List<ObjectId> mTableList;
        
        // --- All tables using the Area Tables styles keyed by the entitiy handle
        private Dictionary<string, TableData> mTableDataDictionary;
        
        // --- Table values read in during import
        private Dictionary<string, TableData> mTableDataImportDictionary;

        // --- Sub Area Names Dictionary associated with each AreaItem
        private Dictionary<string, List<string>> mSubAreaNamesDictionary = new Dictionary<string, List<string>>();

        // --- A list of column numbers used for adding borders to columns once the data has been imported
        private List<ColumnRange> mMainColumnBoxRanges;

        public AreaTableExportLink()
        {
            // Document - Each open drawing will have an associated Document object. The Document object contains
            // information such as the filename, the MFC CDocument object, the current database and the save
            // format of the current drawing.
            mDoc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            
            // Editor – The Editor class is the interface with the user. In this example messages are written to
            // the command line and the selection functions are used to search for AutoCAD Table entities.
            mEd = mDoc.Editor;

            // Database - The Database class represents the AutoCAD drawing file. Each Database object contains
            // the various header variables, symbol tables, table records, entities and objects that make up
            // the drawing.
            mDb = mDoc.Database;
        }

        // ----------------------------------------------------------------------------------------------------------

        // Function to export the areas in the AutoCAD tables out to an Excel spreadsheet

        public void ExportToExcel()
        {
            // --- Find the AutoCAD Tables in the drawing
            FindTables();

            // --- Read the tables to create the table data dictionary
            ReadTableData();

            if (mTableDataDictionary.Count > 0)
            {
                // --- Prompt the user to select an Excel file name
                string filename = SelectExcelFileName(false  /* No need to check file exists */);
                
                // --- If a file was specified
                if (filename != null)
                {
                    // --- Go ahead and write the areas out to the Excel file
                    WriteExcel(filename);
                }
            }
            else
            {
                AcadApp.Application.ShowAlertDialog($"No tables using \"AreaTables\" style were found.");
            }
        }

        // ----------------------------------------------------------------------------------------------------------


        // Function to read the areas back into the drawing from Excel.The first two functions are identical to
        // the first two functions in the ExportToExcel function, and this we use FindTables to get a list of
        // the table objects and ReadTableData to read the tables to create the table data dictionary
        // (mTableDataDictionary). This is because when we import the data from the Excel file, we update
        // just the data that has changed.

        public void ImportFromExcel()
        {
            // --- Find the AutoCAD Tables in the drawing
            FindTables();

            // --- Read the tables to create the table data dictionary
            ReadTableData();

            if (mTableDataDictionary.Count > 0)
            {
                // --- Prompt the user to select an Excel file name
                string filename = SelectExcelFileName(true  /* file does need to exist */);

                // --- If a file was selected
                if (filename != null)
                {
                    // --- Read in the area from the Excel file and save to
                    // --- another Area Table dictionary
                    ReadExcel(filename);

                    if (mTableDataImportDictionary.Count > 0)
                    {
                        // --- Compare the imported data with the current table data
                        // --- and update table's whose areas have changed.
                        UpdateFromImportedValues();
                    }
                    else
                    {
                        AcadApp.Application.ShowAlertDialog($"No Area Tables were imported from the Excel.");
                    }
                }
            }
            else
            {
                AcadApp.Application.ShowAlertDialog($"There are no Area Tables in this drawing.");
            }
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- This function matches up the data in the drawing with the data read in from the spreadsheet
        // --- These are linked by the entity handle. The new data will be compared with the current values
        // --- and if they've changed the current values are updated if they are not read-only.

        private void UpdateFromImportedValues()
        {

            using (Transaction t = mDb.TransactionManager.StartTransaction())
            {
                try
                {
                    // --- Loop through all the tables read from the Excel sheet
                    foreach (KeyValuePair<string, TableData> kvp in mTableDataImportDictionary)
                    {
                        TableData importedTableData = kvp.Value as TableData;

                        if (importedTableData == null) throw new System.Exception("Null TableData");

                        // --- Does a table with the same handle exist in the drawing?
                        if (mTableDataDictionary.ContainsKey(importedTableData.mHandle))
                        {
                            TableData tableData = mTableDataDictionary[importedTableData.mHandle];

                            if (tableData == null) throw new System.Exception("Null TableData");

                            // --- Initially we are only reading the table
                            Table tbl = t.GetObject(tableData.mId, OpenMode.ForRead) as Table;

                            if (tbl == null) throw new System.Exception("Null Table");

                            int totalRows = tbl.Rows.Count;

                            AreaItem currentDWGAreaItem = null;
                            AreaSubItem currentDWGAreaSubItem = null;
                            AreaItem importedAreaItem = null;
                            AreaSubItem importedAreaSubItem = null;
                            double? dwgArea = null;
                            double? importedArea = null;

                            string number = tableData.mNumber;
                            string parzelle = tableData.mParzelle;
                            string address = tableData.mAddress;

                            // --- Remove blank lines and replace with a single space & commma so that address is on one line
                            address = StripAddressLineBreasks(address);

                            // --- If the Area Table number is different update the table
                            if (!number.Equals(importedTableData.mNumber))
                            {
                                tbl.UpgradeOpen();  // --- Upgrade to enable write access to table
                                mEd.WriteMessage($"\nUpdating Number for table {tableData.mNumber}");
                                tbl.Cells[0, 0].TextString = importedTableData.mNumber;
                            }

                            // --- If the Area Table parzell is different update the table
                            if (!parzelle.Equals(importedTableData.mParzelle))
                            {
                                tbl.UpgradeOpen();
                                mEd.WriteMessage($"\nUpdating Parzelle for table {tableData.mNumber}");
                                tbl.Cells[0, 1].TextString = importedTableData.mParzelle;
                            }

                            // --- If the Area Table address is different update the table
                            if (!address.Equals(importedTableData.mAddress))
                            {
                                // --- Reverse the comma substitution and replace commas with line breaks
                                address = importedTableData.mAddress.Replace(", ", "\r\n");
                                tbl.UpgradeOpen();
                                mEd.WriteMessage($"\nUpdating Address for table {tableData.mNumber}");
                                tbl.Cells[1, 1].TextString = address;
                            }

                            for (int row = 3; row < totalRows; row++)
                            {
                                string itemName = tbl.Cells[row, 1].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                                string area = removeAreaSuffix(tbl.Cells[row, 2].GetTextString(FormatOption.IgnoreMtextFormat));
                                double? areaDouble = null;

                                try
                                {
                                    areaDouble = ConvertToDouble(area);

                                    // --- Main Area
                                    if (!itemName.StartsWith("-"))
                                    {
                                        currentDWGAreaItem = tableData.GetAreaItem(itemName);
                                        importedAreaItem = importedTableData.GetAreaItem(itemName);
                                        if ((currentDWGAreaItem != null) && (importedTableData != null))
                                        {
                                            dwgArea = currentDWGAreaItem.mArea;
                                            importedArea = importedAreaItem.mArea;
                                            if (!AreasEqual(dwgArea, importedArea))
                                            {
                                                if (!currentDWGAreaItem.mReadonly)
                                                {
                                                    tbl.UpgradeOpen();
                                                    mEd.WriteMessage($"\nUpdating {itemName} for table {tableData.mNumber}");
                                                    // --- Convert the double to a string and add the "m2" suffix with mtext formatting
                                                    tbl.Cells[row, 2].TextString = @"{" + importedArea.ToString() + @" m\H0.7x;\S2^;}";
                                                }
                                                else
                                                {
                                                    mEd.WriteMessage($"\nNot overriding Field in {itemName} for table {tableData.mNumber}");
                                                }
                                            }
                                        }
                                    }
                                    // --- A Sub Area can't be before a Main Area
                                    else if ((currentDWGAreaItem != null) && itemName.StartsWith("-"))
                                    {
                                        // --- Ignore the "-" prefix
                                        itemName = itemName.Substring(1).Trim();
                                        currentDWGAreaSubItem = tableData.GetSubAreaItem(currentDWGAreaItem.mName, itemName);
                                        importedAreaSubItem = importedTableData.GetSubAreaItem(currentDWGAreaItem.mName, itemName);

                                        if ((currentDWGAreaSubItem != null) && (importedAreaSubItem != null))
                                        {
                                            dwgArea = currentDWGAreaSubItem.mArea;
                                            importedArea = importedAreaSubItem.mArea;
                                            if (!AreasEqual(dwgArea, importedArea))
                                            {
                                                if (!currentDWGAreaSubItem.mReadonly)
                                                {
                                                    tbl.UpgradeOpen();
                                                    mEd.WriteMessage($"\nUpdating {currentDWGAreaItem.mName} -{itemName} for table {tableData.mNumber}");
                                                    // --- Convert the double to a string and add the "m2" suffix 
                                                    tbl.Cells[row, 2].TextString = @"{" + importedArea.ToString() + @" m\H0.7x;\S2^;}";
                                                }
                                                else
                                                {
                                                    mEd.WriteMessage($"\nNot overriding Field in {currentDWGAreaItem.mName} -{itemName} for table {tableData.mNumber}");
                                                }
                                            }
                                        }
                                    }

                                }
                                catch (System.Exception ex)
                                {
                                    mEd.WriteMessage($"Exception {ex.Message} in table {tableData.mHandle}");
                                }
                            }
                        }
                        else
                        {
                            mEd.WriteMessage($"No table found with handle: {importedTableData.mHandle}");
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    AcadApp.Application.ShowAlertDialog($"Exception in UpdateFromImportedValues: {ex.Message}");
                }

                // --- Commit any write changes made the to the AutoCAD table entities
                t.Commit();
            }
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- Test whether two areas are equal

        private bool AreasEqual(double? a1, double? a2)
        {
            // --- Both Null so they're "equal"
            if (!a1.HasValue && !a2.HasValue) return true;

            // --- One is null the other isn't so they're "not equal"
            if (!a1.HasValue && a2.HasValue) return false;
            if (a1.HasValue && !a2.HasValue) return false;

            // --- They both have double values and they're equal
            if (a1.HasValue && a2.HasValue && a1.Value.Equals(a2.Value)) return true;

            return false;
        }

        // ----------------------------------------------------------------------------------------------------------

        // Function uses SelectionFilter to get a list of all the AutoCAD Table entities inserted in Model Space.
        // Then loops through each of the tables and create a new list of the Table entities
        // that use the “AreaTables” TableStyle.

        private void FindTables()
        {
            mTableList = new List<ObjectId>();

            // --- Filter for table objects in model space
            // --- These are based on DXF codes
            TypedValue[] selFilter = new TypedValue[]
            {
                new TypedValue(0, "ACAD_TABLE"),   // https://help.autodesk.com/view/OARX/2020/ENU/?guid=GUID-D8CCD2F0-18A3-42BB-A64D-539114A07DA0
                new TypedValue(410, "Model")       // https://help.autodesk.com/view/OARX/2020/ENU/?guid=GUID-3610039E-27D1-4E23-B6D3-7E60B22BB5BD
            };

            ObjectIdCollection selectedTables = null;

            SelectionFilter oSf = new SelectionFilter(selFilter);

            // --- Use SelectAll with a filter to find all the Table entities in Model Space
            PromptSelectionResult selection = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.SelectAll(oSf);
            if (selection.Status == PromptStatus.OK)
            {
                ObjectId[] ids = selection.Value.GetObjectIds();
                selectedTables = new ObjectIdCollection(ids);

                // --- We need to start a transaction to read the entities
                using (Transaction t = mDb.TransactionManager.StartTransaction())
                {
                    try
                    {
                        // --- Loop through all selected entities
                        foreach (ObjectId objId in selectedTables)
                        {
                            // --- They should all be tables but "as Table" will return null if
                            // --- the DBObject can't be cast to a Table
                            Table table = t.GetObject(objId, OpenMode.ForRead) as Table;

                            // --- If it is a table
                            if (table != null)
                            {
                                // --- Get the ObjectId for the Table's style
                                ObjectId tableStyleId = table.TableStyle;
                                // --- Now open that to read
                                TableStyle tableStyle = t.GetObject(tableStyleId, OpenMode.ForRead) as TableStyle;
                                // --- and get it's name to check it equals AreaTables
                                // --- Suggestion, this could probably be done with the SelectionFilter by including DXF code 342
                                // --- which is the Hard pointer ID of the TABLESTYLE object
                                if (tableStyle.Name.Equals("AreaTables"))
                                {
                                    mTableList.Add(objId);
                                }
                            }
                        }

                        // AcadApp.Application.ShowAlertDialog($"Found {mTableList.Count} tables!");

                    }
                    catch (System.Exception ex)
                    {
                        AcadApp.Application.ShowAlertDialog($"Exception in FindTables: {ex.Message}");
                    }
                }
            }
        }

        // ----------------------------------------------------------------------------------------------------------

        // This function reads the data in the tables and sorts into a dictionary (mTableDataDictionary)
        // keyed with the table’s entity handle.

        private void ReadTableData()
        {
            // --- Dictionary of table data keyed by the table entity handle
            mTableDataDictionary = new Dictionary<string, TableData>();

            using (Transaction t = mDb.TransactionManager.StartTransaction())
            {
                try
                {
                    foreach (ObjectId objId in mTableList)
                    {
                        Table tbl = t.GetObject(objId, OpenMode.ForRead) as Table;

                        // --- Throwing an exception is a handy way of breaking out the 
                        // --- code and handling the situation with a catch block
                        if (tbl == null) throw new System.Exception("Null Table");

                        TableData td = new TableData(objId);

                        // --- The table Number, Parzelle and Address are located in fixed cell locations of the table.
                        td.mHandle = tbl.Handle.ToString().ToUpper();
                        td.mNumber = tbl.Cells[0, 0].GetTextString(FormatOption.IgnoreMtextFormat);
                        td.mParzelle = tbl.Cells[0, 1].GetTextString(FormatOption.IgnoreMtextFormat);

                        string address = tbl.Cells[1, 1].GetTextString(FormatOption.IgnoreMtextFormat);

                        // --- Remove blank lines and replace with a single space & commma so that address is on one line
                        address = StripAddressLineBreasks(address);

                        td.mAddress = address;

                        int totalRows = tbl.Rows.Count;

                        AreaItem currentAreaItem = null;

                        // --- Loop through each of the data rows in the spreadsheet
                        for (int row = 3; row < totalRows; row++)
                        {
                            string itemName = tbl.Cells[row, 1].GetTextString(FormatOption.IgnoreMtextFormat).Trim();
                            Cell areaCell = tbl.Cells[row, 2];

                            // --- Mtext adds a loads of formatting to the text. We want the text without the formatting.
                            string areaStr = areaCell.GetTextString(FormatOption.FormatOptionNone);

                            // --- Remove the "m2" if it is there
                            string area = removeAreaSuffix(areaCell.GetTextString(FormatOption.IgnoreMtextFormat));

                            // --- The ? means that the double is nullable and if it isn't null
                            // --- then we know it has been assigned a value
                            double? areaDouble = null;

                            try
                            {
                                if ((area != null) && (area != "")) areaDouble = Convert.ToDouble(area);

                                if ((currentAreaItem != null) && itemName.StartsWith("-"))
                                {
                                    // --- Ignore the "-" prefix
                                    itemName = itemName.Substring(1).Trim();
                                    
                                    AreaSubItem areaSubItem = new AreaSubItem(itemName, areaDouble);

                                    // --- If the Cell in linked to a Field, mark it as readonly so it can't be overwritten by import
                                    if (areaCell.FieldId != ObjectId.Null)
                                        areaSubItem.mReadonly = true;

                                    currentAreaItem.mSubItems.Add(areaSubItem);
                                }
                                else
                                {
                                    if (currentAreaItem != null) td.mAreaItems.Add(currentAreaItem);

                                    currentAreaItem = new AreaItem(itemName, areaDouble);

                                    // --- If the Cell in linked to a Field, mark it as readonly so it can't be overwritten by import
                                    if (areaCell.FieldId != ObjectId.Null)
                                        currentAreaItem.mReadonly = true;
                                }
                            }
                            catch (System.Exception ex)
                            {
                                mEd.WriteMessage($"Exception {ex.Message} in table {td.mHandle}");
                            }
                        }

                        if (currentAreaItem != null) td.mAreaItems.Add(currentAreaItem);

                        // --- Add a new TableData to the dictionary
                        mTableDataDictionary.Add(td.mHandle, td);
                    }
                }
                catch (System.Exception ex)
                {
                    AcadApp.Application.ShowAlertDialog($"Exception in ReadTableData: {ex.Message}");
                }
            }
        }

        // ----------------------------------------------------------------------------------------------------------
        
        // --- Replace break lines in an address string with commas

        private string StripAddressLineBreasks(string address)
        {
            while (address.Contains("\r\n\n\r\n\n")) address = address.Replace("\r\n\n\r\n\n", "\r\n\n");
            while (address.Contains("\r\n\n")) address = address.Replace("\r\n\n", ", ");
            return address;
        }
        
        // ----------------------------------------------------------------------------------------------------------

        // --- Remove the "m2" from the end of the area strings

        private string removeAreaSuffix(string txt)
        {
            // --- Set to lower case
            txt = txt.ToLower();

            if (txt.Contains($"m2/"))
            {
                // --- Replace "m2" with nothing
                txt = txt.Replace($"m2/", "");
            }

            // --- Remove and leading or trailing spaces
            return txt.Trim();
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- Function to specify a file name for export or select a file for import

        private string SelectExcelFileName(bool checkFileExists)
        {
            // --- The Properties.Settings class is used to saving and restoring default values
            string folder = Properties.Settings.Default.Folder;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (Directory.Exists(folder)) openFileDialog.InitialDirectory = folder;
            openFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx";
            openFileDialog.Title = "Select Excel file";
            openFileDialog.CheckFileExists = checkFileExists;

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // --- Save selected folder as default for next time
                Properties.Settings.Default.Folder = Path.GetDirectoryName(openFileDialog.FileName);
                Properties.Settings.Default.Save();
                return openFileDialog.FileName;
            }

            return null;
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- Test whether the file can be written to
        // --- This will return false if the file cannot be written to which is typically because
        // --- the XLS file is open in Excel.

        private bool isFileOkToOpen(string filename)
        {
            try
            {
                Stream s = File.Open(filename, FileMode.Open, FileAccess.Read, FileShare.None);
                s.Close();
                return true;
            }
            catch (System.Exception)
            {
                return false;
            }
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- Write to Excel using Epplus
        private void WriteExcel(string filename)
        {
            // --- If we get an exception within this Try block it will be caught by the Catch block
            // --- and the exception message will be displayed in an alert but the code won't crash
            try
            {
                // --- If the target file exists, check it can be opened
                if (File.Exists(filename) && !isFileOkToOpen(filename))
                {
                    AcadApp.Application.ShowAlertDialog($"Can't open {filename}.\nCheck it isn't already open in Excel.");
                    return;
                }

                // --- Delete file if it already exists
                if (File.Exists(filename))
                    File.Delete(filename);

                FileInfo fileInfo = new FileInfo(filename);
                ExcelPackage package = new ExcelPackage(fileInfo);

                ExcelWorkbook workbook = package.Workbook;

                // --- Create a new Worksheet called "Area Tables"
                ExcelWorksheet ws = workbook.Worksheets.Add("Area Tables");

                ExcelRange range;

                // --- Define a dictionary for storing unique Sub Area names
                mSubAreaNamesDictionary = new Dictionary<string, List<string>>();


                foreach (KeyValuePair<string, TableData> kvp in mTableDataDictionary)
                {
                    TableData td = kvp.Value as TableData;

                    foreach (AreaItem ai in td.mAreaItems)
                    {
                        List<string> subNames = new List<string>();

                        if (!mSubAreaNamesDictionary.ContainsKey(ai.mName))
                            mSubAreaNamesDictionary.Add(ai.mName, subNames);

                        foreach (AreaSubItem areaSubItem in ai.mSubItems)
                        {
                            if (!mSubAreaNamesDictionary[ai.mName].Contains(areaSubItem.mName))
                                mSubAreaNamesDictionary[ai.mName].Add(areaSubItem.mName);
                        }
                    }
                }


                // --- We create a list of single columns and pairs of columns so
                // --- that we know where to add the borders
                mMainColumnBoxRanges = new List<ColumnRange>();


                int r = 1;
                int c = 1;

                // --- Write the first two rows with the column header names

                // Parz.	Enteig	Er-	Dienstbarkeit	Address	
                // --- "Handle" column - this is used to link the Excel row with the AutoCAD Table entity
                // --- This an other header cells will have a light green solid fill
                range = ws.Cells[r, c];
                range.Value = "Handle";
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                // --- Nothing in the cell below "Handle"
                range = ws.Cells[r + 1, c];
                range.Value = "";
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                // --- Add a border around the two cells
                range = ws.Cells[1, c, 2, c];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                mMainColumnBoxRanges.Add(new ColumnRange(c));

                c++;

                // --- Add the "Parz" column
                range = ws.Cells[r, c];
                range.Value = "Parz.";
                range.Style.WrapText = false;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                // --- Below that add the "Nr" cell
                range = ws.Cells[r + 1, c];
                range.Value = "Nr";
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                range = ws.Cells[1, c, 2, c];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                mMainColumnBoxRanges.Add(new ColumnRange(c));

                c++;

                range = ws.Cells[r, c];
                range.Value = "Enteig";
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                range = ws.Cells[r + 1, c];
                range.Value = "Nr";
                range.Style.WrapText = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                range = ws.Cells[1, c, 2, c];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                mMainColumnBoxRanges.Add(new ColumnRange(c));

                // --- Create a dictionary where the key is the header name and the data
                // --- is the colummn number

                Dictionary<string, int> columnDictionary = new Dictionary<string, int>();

                foreach (KeyValuePair<string, List<string>> kvp in mSubAreaNamesDictionary)
                {
                    string areaItemName = kvp.Key;
                    c++;

                    range = ws.Cells[r, c];
                    range.Value = areaItemName;
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                    if (!columnDictionary.ContainsKey(areaItemName))
                        columnDictionary.Add(areaItemName, c);

                    range = ws.Cells[r + 1, c];
                    range.Value = "m²";
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                    List<string> subAreas = mSubAreaNamesDictionary[areaItemName];

                    int areaStartCol = c;

                    // --- If the Area Item is not "Temp. Nutzung" do this
                    if (!areaItemName.Equals("Temp. Nutzung"))
                    {
                        if (subAreas.Count > 0)
                        {
                            c++;
                            range = ws.Cells[r + 1, c];
                            range.Value = "Art";
                            range.Style.WrapText = true;
                            range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                            range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);

                            if (!columnDictionary.ContainsKey($"{areaItemName}!Art"))
                                columnDictionary.Add($"{areaItemName}!Art", c);

                            // --- Merge Top Row
                            ws.Cells[r, areaStartCol, r, c].Merge = true;
                        }

                        range = ws.Cells[1, areaStartCol, 2, c];
                        range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                        mMainColumnBoxRanges.Add(new ColumnRange(areaStartCol, c));

                    }
                    // --- Else if Area Item is "Temp. Nutzung" add additional "parameter n"
                    // --- columns for each unquie property
                    // --- TODO - Not sure why "Temp. Nutzung" is a special case
                    else
                    {
                        // --- One or more subareas
                        if (subAreas.Count > 0)
                        {
                            int pc = 1;

                            foreach (string subAreaName in subAreas)
                            {
                                if (pc > 1)
                                {
                                    c++;
                                    range = ws.Cells[r + 1, c];
                                    range.Value = "m²";
                                    range.Style.WrapText = true;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                }

                                c++;

                                range = ws.Cells[r + 1, c];
                                range.Value = $"parameter {pc}";  // --- e.g. "parameter 2"
                                range.Style.WrapText = true;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGreen);
                                if (!columnDictionary.ContainsKey($"{areaItemName}!parameter {pc}"))
                                    columnDictionary.Add($"{areaItemName}!parameter {pc}", c);
                                pc++;
                            }

                            // --- Merge Top Row
                            ws.Cells[r, areaStartCol, r, c].Merge = true;

                            mMainColumnBoxRanges.Add(new ColumnRange(areaStartCol, c));

                            range = ws.Cells[1, areaStartCol, 2, c];
                            range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                        }
                    }
                }

                c++;

                range = ws.Cells[r, c];
                range.Value = "Address";
                range.Style.WrapText = false;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                range = ws.Cells[r + 1, c];
                range.Value = "";
                range.Style.WrapText = false;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);

                int lastColumn = c;

                // --- Set default column widths
                for (int i = 1; i < c; i++)
                    ws.Column(i).Width = 20;

                // --- Address Column
                ws.Column(lastColumn).Width = 65;


                // --- ExcelRange this[int FromRow, int FromCol, int ToRow, int ToCol]
                // --- Top two rows
                range = ws.Cells[1, 1, 2, c];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                r++; // --- Skip the 2nd header row

                // --- Now write out each table as a single row to the spreadsheet

                foreach (KeyValuePair<string, TableData> kvp in mTableDataDictionary)
                {
                    r++;

                    TableData td = kvp.Value as TableData;

                    if (td == null) throw new System.Exception("Null TableData");

                    range = ws.Cells[r, 1];
                    range.Style.Numberformat.Format = "@";  // --- Read number text as text and not a number
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Value = td.mHandle;

                    range = ws.Cells[r, 2];
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Value = td.mParzelle;

                    range = ws.Cells[r, 3];
                    range.Style.Numberformat.Format = "@";  // --- Read number text as text and not a number
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                    range.Value = td.mNumber;

                    Debug.WriteLine("Stop");

                    foreach (AreaItem ai in td.mAreaItems)
                    {
                        string areaItemName = ai.mName;
                        List<AreaSubItem> subItems = ai.mSubItems;

                        // --- No Sub-Items so area displayed in first column of area item
                        if (subItems.Count == 0)
                        {
                            if (columnDictionary.ContainsKey(areaItemName))
                            {
                                int col = columnDictionary[areaItemName];
                                range = ws.Cells[r, col];
                                range.Value = ai.mArea.ToString();
                                range.Style.WrapText = true;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                        }
                        else if (!areaItemName.Equals("Temp. Nutzung") && (subItems.Count == 1))
                        {
                            AreaSubItem asi = subItems.First();
                            string colKey = $"{areaItemName}!Art";
                            if (columnDictionary.ContainsKey(colKey))
                            {
                                int col = columnDictionary[colKey];
                                range = ws.Cells[r, col - 1];
                                range.Value = asi.mArea.ToString();
                                range.Style.WrapText = true;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                range = ws.Cells[r, col];
                                range.Value = asi.mName;
                                range.Style.WrapText = true;
                                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                            }
                        }
                        else if (areaItemName.Equals("Temp. Nutzung") && (subItems.Count > 1))
                        {
                            int pc = 1;
                            foreach (AreaSubItem asi in subItems)
                            {
                                string colKey = $"{areaItemName}!parameter {pc}";
                                if (columnDictionary.ContainsKey(colKey))
                                {
                                    int col = columnDictionary[colKey];
                                    range = ws.Cells[r, col - 1];
                                    range.Value = asi.mArea.ToString();
                                    range.Style.WrapText = true;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                                    range = ws.Cells[r, col];
                                    range.Value = asi.mName;
                                    range.Style.WrapText = true;
                                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                                }

                                pc++;
                            }

                        }
                    }

                    range = ws.Cells[r, lastColumn];
                    range.Value = td.mAddress;
                    range.Style.WrapText = true;
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;

                    // --- Add a border around the row
                    range = ws.Cells[r, 1, r, lastColumn];
                    range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                }

                // --- Add borders for columns 

                foreach (ColumnRange cr in mMainColumnBoxRanges)
                {
                    int numCols = cr.mC2 - cr.mC1;
                    // --- Thin box around area / parameter pair columns
                    if (numCols > 2)
                    {
                        for (int i = cr.mC1; i < cr.mC2; i += 2)
                        {
                            range = ws.Cells[2, i, r, i+1];
                            range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        }
                    }
                    // --- Thin box around sub header of a pair
                    else if (numCols == 1)
                    {
                        range = ws.Cells[2, cr.mC1, 2, cr.mC2];
                        range.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }
                    range = ws.Cells[1, cr.mC1, r, cr.mC2];
                    range.Style.Border.BorderAround(ExcelBorderStyle.Medium);
                }

                // --- Add Header Borders

                range = ws.Cells[1, 1, 2, lastColumn];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

                // --- Add outer borders

                range = ws.Cells[1, 1, r, lastColumn];
                range.Style.Border.BorderAround(ExcelBorderStyle.Medium);

#if !DEBUG
                // --- Hide handle column
                ws.Column(1).Hidden = true;
#endif
                // --- Save the Excel file
                package.Save();

                // --- Double check the file exists
                if (File.Exists(filename))
                {
                    // --- Display the file name in an alert
                    AcadApp.Application.ShowAlertDialog($"Area Tables exported to {filename}.");
                }
            }
            catch (System.Exception ex)
            {
                AcadApp.Application.ShowAlertDialog($"Exception in WriteExcel: {ex.Message}");
            }
        }

        // ----------------------------------------------------------------------------------------------------------

        // --- Read table data from an Excel file and store it in a TableData dictionary

        private void ReadExcel(string filename)
        {
            mTableDataImportDictionary = new Dictionary<string, TableData>();

            try
            {
                // --- If the target file exists, check it can be opened
                if (!File.Exists(filename))
                {
                    AcadApp.Application.ShowAlertDialog($"Can't find {filename}.");
                    return;
                }

                FileInfo fileInfo = new FileInfo(filename);
                ExcelPackage package = new ExcelPackage(fileInfo);

                ExcelWorkbook workbook = package.Workbook;

                ExcelWorksheet ws = null;

                ExcelColumns excelColumns = new ExcelColumns();

                // --- Look for the "Area Tables" work sheet
                foreach (ExcelWorksheet worksheet in workbook.Worksheets)
                {
                    if (worksheet.Name.Equals("Area Tables"))
                    {
                        ws = worksheet;
                        break;
                    }
                }

                if (ws == null)
                {
                    package.Dispose();
                    AcadApp.Application.ShowAlertDialog($"Can't find WorkSheet \"Area Tables\".");
                    return;
                }

                int rowCount = ws.Dimension.End.Row;
                int colCount = ws.Dimension.End.Column;

                object cellObj = null;

                // Handle Parz.	Enteig	Landerwerb	Dienstbarkeit		Temp. Nutzung				Address

                // --- Read the first row of the spreadsheet to determine the column numbers for the properties
                for (int c = 1; c <= colCount; c++)
                {
                    object header1Obj = ws.Cells[1, c].Value;
                    object header2Obj = ws.Cells[2, c].Value;
                    if (header1Obj != null)
                    {
                        string header = header1Obj.ToString().Trim();
                        if (header.Equals("Handle", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mHandle = c;
                        }
                        else if (header.Equals("Parz.", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mParz = c;
                        }
                        else if (header.Equals("Enteig", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mEnteig = c;
                        }
                        else if (header.Equals("Landerwerb", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mLanderwerb = c;
                            // --- Look down and ahead for the Art column
                            cellObj = ws.Cells[2, c + 1].Value;
                            if ((cellObj != null) && (cellObj.ToString().Trim().Equals("Art")))
                            {
                                excelColumns.mLanderwerbSubItems.Add(c + 1);
                                c++;
                            }
                        }
                        else if (header.Equals("Dienstbarkeit", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mDienstbarkeit = c;
                            // --- Look down and ahead for the Art column
                            cellObj = ws.Cells[2, c + 1].Value;
                            if ((cellObj != null) && (cellObj.ToString().Trim().Equals("Art")))
                            {
                                excelColumns.mDienstbarkeitSubItems.Add(c + 1);
                                c++;
                            }
                        }
                        else if (header.Equals("Temp. Nutzung", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mTempNutzung = c;
                            int x = 1;
                            // --- Look down and ahead for the "parameter <n>" columns
                            cellObj = ws.Cells[2, c + x].Value;
                            while ((cellObj != null) && (cellObj.ToString().ToLower().StartsWith("parameter")))
                            {
                                excelColumns.mTempNutzungSubItems.Add(c + x);
                                x += 2;
                                cellObj = ws.Cells[2, c + x].Value;
                            }
                        }
                        else if (header.Equals("Address", StringComparison.CurrentCultureIgnoreCase))
                        {
                            excelColumns.mAddress = c;
                        }
                    }
                }

                if (!excelColumns.AllColumnsSet())
                {
                    throw new System.Exception("Missing column(s) in Excel");
                }

                int col;

                string cellText = "";

                // --- Loop through all the data rows
                for (int r = 3; r <= rowCount; r++)
                {
                    string handle = SafeValue(ws.Cells[r, excelColumns.mHandle]).ToUpper();
                    string landerwerbObj = excelColumns.mLanderwerb != -1 ? SafeValue(ws.Cells[r, excelColumns.mLanderwerb]) : null;
                    string dienstbarkeitObj = excelColumns.mDienstbarkeit != -1 ? SafeValue(ws.Cells[r, excelColumns.mDienstbarkeit]) : null;
                    string tempNutzungObj = excelColumns.mTempNutzung != -1 ? SafeValue(ws.Cells[r, excelColumns.mTempNutzung]) : null;

                    if (!handle.Equals(""))
                    {
                        try
                        {
                            if (!mTableDataImportDictionary.ContainsKey(handle))
                            {
                                try
                                {
                                    TableData td = new TableData(handle);
                                    td.mParzelle = SafeValue(ws.Cells[r, excelColumns.mParz]);
                                    td.mNumber = SafeValue(ws.Cells[r, excelColumns.mEnteig]);

                                    // "Landerwerb"

                                    if (landerwerbObj != null)
                                    {
                                        // --- No Sub Paramaters, so area is in same column
                                        if (excelColumns.mLanderwerbSubItems.Count == 0)
                                        {
                                            td.mAreaItems.Add(new AreaItem("Landerwerb", ConvertToDouble(landerwerbObj)));
                                        }
                                        // --- Can be a single Sub parameter
                                        else if (excelColumns.mLanderwerbSubItems.Count == 1)
                                        {
                                            col = excelColumns.mLanderwerbSubItems.First();
                                            string subParamName = SafeValue(ws.Cells[r, col]);

                                            // --- Blank so it isn't a sub-parameter
                                            if (subParamName.Equals(""))
                                            {
                                                td.mAreaItems.Add(new AreaItem("Landerwerb", ConvertToDouble(landerwerbObj)));
                                            }
                                            else
                                            {
                                                AreaItem ai = new AreaItem("Landerwerb", null);
                                                AreaSubItem asi = new AreaSubItem(subParamName, ConvertToDouble(landerwerbObj));
                                                ai.mSubItems.Add(asi);
                                                td.mAreaItems.Add(ai);
                                            }
                                        }
                                    }

                                    // "Dienstbarkeit"
                                    if (dienstbarkeitObj != null)
                                    {
                                        // --- No Sub Paramaters, so area is in same column
                                        if (excelColumns.mDienstbarkeitSubItems.Count == 0)
                                        {
                                            td.mAreaItems.Add(new AreaItem("Dienstbarkeit", ConvertToDouble(dienstbarkeitObj)));
                                        }
                                        // --- Can be a single Sub parameter
                                        else if (excelColumns.mDienstbarkeitSubItems.Count == 1)
                                        {
                                            col = excelColumns.mDienstbarkeitSubItems.First();
                                            string subParamName = SafeValue(ws.Cells[r, col]);

                                            // --- Blank so it isn't a sub-parameter
                                            if (subParamName.Equals(""))
                                            {
                                                td.mAreaItems.Add(new AreaItem("Dienstbarkeit", ConvertToDouble(dienstbarkeitObj)));
                                            }
                                            else
                                            {
                                                AreaItem ai = new AreaItem("Dienstbarkeit", null);
                                                AreaSubItem asi = new AreaSubItem(subParamName, ConvertToDouble(dienstbarkeitObj));
                                                ai.mSubItems.Add(asi);
                                                td.mAreaItems.Add(ai);
                                            }
                                        }
                                    }

                                    // Temp. Nutzung

                                    if (tempNutzungObj != null)
                                    {
                                        // --- No Sub Paramaters, so area is in same column
                                        if (excelColumns.mTempNutzungSubItems.Count == 0)
                                        {
                                            td.mAreaItems.Add(new AreaItem("Temp. Nutzung", ConvertToDouble(tempNutzungObj)));
                                        }
                                        // --- Can be more than one Sub parameters
                                        else if (excelColumns.mTempNutzungSubItems.Count > 0)
                                        {
                                            foreach (int spcol in excelColumns.mTempNutzungSubItems)
                                            {
                                                string subParamName = SafeValue(ws.Cells[r, spcol]);

                                                // --- Blank so it isn't a sub-parameter
                                                if (subParamName.Equals(""))
                                                {
                                                    AreaItem ai = td.AddAreaItem("Temp. Nutzung");
                                                    ai.mArea = ConvertToDouble(tempNutzungObj);
                                                }
                                                else
                                                {
                                                    // --- Get area for preceeding cell
                                                    cellText = SafeValue(ws.Cells[r, spcol - 1]);
                                                    AreaItem ai = td.AddAreaItem("Temp. Nutzung");
                                                    AreaSubItem asi = new AreaSubItem(subParamName, Convert.ToDouble(cellText));
                                                    ai.mSubItems.Add(asi);
                                                }
                                            }
                                        }
                                    }

                                    td.mAddress = SafeValue(ws.Cells[r, excelColumns.mAddress]);

                                    mTableDataImportDictionary.Add(handle, td);
                                }
                                catch (SystemException ex)
                                {
                                    mEd.WriteMessage($"Ivalid data in row {r}");
                                }
                            }
                            else
                            {
                                mEd.WriteMessage($"Ignoring duplicate handle {handle} in row {r}");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            mEd.WriteMessage($"Invalid area in row {r}");
                        }
                    }
                }

                package.Dispose();
            }
            catch (System.Exception ex)
            {
                AcadApp.Application.ShowAlertDialog($"Exception in ReadExcel: {ex.Message}");
            }
        }

        // -------------------------------------------------------------------------------------------------------

        // --- Return the text from a cell or and empty string if there is none
        // --- It is "safe" in that it will always return a string

        private string SafeValue (ExcelRange range)
        {
            object obj = range.Value;
            return obj == null ? "" : obj.ToString().Trim();
        }
        
        // -------------------------------------------------------------------------------------------------------

        // --- This function will "catch" invalid text that can't be converted to a double value
        // --- and will return null if anything went wrong

        private double? ConvertToDouble (string txt)
        {
            double? ret = null;

            if (txt == null) return null;

            try
            {
                ret = Convert.ToDouble(txt);
            }
            catch
            {

            }

            return ret;
        }
    }

    // ------------------------------------------------------------------------------------------------------

    // --- This class stores the column numbers in the spreadsheet for each or the properties

    public class ExcelColumns
    {
        // Handle Parz.	Enteig	Landerwerb	Dienstbarkeit		Temp. Nutzung				Address

        public int mHandle = -1;
        public int mParz = -1;
        public int mEnteig = -1;
        public int mLanderwerb = -1;
        public List<int> mLanderwerbSubItems = new List<int>();
        public int mDienstbarkeit = -1;
        public List<int> mDienstbarkeitSubItems = new List<int>();
        public int mTempNutzung = -1;
        public List<int> mTempNutzungSubItems = new List<int>();
        public int mAddress = -1;
        public ExcelColumns() { }

        // ------------------------------------------------------------------------------------------------------

        public bool AllColumnsSet()
        {
            if (mHandle == -1) return false;
            if (mParz == -1) return false;
            if (mEnteig == -1) return false;
            if (mAddress == -1) return false;

            return true;
        }
    }
    
    // ------------------------------------------------------------------------------------------------------

    // --- The ColumnRange class simply holds a single column number, or a pair of column numbers
    public class ColumnRange
    {
        public int mC1;
        public int mC2;
        public ColumnRange(int c1, int c2)
        {
            mC1 = c1;
            mC2 = c2;
        }
        public ColumnRange(int c)
        {
            mC1 = c;
            mC2 = c;
        }
    }
}
