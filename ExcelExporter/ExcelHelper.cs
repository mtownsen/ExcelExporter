//ExcelHelper.cs
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
namespace CreateExcelSpreadsheet
{
    public class ExcelHelper
    {
        /// <summary>
        /// Write excel file of a list of object as T
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="fileName">Full path of the file name of excel spreadsheet</param>
        /// <param name="objects">List of data</param>
        /// <param name="sheetName">Sheet name of Excel File</param>
        /// <param name="headerNames">Optional list of header names for the spreadsheet</param>
        public void Create<T>(
            string fileName,
            List<T> objects,
            string sheetName,
            List<string> headerNames)
        {
            //Open the copied template workbook. 
            using (SpreadsheetDocument myWorkbook =
                   SpreadsheetDocument.Create(fileName,
                   SpreadsheetDocumentType.Workbook))
            {
                //First lets get the base workbook all setup.  We do this by adding a new workbook to our document as well as its associate worksheet. 
                WorkbookPart workbookPart = myWorkbook.AddWorkbookPart();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                // Create Styles and Insert into Workbook
                var stylesPart =
                    myWorkbook.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                Stylesheet styles = new CustomStylesheet();
                styles.Save(stylesPart);
                string relId = workbookPart.GetIdOfPart(worksheetPart);
                var workbook = new Workbook();
                var fileVersion =
                    new FileVersion
                    {
                        ApplicationName =
                            "Microsoft Office Excel"
                    };

                //Here we check if the caller passed in custom headernames.  If the caller did not we need to set the variable up based on the class provided by the caller. 
                if (headerNames == null || headerNames.Count() == 0)
                {
                    headerNames = new List<string>();
                    //Here we scan that class type the caller passed in to get all public instance variables at the highest level.
                    PropertyInfo[] properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance | BindingFlags.FlattenHierarchy);
                    //Once we have the properties associated with the highest level class in the hierarchy we loop through them and add each to the headernames variable.  
                    //What this allows the callers to do is not worry about column names and rely on the code to generate the names from the class.
                    foreach (PropertyInfo p in properties)
                    {
                        headerNames.Add(p.Name);
                    }
                }
                //Now lets create a worksheet with data.  First setup a few variables. 
                var worksheet = new Worksheet();
                int numCols = headerNames.Count;
                var columns = new Columns();
                //Here we can loop through our headernames and create the actual columns with associated style and header names.  
                for (int col = 0; col < numCols; col++)
                {
                    int width = headerNames[col].Length + 5;
                    Column c = new CustomColumn((UInt32)col + 1,
                                  (UInt32)numCols + 1, width);
                    columns.Append(c);
                }
                //Now append our new columns to the worksheet
                worksheet.Append(columns);
                //Now lets setup the sheets class so that we can append a sheet to add data too. 
                var sheets = new Sheets();
                //Create the sheet for the data with the name passed in and ID's needed to initialize the class
                var sheet = new Sheet { Name = sheetName, SheetId = 1, Id = relId };
                //Append the sheet we just created
                sheets.Append(sheet);
                //Now lets append our sheet to our workbook and get it prepared for the SheetData
                workbook.Append(fileVersion);
                workbook.Append(sheets);
                //Now lets actually create the data that will be appeneded to the workbook sheets. 
                SheetData sheetData = CreateSheetData(objects, headerNames);
                //Now that we have real data append that SheetData to our worksheet
                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                //Now lets save our new Document and close out our workbook. 
                worksheetPart.Worksheet.Save();
                myWorkbook.WorkbookPart.Workbook = workbook;
                myWorkbook.WorkbookPart.Workbook.Save();
                myWorkbook.Close();
            }

        }
        /// <summary>
        /// 
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <param name="objects">list of the object type</param>
        /// <param name="headerNames">Header names of the object</param>
        /// <returns></returns>
        private static SheetData CreateSheetData<T>(List<T> objects,
                       List<string> headerNames)
        {
            var sheetData = new SheetData();
            if (objects != null)
            {
                //Get fields names of object
                List<string> fields = GetPropertyInfo<T>();
                //Generate a list from A to AZ.  This translates to the columns in the excel workbook and is a way for us to keep track as we loop through the data
                var az = new List<Char>(Enumerable.Range('A', 'Z' -
                                      'A' + 1).Select(i => (Char)i).ToArray());
                List<string> columnLetters = new List<string>();
                columnLetters.AddRange(Enumerable.Range('A', 'Z' -
                                      'A' + 1).Select(i => ((char)i).ToString()));
                columnLetters.AddRange(Enumerable.Range('A', 'Z' -
                                      'A' + 1).Select(i => "A" + (char)i));

                //Lets setup a few variables to know how many columns, rows, and the headers we are working with. 
                List<string> headers = columnLetters.GetRange(0, fields.Count);
                int numRows = objects.Count;
                int numCols = fields.Count;
                var header = new Row();
                int index = 1;
                header.RowIndex = (uint)index;
                for (int col = 0; col < numCols; col++)
                {
                    //Here we are appending headers to the SheetData we are working with
                    var c = new HeaderCell(headers[col].ToString(),
                                           headerNames[col], index);
                    header.Append(c);
                }
                sheetData.Append(header);
                //Now lets get to the good stuff.  We need to loop through our data and insert each row one item at a time in the appropriate cell.  
                for (int i = 0; i < numRows; i++)
                {
                    index++;
                    //Lets grab the row from our list and place it in a variable.
                    var obj1 = objects[i];
                    //Next lets create the row we are wroking on so we can add data to it.  
                    var r = new Row { RowIndex = (uint)index };
                    //Now lets loop through our row and grab each value from the specific row in the list.  We can do this since we know how many columns it expects. 
                    for (int col = 0; col < numCols; col++)
                    {
                        string fieldName = fields[col];
                        //Grab the item from our row variable for the specific property we are working on at the moment.  
                        PropertyInfo myf = obj1.GetType().GetProperty(fieldName);
                        //Make sure the value is not null
                        if (myf != null)
                        {
                            //Get the actual value we need to append to the spreadsheet
                            object obj = myf.GetValue(obj1, null);
                            //As long as the value is not null we can go on to append it
                            if (obj != null)
                            {
                                //Lets figure out the type so we can create the proper cell type in Excel
                                //Once we know the type we can append the cell with the correct format.  TextCell, DateCell, FormattedNumberCell.
                                if (obj.GetType() == typeof(string))
                                {
                                    var c = new TextCell(headers[col].ToString(),
                                                obj.ToString(), index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(bool))
                                {
                                    string value =
                                      (bool)obj ? "Yes" : "No";
                                    var c = new TextCell(headers[col].ToString(),
                                                         value, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(DateTime))
                                {
                                    var c = new DateCell(headers[col].ToString(),
                                               (DateTime)obj, index);
                                    r.Append(c);
                                }
                                else if (obj.GetType() == typeof(decimal) ||
                                         obj.GetType() == typeof(double))
                                {
                                    var c = new FormatedNumberCell(
                                                 headers[col].ToString(),
                                                 obj.ToString(), index);
                                    r.Append(c);
                                }
                                else
                                {
                                    //Check if the value is of long type
                                    long value;
                                    if (long.TryParse(obj.ToString(), out value))
                                    {
                                        var c = new NumberCell(headers[col].ToString(),
                                                    obj.ToString(), index);
                                        r.Append(c);
                                    }
                                    else
                                    {
                                        //Otherwise just default to text
                                        var c = new TextCell(headers[col].ToString(),
                                                    obj.ToString(), index);
                                        r.Append(c);
                                    }
                                }
                            }
                        }
                    }
                    //Once we are done looping through our data lets add it to our sheet
                    sheetData.Append(r);
                }
                index++;
            }
            //Now return our new SheetData object to be added to our workbook
            return sheetData;
        }
        /// <summary>
        /// Takes a class and pulls the properties out of the class and returns the name in a list of strings.  
        /// </summary>
        /// <typeparam name="T">Object type to pass in</typeparam>
        /// <returns>List of names from the object passed in.</returns>
        private static List<string> GetPropertyInfo<T>()
        {
            PropertyInfo[] propertyInfos = typeof(T).GetProperties();
            return propertyInfos.Select(propertyInfo => propertyInfo.Name).ToList();
        }
    }
}