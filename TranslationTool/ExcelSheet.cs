using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace TranslationTool
{
    internal class ExcelSheet
    {
        public Dictionary<string, string> _IDandEnglishDictionary { get; set; }
        public Dictionary<string, string[]> _IDandArrayDictioinary { get; set; }
        // Dictionary to hold pretranslated items before they are in Revit.
        public Dictionary<string, string[]> _HoldDictioinary { get; set; }
        public List<string> CompareList { get; set; }

        //***********************************Read***********************************
        public Dictionary<string, string> Read(string path, int workSheet)
        {
            Dictionary<string, string> TranslationDictioinary_EnglishAndChinese = new Dictionary<string, string>();
            Dictionary<string, string[]> TranslationDictioinary_IdAndArray = new Dictionary<string, string[]>();
            Dictionary<string, string> TranslationDictioinary_IdAndEnglish = new Dictionary<string, string>();
            //Dictionary to not delete but hold translation in Excel until the model item is created.
            Dictionary<string, string[]> TranslationDictioinary_Hold = new Dictionary<string, string[]>();
            List<int> listToDelete = new List<int>();

            List<string> compareList = new List<string>();

            try
            {
                var package = new ExcelPackage(new FileInfo(path));

                ExcelWorksheet workSheet_Member = package.Workbook.Worksheets[workSheet];

                var start_Member = workSheet_Member.Dimension.Start;
                var end_Member = workSheet_Member.Dimension.End;

                int count = 2;
                while (workSheet_Member.Cells[count, 2].Text.Length > 0)
                {
                    string english = workSheet_Member.Cells[count, 2].Text;
                    //list to compare items to remove duplicates
                    compareList.Add(english);
                    count++;
                }

                int row = 2;
                while (workSheet_Member.Cells[row, 1].Text != "")
                {
                    try
                    {
                        string id = workSheet_Member.Cells[row, 1].Text;
                        string english = workSheet_Member.Cells[row, 2].Text;
                        string chinese = workSheet_Member.Cells[row, 3].Text;
                        string project = workSheet_Member.Cells[row, 4].Text;

                        // dictionary for english and chinese.
                        if (!TranslationDictioinary_EnglishAndChinese.ContainsKey(english))
                            TranslationDictioinary_EnglishAndChinese.Add(english, chinese);

                        // dictionary Id number to identify english word.
                        if (id != "" || id != null)
                            TranslationDictioinary_IdAndEnglish.Add(id, english);

                        string[] array = new string[3];
                        array[0] = english;
                        array[1] = chinese;
                        array[2] = project;

                        //If the first Char is $ do not remove from Excel.
                        //Added to new list to compare to project.
                        if (project == "")
                        {
                            TranslationDictioinary_Hold.Add(id, array);
                            continue;
                        }

                        TranslationDictioinary_IdAndArray.Add(id, array);
                    }
                    catch { }
                    row++;
                }

                // Close the Excel file.
                package.Stream.Close();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open " + path + " for reading. Exception raised - " + ex.Message);
            }

            _IDandEnglishDictionary = TranslationDictioinary_IdAndEnglish;
            _IDandArrayDictioinary = TranslationDictioinary_IdAndArray;
            _HoldDictioinary = TranslationDictioinary_Hold;

            CompareList = compareList;
            return TranslationDictioinary_EnglishAndChinese;
        }
  

        //***********************************Update***********************************
        public void Update(Dictionary<string, string[]> _IDandArrayDictioinary, List<string[]> NotTranslated, string path, int workSheet, string BuildingColor)
        {
            try
            {
                var package = new ExcelPackage(new FileInfo(path));

                ExcelWorksheet workSheet_Member = package.Workbook.Worksheets[workSheet];
                //End of rows.
                var end_Member = workSheet_Member.Dimension.End;
                //End of row plus one to get the empty cell.
                int row = end_Member.Row + 1;

                foreach (string[] array in NotTranslated)
                {
                    // If its part of a Z Sheet don't write to Excel. 
                    if (array[0].StartsWith("Z"))
                        continue;
                    // If the english is empty don't write to Excel. 
                    if (array[1] == "")
                        continue;

                    //Update existing cells in Excel.
                    //Check if Id and Array dictionary is not empty. 
                    if (_IDandArrayDictioinary != null)
                    {
                        //Check if the Id exist in the dictionary. 
                        if (_IDandArrayDictioinary.Keys.Contains(array[0]))
                        {
                            int r = 2;
                            //Iterate over existing Excel rows. 
                            while (workSheet_Member.Cells[r, 1].Text != "")
                            {
                                //If the Id matches the Excel Id update.
                                if (workSheet_Member.Cells[r, 1].Text == array[0])
                                {
                                    // Update all values to match. 
                                    UpdateRow(workSheet_Member, r, array, BuildingColor);

                                    continue;
                                }
                            }
                        }
                    }
                    //Append new item to Excel.
                    UpdateRow(workSheet_Member, row, array, BuildingColor);

                    row++;
                }

                // Close the Excel file.
                package.Save();
                package.Stream.Close();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error" + path + " for reading. Exception raised - " + ex.Message);
            }
        }

        //***********************************UpdateRow***********************************
        public void UpdateRow(ExcelWorksheet workSheet_Member, int row, string[] array, string BuildingColor)
        {
            // Id cell
            workSheet_Member.Cells[row, 1].Value = array[0];
            // English cell
            workSheet_Member.Cells[row, 2].Value = array[1];
            // Chinese cell
            workSheet_Member.Cells[row, 3].Value = array[2];
            // Project Name
            workSheet_Member.Cells[row, 4].Value = array[3];
            // Set color in Excel
            ExcelColor(workSheet_Member, row, BuildingColor);
        }

        //***********************************Delete***********************************
        public void Delete(string path, int workSheet, int k)
        {
            // Dictionary to check if any items have duplicates.
            Dictionary<string, string> TranslationDictioinary_RowandEnglish = new Dictionary<string, string>();
            // List of elements which have duplicates.
            List<int> listToDelete = new List<int>();
            try
            {
                var package = new ExcelPackage(new FileInfo(path));

                ExcelWorksheet workSheet_Member = package.Workbook.Worksheets[workSheet];

                var start_Member = workSheet_Member.Dimension.Start;
                var end_Member = workSheet_Member.Dimension.End;
                // Start on row 2
                for (int row = 2; row <= end_Member.Row; row++)
                {
                    try
                    {
                        string key = workSheet_Member.Cells[row, k].Text;
                        string project = workSheet_Member.Cells[row, 4].Text;
                        // Check if the english has been added to the dictionary.
                        // if the dictionary contains a duplicate english it adds to list to delete.
                        if (TranslationDictioinary_RowandEnglish.ContainsKey(key))
                            // if annotation don't compare building path.
                            listToDelete.Add(row);

                        // items added to dictionary if its unique.
                        else
                            TranslationDictioinary_RowandEnglish.Add(key, row.ToString());
                    }
                    catch { }
                }

                // list to delete.
                if (listToDelete.Count > 0)
                {
                    DeleteItems(listToDelete, workSheet_Member);
                }

                // Close and save the Excel file.
                package.Save();
                package.Stream.Close();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open " + path + " for reading. Exception raised - " + ex.Message);
            }
        }

        //***********************************Delete***********************************
        public void Delete(string path, int workSheet, int k, Dictionary<string, string[]> DeleteFromExcel)
        {
            // Dictionary to check if any items have duplicates.

            // List of elements which have duplicates.
            List<int> listToDelete = new List<int>();
            try
            {
                var package = new ExcelPackage(new FileInfo(path));

                ExcelWorksheet workSheet_Member = package.Workbook.Worksheets[workSheet];

                var start_Member = workSheet_Member.Dimension.Start;
                var end_Member = workSheet_Member.Dimension.End;
                // Start on row 2
                for (int row = 2; row <= end_Member.Row; row++)
                {
                    try
                    {
                        string key = workSheet_Member.Cells[row, k].Text;
                        if (DeleteFromExcel.ContainsKey(key))
                            listToDelete.Add(row);
                    }
                    catch { }
                }

                // list to delete.
                if (listToDelete != null)
                {
                    DeleteItems(listToDelete, workSheet_Member);
                }

                // Close and save the Excel file.
                package.Save();
                package.Stream.Close();
                package.Dispose();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Cannot open " + path + " for reading. Exception raised - " + ex.Message);
            }
        }

        //***********************************Delete***********************************
        public void DeleteItems(List<int> ListToDelete, ExcelWorksheet workSheet)
        {
            ListToDelete = ListToDelete.OrderByDescending(v => v).ToList();
            foreach (int row in ListToDelete)
                workSheet.DeleteRow(row, 1, true);
        }

        //***********************************ExcelColor***********************************
        public void ExcelColor(ExcelWorksheet workSheet_Member, int row, string BuildingColor)
        {            
            Color lightGreen = System.Drawing.Color.FromArgb(144,238,144);
            Color lightYellow = System.Drawing.Color.FromArgb(255, 255, 153);
            Color lightPurple = System.Drawing.Color.FromArgb(230, 230, 230);
            if (BuildingColor != "")
            {
                switch (BuildingColor)
                {
                    case "1":
                        SetExcelColor(workSheet_Member, row, lightGreen);
                        break;
                    case "2":
                        SetExcelColor(workSheet_Member, row, lightYellow);
                        break;
                    case "3":
                        SetExcelColor(workSheet_Member, row, lightPurple);
                        break;
                }
            }

            if (BuildingColor == "")
            {
                
            }
        }

        //***********************************ExcelColor***********************************
        public void SetExcelColor(ExcelWorksheet workSheet_Member, int row, Color color)
        {
            // Id cell
            workSheet_Member.Cells[row, 1].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet_Member.Cells[row, 1].Style.Fill.BackgroundColor.SetColor(color);
            // English cell
            workSheet_Member.Cells[row, 2].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet_Member.Cells[row, 2].Style.Fill.BackgroundColor.SetColor(color);
            // Chinese cell
            workSheet_Member.Cells[row, 3].Style.Fill.PatternType = ExcelFillStyle.Solid;
            workSheet_Member.Cells[row, 3].Style.Fill.BackgroundColor.SetColor(color);
        }
    }
}