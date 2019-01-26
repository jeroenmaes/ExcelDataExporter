﻿using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using System.Diagnostics;
using ExcelToXmlParser.Helpers;
using System.Xml.Serialization;
using System.Text;

namespace ExcelToXmlParser.Services
{
    public class ExcelConverter : IExcelConverter
    {
        private DataTable ReadExcelFile(Stream contents)
        {
            // Initialize an instance of DataTable
            DataTable dt = new DataTable();

            try
            {
                var wb = new XLWorkbook(contents, XLEventTracking.Disabled);
                var ws = wb.Worksheets.First();

                foreach (var cell in ws.RowsUsed().First().CellsUsed())
                {
                    var columnName = "_" + cell.Value.ToString().ToAlphaNumericOnly();
                    dt.Columns.Add(columnName);                    
                }

                foreach (var row in ws.RowsUsed().Where(x => x.RowNumber() > 1))
                {
                    DataRow temprow = dt.NewRow();

                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        try
                        {
                            var cellValue = row.Cells().ElementAt(i).Value.ToString();
                            temprow[i] = cellValue.TrimWhiteSpaces().RemoveLineBreaks();
                        }
                        catch (Exception)
                        {
                            Trace.WriteLine("Empty Cell");                            
                        }
                    }

                    dt.Rows.Add(temprow);
                }

                wb.Dispose();

                return dt;
            }
            catch (IOException ex)
            {
                throw new IOException(ex.Message);
            }
        }

        /// <summary>
        /// Convert DataTable to Xml format
        /// </summary>
        /// <param name="filePath">Excel File Path</param>
        /// <returns>Xml format string</returns>
        public string GetXML(Stream contents)
        {
            using (DataSet ds = new DataSet())
            {
                ds.Tables.Add(this.ReadExcelFile(contents));

                return ds.GetXml();
            }
        }

        public string GetXML(Stream contents, string rootNodeName, string parentNodeName, string XmlNamespace)
        {
            using (DataSet ds = new DataSet())
            {
                ds.Namespace = XmlNamespace;
                ds.DataSetName = rootNodeName;

                var table = this.ReadExcelFile(contents);
                table.TableName = parentNodeName;

                ds.Tables.Add(table);

                return ds.GetXml();
            }
        }                
    }
}
