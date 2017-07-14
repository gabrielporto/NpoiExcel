using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//add
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.HPSF;

namespace NPOI
{
    class Excel
    {
        HSSFWorkbook hssfworkbook;

        public void Carregar(string fileStreamPath = null)
        {
            if (fileStreamPath != null)
            {
                FileStream file = new FileStream(fileStreamPath, FileMode.Open, FileAccess.Read);

                HSSFWorkbook hssfworkbook = new HSSFWorkbook(file);

                if (hssfworkbook != this.hssfworkbook)
                {
                    this.hssfworkbook = hssfworkbook;
                }
            }
        }

        public DataTable GetDataTable(string fileStreamPath = null)
        {
            DataTable dt = new DataTable();

            if (fileStreamPath != null)
            {
                Carregar(fileStreamPath);
                return GetDataTable();
            }
            else
            {
                //A tabela
                ISheet sheet = hssfworkbook.GetSheetAt(0);

                int numeroDeLinhas = sheet.PhysicalNumberOfRows;
                int numeroDeColunas = 0;

                //Achar o numero de colunas
                for (int i = 0; i < sheet.PhysicalNumberOfRows; i++)
                {
                    //IRow row = sheet.GetRow(i);
                    //int cellsCont = row.Cells.Count();

                    if (sheet.GetRow(i) != null && sheet.GetRow(i).Cells.Count() > numeroDeColunas)
                    {
                        numeroDeColunas = sheet.GetRow(i).Cells.Count();
                    }
                }

                //adicionar as colunas
                for (int i = 0; i < numeroDeColunas; i++)
                {
                    dt.Columns.Add();
                }

                //Pegar os valores do sheet e colcoar no DT
                for (int r = 0; r < sheet.PhysicalNumberOfRows; r++)
                {
                    dt.Rows.Add(dt.NewRow());
                    for (int c = 0; c < numeroDeColunas; c++)
                    {

                        IRow row = sheet.GetRow(r);
                        if (sheet.GetRow(r) == null || sheet.GetRow(r).GetCell(c) == null)
                        {
                            dt.Rows[r][c] = "";
                        }
                        else
                        {
                            ICell cell = sheet.GetRow(r).GetCell(c);

                            //Verificar qual o tipo de valor esta na tabela
                            switch (sheet.GetRow(r).GetCell(c).CellType)
                            {
                                case CellType.Numeric:
                                    dt.Rows[r][c] = sheet.GetRow(r).GetCell(c).NumericCellValue.ToString();
                                    break;
                                case CellType.String:
                                    dt.Rows[r][c] = sheet.GetRow(r).GetCell(c).StringCellValue;
                                    break;
                                case CellType.Blank:
                                    //dt.Rows[r][c] = CellType.Blank;
                                    dt.Rows[r][c] = "";
                                    break;
                                //Caso seja uma Formula Verificar qual valor tem na formula
                                case CellType.Formula:
                                    switch (sheet.GetRow(r).GetCell(c).CachedFormulaResultType)
                                    {
                                        //Caso seja numerico
                                        case CellType.Numeric:
                                            dt.Rows[r][c] = sheet.GetRow(r).GetCell(c).NumericCellValue.ToString();
                                            break;
                                    }
                                    break;
                                default:
                                    dt.Rows[r][c] = "T" + sheet.GetRow(r).GetCell(c).CellType;
                                    break;
                            }
                        }
                    }

                }
            }
            return dt;
        }

        public void SalvarDataTable(DataTable dt, string path)
        {
            HSSFWorkbook hssfworkbook = new HSSFWorkbook();

            //Criar uma tabela em branco
            ISheet sheet = hssfworkbook.CreateSheet();

            //Colocar dados nela
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                sheet.CreateRow(i);
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    sheet.GetRow(i).CreateCell(j).SetCellValue(dt.Rows[i][j].ToString());
                }
            }

            hssfworkbook.CreateSheet();
            hssfworkbook.SetSheetName(hssfworkbook.GetSheetIndex(sheet), sheet.SheetName);

            ////create a entry of DocumentSummaryInformation
            DocumentSummaryInformation dsi = PropertySetFactory.CreateDocumentSummaryInformation();
            dsi.Company = "NPOI Team";
            hssfworkbook.DocumentSummaryInformation = dsi;

            ////create a entry of SummaryInformation
            SummaryInformation si = PropertySetFactory.CreateSummaryInformation();
            si.Subject = "NPOI SDK Example";
            hssfworkbook.SummaryInformation = si;

            //Write the stream data of workbook to the root directory
            FileStream file = new FileStream(path, FileMode.Create);
            hssfworkbook.Write(file);
            file.Close();
        }
    }
}