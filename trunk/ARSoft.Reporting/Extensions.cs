namespace ARSoft.NPOI
{
    using global::NPOI.HSSF.UserModel;
    using global::NPOI.SS.UserModel;
    using global::NPOI.SS.Util;

    public static class Extensions
    {
        public static void CopyRow(this ISheet worksheet, ISheet sourceSheet, int sourceRowNum, int destinationRowNum)
        {
            // Get the source / new row
            var newRow = worksheet.GetRow(destinationRowNum);
            var sourceRow = sourceSheet.GetRow(sourceRowNum);

            // If the row exist in destination, push down all rows by 1 else create a new row
            if (newRow != null)
            {
                worksheet.ShiftRows(destinationRowNum, worksheet.LastRowNum, 1);
            }
            else
            {
                newRow = worksheet.CreateRow(destinationRowNum);
            }

            // Loop through source columns to add to new row
            for (int i = 0; i < sourceRow.LastCellNum; i++)
            {
                // Grab a copy of the old/new cell
                var oldCell = sourceRow.GetCell(i);
                var newCell = newRow.CreateCell(i);

                // If the old cell is null jump to next cell
                if (oldCell == null)
                {
                    newCell = null;
                    continue;
                }

                //newCell.CellStyle = oldCell.CellStyle;
                // Copy style from old cell and apply to new cell
                //var newCellStyle = worksheet.Workbook.CreateCellStyle();
                //newCellStyle.CloneStyleFrom(oldCell.CellStyle);
                //newCell.CellStyle = newCellStyle;

                // If there is a cell comment, copy
                if (newCell.CellComment != null) newCell.CellComment = oldCell.CellComment;

                // If there is a cell hyperlink, copy
                if (oldCell.Hyperlink != null) newCell.Hyperlink = oldCell.Hyperlink;

                // Set the cell data type
                newCell.SetCellType(oldCell.CellType);

                // Set the cell data value
                switch (oldCell.CellType)
                {
                    case CellType.BLANK:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                    case CellType.BOOLEAN:
                        newCell.SetCellValue(oldCell.BooleanCellValue);
                        break;
                    case CellType.ERROR:
                        newCell.SetCellErrorValue(oldCell.ErrorCellValue);
                        break;

                    //case CellType.FORMULA:
                    //    newCell.sSetCellFormula(oldCell.CellFormula);
                    //    break;
                    case CellType.NUMERIC:
                        newCell.SetCellValue(oldCell.NumericCellValue);
                        break;
                    case CellType.STRING:
                        newCell.SetCellValue(oldCell.RichStringCellValue);
                        break;
                    case CellType.Unknown:
                        newCell.SetCellValue(oldCell.StringCellValue);
                        break;
                }
            }

            // If there are are any merged regions in the source row, copy to new row
            for (int i = 0; i < sourceSheet.NumMergedRegions; i++)
            {
                CellRangeAddress cellRangeAddress = worksheet.GetMergedRegion(i);
                if (cellRangeAddress.FirstRow == sourceRow.RowNum)
                {
                    CellRangeAddress newCellRangeAddress = new CellRangeAddress(newRow.RowNum,
                                                                                newRow.RowNum +
                                                                                 (cellRangeAddress.FirstRow -
                                                                                  cellRangeAddress.LastRow),
                                                                                cellRangeAddress.FirstColumn,
                                                                                cellRangeAddress.LastColumn);
                    worksheet.AddMergedRegion(newCellRangeAddress);
                }
            }

        }

        /// <summary>
        /// HSSFRow Copy Command
        ///
        /// Description:  Inserts a existing row into a new row, will automatically push down
        ///               any existing rows.  Copy is done cell by cell and supports, and the
        ///               command tries to copy all properties available (style, merged cells, values, etc...)
        /// </summary>
        /// <param name="workbook">Workbook containing the worksheet that will be changed</param>
        /// <param name="worksheet">WorkSheet containing rows to be copied</param>
        /// <param name="sourceRowNum">Source Row Number</param>
        /// <param name="destinationRowNum">Destination Row Number</param>
        public static void CopyRow(this ISheet worksheet, int sourceRowNum, int destinationRowNum)
        {
            CopyRow(worksheet, worksheet, sourceRowNum, destinationRowNum);
        }

        /**
         * Remove a row by its index
         * @param sheet a Excel sheet
         * @param rowIndex a 0 based index of removing row
         */
        public static void RemoveRow(this ISheet sheet, int rowIndex)
        {
            int lastRowNum = sheet.LastRowNum;
            if (rowIndex >= 0 && rowIndex < lastRowNum)
            {
                sheet.ShiftRows(rowIndex + 1, lastRowNum, -1);
            }
            if (rowIndex == lastRowNum)
            {
                var removingRow = sheet.GetRow(rowIndex);
                if (removingRow != null)
                {
                    sheet.RemoveRow(removingRow);
                }
            }
        }
    }
}
