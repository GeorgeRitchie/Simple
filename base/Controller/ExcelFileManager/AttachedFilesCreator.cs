using System;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;

namespace _base.Controller
{
	class AttachedFilesCreator
	{
		private Excel.Application appCopyTo = null;
		private Excel.Workbook workbookCopyTo = null;
		private Excel.Worksheet worksheetCopyTo = null;

		public void CreateAttachedFile(Excel.Range dataRange, string fullFileName, int fileFormat)
		{
			Create(SelectOnlyDataRangeWithoutTriggersRows(dataRange), fullFileName, fileFormat);
		}
		private Excel.Range SelectOnlyDataRangeWithoutTriggersRows(Excel.Range Range)
		{
			Excel.Range firstCellUnderStartTriggerCell = Range.Item[2, 1];
			Excel.Range firstCellAboveEndTriggerCell = Range.Item[Range.Rows.Count - 1, Range.Columns.Count];
			Excel.Range DataRangeWithoutTriggersRows = Range.Worksheet.Range[firstCellUnderStartTriggerCell, firstCellAboveEndTriggerCell];

			return DataRangeWithoutTriggersRows;
		}



		private void Create(Excel.Range dataRange, string fullFileName, int fileFormat)
		{
			try
			{
				if (fileFormat == 0)
				{
					ExportRangeAsPNG(dataRange, fullFileName);
				}

				OpenExcelProgramForCopying();
				Copy(dataRange);
				Paste();
				Save(fullFileName, fileFormat);
			}
			catch (Exception e)
			{
				Logger.Log("The exception was thrown during creating attached file for a receiver. Creating was not successful! Exception message: " + e.Message, typeof(AttachedFilesCreator).FullName + '.' + nameof(Create));
			}
			finally
			{
				QuitExcelProgramForCopying();
			}
		}

		private void OpenExcelProgramForCopying()
		{
			appCopyTo = new Excel.Application();
			appCopyTo.Visible = false;
			appCopyTo.AutomationSecurity = MsoAutomationSecurity.msoAutomationSecurityForceDisable;
			appCopyTo.CSVDisplayNumberConversionWarning = false;

			appCopyTo.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
			workbookCopyTo = appCopyTo.Workbooks[1];
		}

		private void Copy(Excel.Range dataRange)
		{
			dataRange.Copy();
		}

		private void Paste()
		{
			workbookCopyTo.Activate();

			worksheetCopyTo = workbookCopyTo.Worksheets[1];
			try
			{
				((Excel.Range)worksheetCopyTo.Cells[1, 1]).PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
			}
			catch
			{
				((Excel.Range)worksheetCopyTo.Cells[1, 1]).PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
			}
		}

		private void Save(string fullFileName, int fileFormat)
		{
			workbookCopyTo.SaveAs2(Filename: fullFileName, FileFormat: fileFormat);
		}

		private void QuitExcelProgramForCopying()
		{
			if (appCopyTo != null)
			{
				if (appCopyTo.Workbooks.Count != 0)
				{
					appCopyTo.Workbooks[1].Close(SaveChanges: false);
				}

				appCopyTo.Quit();

				if (worksheetCopyTo != null)
					Marshal.FinalReleaseComObject(worksheetCopyTo);
				if (workbookCopyTo != null)
					Marshal.FinalReleaseComObject(workbookCopyTo);
				if (appCopyTo != null)
					Marshal.FinalReleaseComObject(appCopyTo);
			}
			GC.Collect();
			GC.WaitForPendingFinalizers();
		}



		private void ExportRangeAsPNG(Excel.Range rangeToPNG, string fullFileName)
		{
			//rangeToPNG.PrintOut(1, 1, 1, false, Type.Missing, true, Type.Missing, fullFileName);



			//Excel.ChartObject chartObj = worksheet.ChartObjects().Add(rangeToPNG.Left, rangeToPNG.Top, rangeToPNG.Width, rangeToPNG.Height);

			//chartObj.Activate();
			//Excel.Chart chart = chartObj.Chart;
			//chart.Paste();
			//chart.Export(fullFileName, "PNG");
			//chartObj.Delete();
		}
	}
}
