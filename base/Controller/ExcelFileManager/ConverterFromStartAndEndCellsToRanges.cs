using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace _base.Controller
{
	class ConverterFromStartAndEndCellsToRanges
	{
		class ListOfClassifiedCellsByDimention
		{
			public List<List<Excel.Range>> classifiedCellList { get; private set; } = new List<List<Excel.Range>>();

			public void CreateNewDemention()
			{
				classifiedCellList.Add(new List<Excel.Range>());
			}

			public void AddCellToCurrentDemention(Excel.Range cell)
			{
				classifiedCellList.Last().Add(cell);
			}
		}


		public List<Excel.Range> GetRanges(List<Excel.Range> StartCells, List<Excel.Range> EndCells)
		{
			ValidateStartAndEndCells(StartCells, EndCells);

			List<List<Excel.Range>> startCellList = GetCellsClassifiedByItsDemension(StartCells);
			List<List<Excel.Range>> endCellList = GetCellsClassifiedByItsDemension(EndCells);

			return MakeRangeFromStartAndEndCells(startCellList, endCellList);
		}

		private void ValidateStartAndEndCells(List<Excel.Range> StartCells, List<Excel.Range> EndCells)
		{
			if (StartCells == null || StartCells.Count == 0)
			{
				throw new ArgumentNullException("StartCells is null or does not contain anything");
			}

			if (EndCells == null || EndCells.Count == 0)
			{
				throw new ArgumentNullException("EndCells is null or does not contain anything");
			}
		}

		//TODO [!] logic is not totally safe for random located data areas
		private List<List<Excel.Range>> GetCellsClassifiedByItsDemension(List<Excel.Range> listOfCells)
		{
			ListOfClassifiedCellsByDimention listOfClassifiedCells = new ListOfClassifiedCellsByDimention();
			Excel.Range previousCell = null;

			foreach (Excel.Range currentCell in listOfCells.OrderBy(u => u.Column).ThenBy(u => u.Row))
			{
				if (ShouldCreateNewDimention(previousCell, currentCell))
				{
					listOfClassifiedCells.CreateNewDemention();
				}

				listOfClassifiedCells.AddCellToCurrentDemention(currentCell);
				previousCell = currentCell;
			}

			return listOfClassifiedCells.classifiedCellList;
		}

		private bool ShouldCreateNewDimention(Excel.Range previousCell, Excel.Range currentCell)
		{
			return previousCell == null || previousCell.Column < currentCell.Column;
		}

		private List<Excel.Range> MakeRangeFromStartAndEndCells(List<List<Excel.Range>> startCellList, List<List<Excel.Range>> endCellList)
		{
			List<Excel.Range> ranges = new List<Excel.Range>();
			Excel.Worksheet worksheet = GetWorksheetFromCell(startCellList[0][0]);

			for (int i = 0; i < startCellList.Count; i++)
			{
				for (int j = 0; j < startCellList[i].Count; j++)
				{
					ranges.Add(worksheet.Range[worksheet.Cells[startCellList[i][j].Row, startCellList[i][j].Column], worksheet.Cells[endCellList[i][j].Row, endCellList[i][j].Column]]);
				}
			}

			return ranges;
		}

		private Excel.Worksheet GetWorksheetFromCell(Excel.Range cell)
		{
			return cell.Worksheet;
		}
	}
}
