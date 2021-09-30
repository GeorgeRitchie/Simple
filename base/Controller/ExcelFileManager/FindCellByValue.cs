using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace _base.Controller
{
	class FindCellByValue
	{
		private Excel.Range TotalRanges;

		public List<Excel.Range> FoundCells { get; set; } = null;

		public FindCellByValue(Excel.Range TotalRanges)
		{
			this.TotalRanges = TotalRanges;
		}

		public void FindCells(string findWhat)
		{
			List<Excel.Range> listOfFoundCells = new List<Excel.Range>();

			try
			{
				listOfFoundCells.Add(GetFirstFound(findWhat));
				while (true)
				{
					listOfFoundCells.Add(GetNextFound(findWhat, listOfFoundCells));
				}
			}
			catch
			{ }

			FoundCells = listOfFoundCells;
		}

		private Excel.Range GetFirstFound(string findWhat)
		{
			Excel.Range first = TotalRanges.Find(findWhat);
			if (first == null)
			{
				throw new InvalidOperationException($"There is no cell with {findWhat} value");
			}

			return first;
		}

		private Excel.Range GetNextFound(string findWhat, List<Excel.Range> listOfFoundCells)
		{
			Excel.Range nextCell = TotalRanges.FindNext(listOfFoundCells.LastOrDefault());

			if (nextCell.Address == listOfFoundCells.FirstOrDefault().Address)
			{
				throw new InvalidOperationException($"There is no next cell with {findWhat} value");
			}

			return nextCell;
		}

		public bool IsAnyCellFound()
		{
			return FoundCells != null && FoundCells.Count > 0;
		}
	}
}
