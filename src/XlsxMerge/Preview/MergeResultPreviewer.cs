using XlsxMerge.Features.Diffs;
using XlsxMerge.Features.Merges;
using XlsxMerge.ViewModel;

namespace XlsxMerge
{
    class MergeResultPreviewer
	{
		public static void RefreshDataGridViewContents(
            DiffViewModel diffViewModel,
			SheetMergeDecision sheetMergeDecision,
			DataGridView dataGridView,
			MergeResultPreviewData previewData)
		{
			if (sheetMergeDecision == null)
				return;

			var sheetResult = sheetMergeDecision.SheetDiffResult;
			var parsedWorksheetData = diffViewModel.GetWorksheets(sheetResult.WorksheetName);

            // 열 생성
            dataGridView.Rows.Clear();
            dataGridView.Columns.Clear();

            var columnWidthList = diffViewModel.CalcColumnWidthList(sheetResult.WorksheetName);
            var columns = MakeColumns(columnWidthList);
            dataGridView.Columns.AddRange(columns.ToArray());

            if (columnWidthList.Count == 0)
            {
                int rowIndex = dataGridView.Rows.Add();
                dataGridView.Rows[rowIndex].Cells["C1"].Value = "(빈 워크시트입니다.)";
                return;
            }

			var cachedTempPreviewLines = previewData.RowInfoList;

			Font strikeoutFont = new Font(SystemFonts.DefaultFont, FontStyle.Strikeout);
			dataGridView.RowCount = cachedTempPreviewLines.Count();
			for (int currentRowIdx = 0; currentRowIdx < cachedTempPreviewLines.Count(); currentRowIdx++)
			{
				var eachRow = cachedTempPreviewLines[currentRowIdx];
				var dgvRow = dataGridView.Rows[currentRowIdx];

				string[] token = eachRow.Split(new char[] { ':' });
				dgvRow.Cells["hunk_no"].Value = "";
				{
					var candidateHunkIdx = previewData.GetHunkIdxByRowNumber(currentRowIdx);
                    if (candidateHunkIdx >= 0)
                    {
                        if (sheetMergeDecision.HunkMergeDecisionList[candidateHunkIdx].DocMergeOrder == null)
                            dgvRow.Cells["hunk_no"].Style.SelectionBackColor = Color.Red;
                        else
                            dgvRow.Cells["hunk_no"].Style.SelectionBackColor = Color.LightSlateGray;
                        dgvRow.Cells["hunk_no"].Value = $"#{candidateHunkIdx + 1}";
                    }
				}

				string sourceLineText = token[0];
				if (token.Length == 1)
				{
					dgvRow.DefaultCellStyle.BackColor = Color.Yellow;
					dgvRow.Cells["source_line"].Value = sourceLineText;
					continue;
				}

				if (token.Length > 1)
					sourceLineText = sourceLineText + $": {token[1]}";

				bool isRemovedLine = token.Length > 2 && token[2].EndsWith("-1");
                if (isRemovedLine)
                    sourceLineText = sourceLineText + " [-]";
                int refBaseRowNumber = int.MinValue;
				if (token.Length > 2)
					refBaseRowNumber = int.Parse(token[2]);

				Color backColor = Color.White;
				int rowNumber = int.Parse(token[1]);
				var refWorksheet = parsedWorksheetData[DocOrigin.Base];
				if (token[0].StartsWith("base"))
				{
					refWorksheet = parsedWorksheetData[DocOrigin.Base];
					backColor = ColorScheme.BaseBackground;
				}
				else if (token[0].StartsWith("mine"))
				{
					refWorksheet = parsedWorksheetData[DocOrigin.Mine];
					backColor = ColorScheme.MineBackground;
				}
				else if (token[0].StartsWith("theirs"))
				{
					refWorksheet = parsedWorksheetData[DocOrigin.Theirs];
					backColor = ColorScheme.TheirsBackground;
				}
				else
				{
					sourceLineText = "=";
				}

				dgvRow.DefaultCellStyle.BackColor = backColor;
				if (isRemovedLine)
					dgvRow.DefaultCellStyle.Font = strikeoutFont;
				dgvRow.Cells["source_line"].Value = sourceLineText;

				if (refWorksheet.RowCount == 0)
					continue;


				int maxColumn = refWorksheet.ColumnCount;
				for (int cellNumber = 1; cellNumber <= maxColumn; cellNumber++)
				{
					var currentCell = refWorksheet.Cell(rowNumber, cellNumber);
					var currentCellDgv = dgvRow.Cells["C" + cellNumber.ToString()];
					currentCellDgv.Value = currentCell.Value2String;

					if (refBaseRowNumber <= 0 || parsedWorksheetData[DocOrigin.Base] == null)
						continue;

					var baseCell = parsedWorksheetData[DocOrigin.Base].Cell(refBaseRowNumber, cellNumber);
					if (currentCell.ContentsForDiff3 == baseCell.ContentsForDiff3)
						continue;

					if (dgvRow.DefaultCellStyle.BackColor == ColorScheme.MineBackground)
						currentCellDgv.Style.BackColor = ColorScheme.MineHighlight;
					if (dgvRow.DefaultCellStyle.BackColor == ColorScheme.TheirsBackground)
						currentCellDgv.Style.BackColor = ColorScheme.TheirsHighlight;
				}
			}
		}

        private static IEnumerable<DataGridViewTextBoxColumn> MakeColumns(List<double> columnWidthList)
        {
            var columns = new List<DataGridViewTextBoxColumn>();
            // 기본 열
            var defaultColumn = new DataGridViewTextBoxColumn
            {
                Name = "hunk_no",
                HeaderText = "변경 위치",
                Width = 80,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            columns.Add(defaultColumn);

            defaultColumn = new DataGridViewTextBoxColumn
            {
                Name = "source_line",
                HeaderText = "소스(행)",
                Width = 80,
                SortMode = DataGridViewColumnSortMode.NotSortable
            };
            columns.Add(defaultColumn);

            if (columnWidthList.Count == 0)
            {
                // 모두 빈 워크시트입니다.
                var emptyColumn = new DataGridViewTextBoxColumn
                {
                    Name = "C1",
                    HeaderText = "정보",
                    Width = 400,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                columns.Add(emptyColumn);
                return columns;
            }

            for (int i = 0; i < columnWidthList.Count; i++)
            {
                int columnNo = i + 1;
                string columnText = HelperFunctions.GetExcelColumnName(columnNo);
                int columnWidth = (int)(columnWidthList[i]);

                var dataGridViewTextBoxColumn = new DataGridViewTextBoxColumn
                {
                    Name = $"C{columnNo}",
                    HeaderText = $"{columnText}::[]",
                    Width = columnWidth,
                    SortMode = DataGridViewColumnSortMode.NotSortable
                };
                columns.Add(dataGridViewTextBoxColumn);
            }

            return columns;
        }
	}
}
