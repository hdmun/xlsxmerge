using XlsxMerge.Extensions;
using XlsxMerge.Features.Diffs;
using XlsxMerge.Features.Excels;
using XlsxMerge.Features.Merges;
using XlsxMerge.ViewModel;

namespace XlsxMerge
{
    class MergeResultPreviewer
	{
        private static readonly Font StrikeOutFont = new (SystemFonts.DefaultFont, FontStyle.Strikeout);

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

            // iterate rows
			dataGridView.RowCount = cachedTempPreviewLines?.Count() ?? 0;
			for (int currentRowIdx = 0; currentRowIdx < dataGridView.RowCount; currentRowIdx++)
			{
				var eachRow = cachedTempPreviewLines[currentRowIdx];
				DataGridViewRow dgvRow = dataGridView.Rows[currentRowIdx];

                // set 'hunk_no' cell
                var hunkNoCell = dgvRow.Cells["hunk_no"];
                hunkNoCell.Value = "";
				{
					var candidateHunkIdx = previewData.GetHunkIdxByRowNumber(currentRowIdx);
                    if (candidateHunkIdx >= 0)
                    {
                        if (sheetMergeDecision.HunkMergeDecisionList[candidateHunkIdx].DocMergeOrder == null)
                            hunkNoCell.Style.SelectionBackColor = Color.Red;
                        else
                            hunkNoCell.Style.SelectionBackColor = Color.LightSlateGray;

                        hunkNoCell.Value = $"#{candidateHunkIdx + 1}";
                    }
				}

                // parse token
                string[] token = eachRow.Split(new char[] { ':' });
                if (token.Length == 1)
				{
					dgvRow.DefaultCellStyle.BackColor = Color.Yellow;
					dgvRow.Cells["source_line"].Value = token.First();
					continue;
				}

                // TODO: 정리필요
                string sourceLineText = token.First();
				if (token.Length > 1)
					sourceLineText = sourceLineText + $": {token[1]}";

                bool isRemovedLine = false;

                int refBaseRowNumber = int.MinValue;
				if (token.Length > 2)
                {
                    var rowNumberToken = token[2];
                    refBaseRowNumber = int.Parse(rowNumberToken);

                    isRemovedLine = rowNumberToken.EndsWith("-1");
                }

                if (isRemovedLine)
                {
                    sourceLineText = sourceLineText + " [-]";
                    dgvRow.DefaultCellStyle.Font = StrikeOutFont;
                }

                string firstToken = token.First();
                var docOrigin = firstToken.ToDocOrigin();
                if (docOrigin is null)
                    sourceLineText = "=";

                // set 'source_line' cell
                dgvRow.Cells["source_line"].Value = sourceLineText;

                // set back color
                Color backColor = docOrigin switch
                {
                    DocOrigin.Base => ColorScheme.BaseBackground,
                    DocOrigin.Mine => ColorScheme.MineBackground,
                    DocOrigin.Theirs => ColorScheme.TheirsBackground,
                    _ => Color.White
                };
                dgvRow.DefaultCellStyle.BackColor = backColor;

                // get refWorksheet
                var refWorksheet = docOrigin switch
                {
                    DocOrigin.Base => parsedWorksheetData[DocOrigin.Base],
                    DocOrigin.Mine => parsedWorksheetData[DocOrigin.Mine],
                    DocOrigin.Theirs => parsedWorksheetData[DocOrigin.Theirs],
                    _ => parsedWorksheetData[DocOrigin.Base]
                };
                if (refWorksheet?.RowCount == 0)
					continue;

                // iterate columns
                int rowNumber = int.Parse(token[1]);
                var baseWorksheet = parsedWorksheetData[DocOrigin.Base];
                UpdateColumns(rowNumber, refWorksheet, refBaseRowNumber, baseWorksheet, dgvRow);
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

        private static void UpdateColumns(int rowNumber, ExcelWorksheet? refWorksheet, int refBaseRowNumber, ExcelWorksheet baseWorksheet, DataGridViewRow dgvRow)
        {
            int maxColumn = refWorksheet?.ColumnCount ?? 0;
            for (int cellNumber = 1; cellNumber <= maxColumn; cellNumber++)
            {
                var currentCell = refWorksheet.Cell(rowNumber, cellNumber);
                var columnName = $"C{cellNumber}";
                var currentCellDgv = dgvRow.Cells[columnName];
                currentCellDgv.Value = currentCell.Value2String;

                if (refBaseRowNumber <= 0 || baseWorksheet == null)
                    continue;

                var baseCell = baseWorksheet.Cell(refBaseRowNumber, cellNumber);
                if (currentCell.ContentsForDiff3 == baseCell.ContentsForDiff3)
                    continue;

                if (dgvRow.DefaultCellStyle.BackColor == ColorScheme.MineBackground)
                    currentCellDgv.Style.BackColor = ColorScheme.MineHighlight;
                if (dgvRow.DefaultCellStyle.BackColor == ColorScheme.TheirsBackground)
                    currentCellDgv.Style.BackColor = ColorScheme.TheirsHighlight;
            }
        }
	}
}
