using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using XlsxMerge;
using XlsxMerge.Extensions;
using XlsxMerge.Diff;
using XlsxMerge.ViewModel;

namespace XlsxMerge.View
{
	public partial class FormMainDiff : Form
	{
		private readonly PathViewModel _pathViewModel;

        public FormMainDiff(PathViewModel pathViewModel)
		{
			InitializeComponent();

            // UI 준비
            labelPathBase.BackColor = ColorScheme.BaseBackground;
            labelPathMine.BackColor = ColorScheme.MineBackground;
            labelPathTheirs.BackColor = ColorScheme.TheirsBackground;
            labelPathResult.BackColor = ColorScheme.DiffHunk;

            _pathViewModel = pathViewModel;
		}

		XlsxMergeDecision _xlsxMergeDecision = new XlsxMergeDecision();
		Dictionary<String, MergeResultPreviewData> previewDataCache = new Dictionary<string, MergeResultPreviewData>();
		public bool MergeSuccessful = false;

		private bool _dataGridViewCellUpdatingInProgress = false;
		private int _focusedHunkIdx = -1;


        private void FormMainDiff_Load(object sender, EventArgs e)
        {
            this.Text = VersionName.GetFormTitleText();

            // 데이터 바인딩
            labelPathBase.BindingText(_pathViewModel, nameof(_pathViewModel.BasePathLabelText));
            labelPathMine.BindingText(_pathViewModel, nameof(_pathViewModel.MinePathLabelText));

            labelPathTheirs.BindingVisible(_pathViewModel, nameof(_pathViewModel.VisibleTheirsPath));
            labelPathTheirs.BindingText(_pathViewModel, nameof(_pathViewModel.TheirsPathLabelText));

            labelPathResult.BindingVisible(_pathViewModel, nameof(_pathViewModel.VisibleResultPath));
            labelPathResult.BindingText(_pathViewModel, nameof(_pathViewModel.ResultPathLabelText));

            if (_pathViewModel.VisibleResultPath)
                buttonSaveMergeResult.Text = "머지 결과 저장 후 닫기";

            panelTop.Height = labelPathResult.Bounds.Bottom;

            // TODO: 나중에 정리
            // textBox1.Text = "(정보 없음)";
            // if (string.IsNullOrEmpty(MergeArgs.ExtraInfoPath) == false)
            // {
            //     try
            //     {
            //         var fileContent = File.ReadAllText(MergeArgs.ExtraInfoPath);
            //         textBox1.Text = fileContent;
            //     }
            //     catch
            //     {
            //     }
            // }

            // initialize listViewWorksheets columns
            listViewWorksheets.Clear();
			listViewWorksheets.Columns.Add("워크시트 이름", 150);
			if (_pathViewModel.ComparisonMode == ComparisonMode.ThreeWay)
				listViewWorksheets.Columns.Add("충돌", 50);

			listViewWorksheets.Columns.Add("Mine (Dest/Curr)", 60);
			if (_pathViewModel.ComparisonMode == ComparisonMode.ThreeWay)
				listViewWorksheets.Columns.Add("Theirs (Src/Others)", 60);

			// double-buffer
			dataGridView1.GetType()?.
				GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.
				SetValue(dataGridView1, true, null);
			splitContainerBottom.GetType()?.
				GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)?.
				SetValue(splitContainerBottom, true, null);

			// progress box set-up
			FakeBackgroundWorker.OnUpdateProgress = onUpdateProgress;

			// diff3 진행
			var diff3Data = new XlsxDiff3Core();
			var compareResults = diff3Data.Run(_pathViewModel);

			FakeBackgroundWorker.OnUpdateProgress("xlsx 파일 비교", "비교 완료. 기본 설정 진행 중..");
			_xlsxMergeDecision = new XlsxMergeDecision(diff3Data, compareResults);

			foreach (var eachSheetResult in compareResults)
			{
				String sheetName = eachSheetResult.WorksheetName;

				var lvi = listViewWorksheets.Items.Add(sheetName);
				lvi.UseItemStyleForSubItems = false;

				if (_pathViewModel.ComparisonMode == ComparisonMode.ThreeWay)
				{
					if (eachSheetResult.HunkList.Find(r => r.hunkStatus == Diff3HunkStatus.Conflict) != null)
						lvi.SubItems.Add("발생").BackColor = Color.Gold;
					else
						lvi.SubItems.Add("");
				}

                var modificationMine = eachSheetResult.GetModificationSummary(DocOrigin.Mine);
                var subItem = lvi.SubItems.Add(modificationMine.Name);
                subItem.BackColor = modificationMine.Color;

                if (_pathViewModel.ComparisonMode == ComparisonMode.ThreeWay)
                {
                    var modificationTheirs = eachSheetResult.GetModificationSummary(DocOrigin.Theirs);
                    subItem = lvi.SubItems.Add(modificationTheirs.Name);
                    subItem.BackColor = modificationTheirs.Color;
                }
            }

			FakeBackgroundWorker.OnUpdateProgress(null);
		}

		private void listViewWorksheets_SelectedIndexChanged(object sender, EventArgs e)
		{
			ChangeFocusedHunk(0);
			UpdatePreviewWindow();
			HighlightFocusedHunk();
		}

		private void checkBoxShowFirstRowContentsOnTop_CheckedChanged(object sender, EventArgs e)
		{
			UpdateDataGridViewColumnName();
		}

		private void buttonSaveMergeResult_Click(object sender, EventArgs e)
		{
			int unResolvedConflictCount = 0;
			foreach (var dc in _xlsxMergeDecision.SheetMergeDecisionList)
			{
				if (dc.MergeModeDecision != WorksheetMergeMode.Merge)
					continue;
				foreach (var hunkDc in dc.HunkMergeDecisionList)
					if (hunkDc.DocMergeOrder == null)
						unResolvedConflictCount++;
			}

			if (unResolvedConflictCount > 0)
			{
				if (MessageBox.Show($"충돌 상태로 둔 변경 지점이 {unResolvedConflictCount}곳 있습니다. 이 상태로 결과를 저장할까요?", "충돌 상태",
						MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
					return;
			}

            if (!validateResultPath(out var mergedFilePath))
                return;

            _pathViewModel.ResultPath = mergedFilePath;

            XlsxMergeCommand cmd = new XlsxMergeCommand();
			cmd.Init(_xlsxMergeDecision);

			using (var runner = new XlsxMergeCommandRunner())
			{
				runner.Run(cmd, _pathViewModel.BasePath, _pathViewModel.MinePath, _pathViewModel.TheirsPath, mergedFilePath);
			}

			if (string.IsNullOrEmpty(_pathViewModel.ResultPath))
			{
                var dialogResult = MessageBox.Show("결과를 저장했습니다. 파일을 지금 열어보시겠습니까?", "완료", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
					ProcessStartInfo psi = new ProcessStartInfo()
					{
						FileName = mergedFilePath,
						UseShellExecute = true
					};
					Process.Start(psi);
				}
			}
			else
			{
				MergeSuccessful = true;
				Close();
			}
		}

        private bool validateResultPath(out string resultPath)
        {
            resultPath = _pathViewModel.ResultPath;
            if (string.IsNullOrEmpty(resultPath))
            {
                if (saveFileDialog1.ShowDialog() != DialogResult.OK)
                    return false;

                resultPath = saveFileDialog1.FileName;
            }

            var alertFileMap = new Dictionary<string, string>()
                {
                    { "Base", _pathViewModel.BasePath },
                    { "Mine", _pathViewModel.MinePath },
                    { "Theirs", _pathViewModel.TheirsPath},
                };
            foreach (var alertItems in alertFileMap)
            {
                if (string.IsNullOrEmpty(alertItems.Value))
                    continue;

                if (HelperFunctions.IsTwoPathEqual(resultPath, alertItems.Value) == false)
                    continue;

                if (MessageBox.Show(
                    $"결과를 저장할 경로와 {alertItems.Key} 파일 경로가 동일합니다.{Environment.NewLine}" +
                    $"다음에 다시 머지를 시도할 경우 변경된 파일이 쓰이게 됩니다.{Environment.NewLine}" +
                    "계속 진행하시겠습니까?",
                    "파일 경고",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                    return false;
            }

            return true;
        }

        private void linkLabelChangeWorksheetMergeMode_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			var menuItems = contextMenuStrip1.Items.Cast<ToolStripMenuItem>().ToArray();
			foreach (ToolStripMenuItem menuItem in menuItems)
			{
				menuItem.Click -= MergeModeClick;
				menuItem.Dispose();
			}

			contextMenuStrip1.Items.Clear();
			foreach (var eachDecision in sheetDecision.MergeModeCandidates)
			{
				var newMenuItem = new ToolStripMenuItem();
				newMenuItem.Text = eachDecision.GetDisplayText();
                newMenuItem.ShortcutKeyDisplayString = eachDecision.ToString(); 
				newMenuItem.Tag = new List<WorksheetMergeMode>() { eachDecision };
				newMenuItem.Click += MergeModeClick;
				contextMenuStrip1.Items.Add(newMenuItem);
			};
			contextMenuStrip1.Show(linkLabelChangeWorksheetMergeMode, new Point(0, 0));
		}

		private void MergeModeClick(object sender, EventArgs eventArgs)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			var selectedMenuItem = (sender as ToolStripMenuItem);
			if (selectedMenuItem == null)
				return;

			var newSheetDecisionMode = (selectedMenuItem.Tag as List<WorksheetMergeMode>)[0];
			if (sheetDecision.MergeModeDecision == newSheetDecisionMode)
				return;

			// change sheet decision
			sheetDecision.MergeModeDecision = newSheetDecisionMode;
			ChangeFocusedHunk(0);
			UpdatePreviewWindow();
			HighlightFocusedHunk();
		}

		private void checkBoxHideRemovedLines_CheckedChanged(object sender, EventArgs e)
		{
			UpdatePreviewWindow();
			HighlightFocusedHunk();
		}

		private void checkBoxHideEqualLines_CheckedChanged(object sender, EventArgs e)
		{
			UpdatePreviewWindow();
			HighlightFocusedHunk();
		}

		private void linkLabelChangeMergeOrder_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			if (_focusedHunkIdx < 0 || _focusedHunkIdx >= sheetDecision.HunkMergeDecisionList.Count)
				return;

			var menuItems = contextMenuStrip1.Items.Cast<ToolStripMenuItem>().ToArray();
			foreach (ToolStripMenuItem menuItem in menuItems)
			{
				menuItem.Click -= MergeModeClick;
				menuItem.Dispose();
			}

			contextMenuStrip1.Items.Clear();

			ToolStripMenuItem subMenuItem = null;

			foreach (var candidate in sheetDecision.HunkMergeDecisionList[_focusedHunkIdx].DocMergeOrderCandidates)
			{
				var newMenuItem = new ToolStripMenuItem();
				newMenuItem.Text = candidate.GetDisplayText();

				if (candidate == null)
					newMenuItem.ShortcutKeyDisplayString = "Conflict";
				else if (candidate.Count == 0)
					newMenuItem.ShortcutKeyDisplayString = "Delete";
				newMenuItem.Tag = candidate;
				newMenuItem.Click += HunkDecisionClick;

				if (candidate != null && candidate.Count > 1)
				{
					if (subMenuItem == null)
					{
						subMenuItem = new ToolStripMenuItem();
						subMenuItem.Text = "변경 사항 조합하기";
						contextMenuStrip1.Items.Add(subMenuItem);
					}
					subMenuItem.DropDownItems.Add(newMenuItem);
				}
				else
					contextMenuStrip1.Items.Add(newMenuItem);
			};
			contextMenuStrip1.Show(linkLabelChangeMergeOrder, new Point(0, 0));
		}


		private void HunkDecisionClick(object sender, EventArgs eventArgs)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			if (_focusedHunkIdx < 0 || _focusedHunkIdx >= sheetDecision.HunkMergeDecisionList.Count)
				return;


			var selectedMenuItem = (sender as ToolStripMenuItem);
			if (selectedMenuItem == null)
				return;

			var menuResultMergeOrder = selectedMenuItem.Tag as List<DocOrigin>;

			sheetDecision.HunkMergeDecisionList[_focusedHunkIdx].DocMergeOrder = menuResultMergeOrder;

			ChangeFocusedHunk(_focusedHunkIdx);
			UpdatePreviewWindow();
			HighlightFocusedHunk();
		}

		private void dataGridView1_SelectionChanged(object sender, EventArgs e)
		{
			if (_dataGridViewCellUpdatingInProgress == false)
				HighlightFocusedHunk();
		}

		private XlsxMergeDecision.SheetMergeDecision getCurrentSheetDecision()
		{
			int selectedWorksheetIndex = listViewWorksheets.SelectedIndices.Count > 0 ? listViewWorksheets.SelectedIndices[0] : -1;
			if (selectedWorksheetIndex < 0)
				return null;

			return _xlsxMergeDecision.SheetMergeDecisionList[selectedWorksheetIndex];
		}


		private void UpdatePreviewWindow()
		{
			label2.Text = "선택한 워크시트 : ";
			labelCurrentWorksheetMergeMode.Text = "---";
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			_dataGridViewCellUpdatingInProgress = true;

			var sheetResult = sheetDecision.SheetDiffResult;

			label2.Text = label2.Text + sheetResult.WorksheetName;
			labelCurrentWorksheetMergeMode.Text = sheetDecision.MergeModeDecision.GetDisplayText();
			labelTotalDiffHunks.Text = $"{sheetDecision.HunkMergeDecisionList.Count}";

			if (sheetDecision.MergeModeDecision == WorksheetMergeMode.Merge)
			{
				panelMergeHunksOn.Visible = true;
				panelMergeHunksOff.Visible = false;
			}
			else
			{ 
				panelMergeHunksOn.Visible = false;
				panelMergeHunksOff.Visible = true;
			}



			var worksheetBase = _xlsxMergeDecision.DiffResult.GetParsedWorksheetData(sheetResult.WorksheetName)[DocOrigin.Base];
			var previewData = MergeResultPreviewData.GeneratePreviewData(getCurrentSheetDecision(), worksheetBase == null ? 0 : worksheetBase.RowCount, checkBoxHideRemovedLines.Checked, checkBoxHideEqualLines.Checked);
			previewDataCache[sheetResult.WorksheetName] = previewData;
			MergeResultPreviewer.RefreshDataGridViewContents(_xlsxMergeDecision, sheetDecision, dataGridView1, previewData);
			UpdateDataGridViewColumnName();

			_dataGridViewCellUpdatingInProgress = false;
		}

		private void ChangeFocusedHunk(int hunkIdx)
		{
			_focusedHunkIdx = hunkIdx;
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				_focusedHunkIdx = -1;

			labelCurrentDiffHunkIdx.Text = $"{_focusedHunkIdx + 1}";
			var displayString = "---";
			if (sheetDecision != null && _focusedHunkIdx >= 0 && _focusedHunkIdx < sheetDecision.HunkMergeDecisionList.Count)
				displayString = sheetDecision.HunkMergeDecisionList[_focusedHunkIdx].DocMergeOrder.GetDisplayText();
			labelCurrentMergeOrder.Text = displayString;
		}

		private void HighlightFocusedHunk()
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			_dataGridViewCellUpdatingInProgress = true;

			dataGridView1.ClearSelection();

			var previewData = previewDataCache[sheetDecision.WorksheetName];
			if (_focusedHunkIdx != -1 && previewData.HunkStartsPosList.Count > 0)
			{
				for (int rowIndex = previewData.HunkStartsPosList[_focusedHunkIdx];
					rowIndex < previewData.HunkEndsPosList[_focusedHunkIdx];
					rowIndex++)
				{
					dataGridView1.Rows[rowIndex].Cells[0].Selected = true;
				}
				dataGridView1.FirstDisplayedScrollingRowIndex = previewData.HunkStartsPosList[_focusedHunkIdx];
			}
			_dataGridViewCellUpdatingInProgress = false;

		}

		private void UpdateDataGridViewColumnName()
		{
			if (dataGridView1.Rows.Count < 1)
				return;

			foreach (DataGridViewColumn eachCol in dataGridView1.Columns)
			{
				if (eachCol.Name.StartsWith("C"))
				{
					int columNumber = int.Parse(eachCol.Name.Substring(1));
					string headerText = HelperFunctions.GetExcelColumnName(columNumber);
					if (checkBoxShowFirstRowContentsOnTop.Checked)
					{
						string secondaryText = "";
						var cellValue = dataGridView1.Rows[0].Cells[eachCol.Name].Value;
						if (cellValue != null)
							secondaryText = cellValue.ToString();
						headerText = headerText + Environment.NewLine + secondaryText;
					}
					eachCol.HeaderText = headerText;
				}
			}

		}

		private void buttonNavPrev_Click(object sender, EventArgs e)
		{
			NavigateHunk(-1);
		}

		private void buttonNavNext_Click(object sender, EventArgs e)
		{
			NavigateHunk(1);
		}

		private void NavigateHunk(int direction)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			if (previewDataCache.ContainsKey(sheetDecision.WorksheetName) == false)
				return;

			if (_focusedHunkIdx == -1)
				return;

			var startPosList = previewDataCache[sheetDecision.WorksheetName].HunkStartsPosList;

			var newHunkIdx = (_focusedHunkIdx + direction + startPosList.Count) % startPosList.Count;
			_focusedHunkIdx = newHunkIdx;
			ChangeFocusedHunk(_focusedHunkIdx);
			HighlightFocusedHunk();
		}
		
		private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
		{
			var sheetDecision = getCurrentSheetDecision();
			if (sheetDecision == null)
				return;

			var sheetResult = sheetDecision.SheetDiffResult;
			var hunkIdx = previewDataCache[sheetResult.WorksheetName].GetHunkIdxByRowNumber(e.RowIndex);
			if (hunkIdx >= 0)
			{
				ChangeFocusedHunk(hunkIdx);
				UpdatePreviewWindow();
				HighlightFocusedHunk();
			}
		}

		private Label _notifyLabel = null;
		private void onUpdateProgress(string title = null, params string[] message)
		{
			if (title == null)
			{
				if (_notifyLabel != null)
					this.Controls.Remove(_notifyLabel);
				this.Enabled = true;
				return;
			}

			this.Enabled = false;
			if (_notifyLabel == null)
			{
				_notifyLabel = new Label();
				_notifyLabel.AutoSize = false;
				_notifyLabel.TextAlign = ContentAlignment.MiddleCenter;
				_notifyLabel.BackColor = Color.PaleGreen;
				_notifyLabel.Font = new Font("Courier New", 14);
			}

			_notifyLabel.Text = "";
			_notifyLabel.Location = new Point(this.Width / 4, this.Height / 4);
			_notifyLabel.Size = new Size(this.Width / 2, this.Height / 2);
			this.Controls.Add(_notifyLabel);
			_notifyLabel.BringToFront();


			_notifyLabel.Text = title + Environment.NewLine + Environment.NewLine + String.Join(Environment.NewLine, message);
			Application.DoEvents();
		}

        private void buttonCopyTableContents_Click(object sender, EventArgs e)
        {
            _dataGridViewCellUpdatingInProgress = true;
            dataGridView1.SelectAll();
            DataObject dataObj = dataGridView1.GetClipboardContent();
            Clipboard.SetDataObject(dataObj, true);
            dataGridView1.ClearSelection();
            _dataGridViewCellUpdatingInProgress = false;

            HighlightFocusedHunk();
            MessageBox.Show("복사되었습니다.", "테이블 내용 복사", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
