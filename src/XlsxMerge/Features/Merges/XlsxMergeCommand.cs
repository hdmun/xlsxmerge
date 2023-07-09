using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using XlsxMerge.Features.Diffs;

namespace XlsxMerge.Features.Merges
{
    // XlsxMergeDecision으로부터, 실제 조립 명령어를 만들어 냅니다.
    class XlsxMergeCommand
    {
        public DocOrigin _docOriginMergeInto = DocOrigin.Mine;

        public void Init(XlsxMergeDecision mergeDecision, DocOrigin mergeInto = DocOrigin.Mine)
        {
            _docOriginMergeInto = mergeInto;

            CommandList.Clear();
            // MINE 을 기본으로 합니다.
            foreach (var sheetDecision in mergeDecision.SheetMergeDecisionList)
            {
                switch (sheetDecision.MergeModeDecision)
                {
                    case WorksheetMergeMode.Unchanged:
                        // do nothing
                        break;
                    case WorksheetMergeMode.Delete:
                        if (sheetDecision.SheetDiffResult.HasDocOrigin(_docOriginMergeInto))
                            CommandList.Add(XlsxMergeCommandItem.DeleteSheet(_docOriginMergeInto, sheetDecision.WorksheetName));
                        break;
                    case WorksheetMergeMode.UseBase:
                        if (_docOriginMergeInto != DocOrigin.Base)
                        {
                            if (sheetDecision.SheetDiffResult.HasDocOrigin(_docOriginMergeInto))
                                CommandList.Add(XlsxMergeCommandItem.DeleteSheet(_docOriginMergeInto, sheetDecision.WorksheetName));
                            CommandList.Add(XlsxMergeCommandItem.CopySheet(DocOrigin.Base, _docOriginMergeInto, sheetDecision.WorksheetName));
                        }
                        break;
                    case WorksheetMergeMode.UseMine:
                        if (_docOriginMergeInto != DocOrigin.Mine)
                        {
                            if (sheetDecision.SheetDiffResult.HasDocOrigin(_docOriginMergeInto))
                                CommandList.Add(XlsxMergeCommandItem.DeleteSheet(_docOriginMergeInto, sheetDecision.WorksheetName));
                            CommandList.Add(XlsxMergeCommandItem.CopySheet(DocOrigin.Mine, _docOriginMergeInto, sheetDecision.WorksheetName));
                        }
                        break;
                    case WorksheetMergeMode.UseTheirs:
                        if (_docOriginMergeInto != DocOrigin.Theirs)
                        {
                            if (sheetDecision.SheetDiffResult.HasDocOrigin(_docOriginMergeInto))
                                CommandList.Add(XlsxMergeCommandItem.DeleteSheet(_docOriginMergeInto, sheetDecision.WorksheetName));
                            CommandList.Add(XlsxMergeCommandItem.CopySheet(DocOrigin.Theirs, _docOriginMergeInto, sheetDecision.WorksheetName));
                        }

                        break;
                    case WorksheetMergeMode.Merge:
                        ProcessMergeDecision(sheetDecision);
                        break;
                }
            }
        }

        private int AddRowCopyCommand(DocOrigin docOriginSource, string worksheetName, int rowNumberInsertAt, HunkMergeDecision hunk)
        {
            int rowNumber = hunk.BaseHunkInfo.GetRowNumber(docOriginSource);
            int rowCount = hunk.BaseHunkInfo.GetRowCount(docOriginSource);
            if (docOriginSource != _docOriginMergeInto && rowCount > 0)
                CommandList.Add(XlsxMergeCommandItem.CopyRow(docOriginSource, _docOriginMergeInto, worksheetName, rowCount, rowCount, rowNumberInsertAt));

            return rowNumberInsertAt + rowCount;
        }

        public void ProcessMergeDecision(SheetMergeDecision sheetMergeDecision)
        {
            // 맨 끝 변경사항부터 첫 변경사항으로 적용합니다.
            // 이 방법을 사용할 경우 행 추가/삭제로 인한 줄 번호 재계산을 할 필요가 없어집니다.

            var worksheetName = sheetMergeDecision.WorksheetName;
            var hunkMergeDecisionList = sheetMergeDecision.HunkMergeDecisionList
                .OrderBy(r => -r.BaseHunkInfo.GetRowNumber(_docOriginMergeInto));
            foreach (var hunk in hunkMergeDecisionList)
            {
                int currentRowNumber = hunk.BaseHunkInfo.GetRowNumber(_docOriginMergeInto);
                if (hunk.IsConflict)
                {
                    var commandText = XlsxMergeCommandItem.InsertText(_docOriginMergeInto, worksheetName, "[XlsxMerge충돌] 충돌 지점 시작 >>>", currentRowNumber);
                    CommandList.Add(commandText);
                    currentRowNumber++;

                    commandText = XlsxMergeCommandItem.InsertText(_docOriginMergeInto, worksheetName, ">Base<", currentRowNumber);
                    CommandList.Add(commandText);
                    currentRowNumber++;
                    currentRowNumber = AddRowCopyCommand(DocOrigin.Base, worksheetName, currentRowNumber, hunk);

                    commandText = XlsxMergeCommandItem.InsertText(_docOriginMergeInto, worksheetName, ">Mine(Destination)<", currentRowNumber);
                    CommandList.Add(commandText);
                    currentRowNumber++;
                    currentRowNumber = AddRowCopyCommand(DocOrigin.Mine, worksheetName, currentRowNumber, hunk);

                    commandText = XlsxMergeCommandItem.InsertText(_docOriginMergeInto, worksheetName, ">Theirs(Source)<", currentRowNumber);
                    CommandList.Add(commandText);
                    currentRowNumber++;
                    currentRowNumber = AddRowCopyCommand(DocOrigin.Theirs, worksheetName, currentRowNumber, hunk);

                    commandText = XlsxMergeCommandItem.InsertText(_docOriginMergeInto, worksheetName, "<<< 충돌 지점 끝", currentRowNumber);
                    CommandList.Add(commandText);
                    currentRowNumber++;
                }
                else
                {
                    foreach (var docOrigin in hunk.DocMergeOrder)
                        currentRowNumber = AddRowCopyCommand(docOrigin, worksheetName, currentRowNumber, hunk);

                    var rowCount = hunk.BaseHunkInfo.GetRowCount(_docOriginMergeInto);
                    if (hunk.DocMergeOrder.Contains(_docOriginMergeInto) == false && rowCount > 0)
                    {
                        var commandItem = XlsxMergeCommandItem.DeleteRow(_docOriginMergeInto, worksheetName, currentRowNumber, rowCount);
                        CommandList.Add(commandItem);
                    }
                }
            }
        }

        public List<XlsxMergeCommandItem> CommandList = new List<XlsxMergeCommandItem>();

        public string Dump()
        {
            StringBuilder sb = new StringBuilder();

            foreach (var eachCommand in CommandList)
            {
                sb.Append(eachCommand.Cmd);
                sb.Append("\t" + (eachCommand.sourceOrigin.HasValue ? eachCommand.sourceOrigin.Value.ToString() : ""));
                sb.Append("\t" + (eachCommand.destOrigin.HasValue ? eachCommand.destOrigin.Value.ToString() : ""));
                sb.Append("\t" + eachCommand.param1);
                sb.Append("\t" + eachCommand.param2);
                sb.Append("\t" + (eachCommand.intParam1 == -1 ? "" : eachCommand.intParam1.ToString()));
                sb.Append("\t" + (eachCommand.intParam2 == -1 ? "" : eachCommand.intParam2.ToString()));
                sb.Append("\t" + (eachCommand.intParam3 == -1 ? "" : eachCommand.intParam3.ToString()));
                sb.AppendLine();
            }

            return sb.ToString();
        }
    }

    public class XlsxMergeCommandItem
    {
        public string Cmd;
        public DocOrigin? sourceOrigin;
        public DocOrigin? destOrigin;
        public string param1 = "";
        public string param2 = "";
        public int intParam1 = -1;
        public int intParam2 = -1;
        public int intParam3 = -1;

        public static XlsxMergeCommandItem DeleteSheet(DocOrigin docOrigin, string worksheetName)
        {
            return new XlsxMergeCommandItem()
            {
                Cmd = "DELETE_SHEET",
                destOrigin = docOrigin,
                param1 = worksheetName
            };
        }

        public static XlsxMergeCommandItem CopySheet(DocOrigin docOriginFrom, DocOrigin docOriginTo, string worksheetNameFrom)
        {
            return new XlsxMergeCommandItem()
            {
                Cmd = "COPY_SHEET",
                sourceOrigin = docOriginFrom,
                destOrigin = docOriginTo,
                param1 = worksheetNameFrom,
            };
        }

        public static XlsxMergeCommandItem CopyRow(DocOrigin docOriginFrom, DocOrigin docOriginTo, string worksheetName, int rowNumberFrom, int rowCountFrom, int rowNumberInsertAt)
        {
            return new XlsxMergeCommandItem()
            {
                Cmd = "COPY_ROW",
                sourceOrigin = docOriginFrom,
                destOrigin = docOriginTo,
                param1 = worksheetName,
                intParam1 = rowNumberInsertAt,
                intParam2 = rowNumberFrom,
                intParam3 = rowCountFrom,
            };
        }

        public static XlsxMergeCommandItem InsertText(DocOrigin docOrigin, string worksheetName, string text, int rowNumberInsertAt)
        {
            return new XlsxMergeCommandItem()
            {
                Cmd = "INSERT_TEXT",
                destOrigin = docOrigin,
                param1 = worksheetName,
                param2 = text,
                intParam1 = rowNumberInsertAt
            };
        }

        public static XlsxMergeCommandItem DeleteRow(DocOrigin docOrigin, string worksheetName, int rowNumberDeleteAt, int rowCount)
        {
            return new XlsxMergeCommandItem()
            {
                Cmd = "DELETE_ROW",
                destOrigin = docOrigin,
                param1 = worksheetName,
                intParam1 = rowNumberDeleteAt,
                intParam2 = rowCount,
            };
        }
    }
}
