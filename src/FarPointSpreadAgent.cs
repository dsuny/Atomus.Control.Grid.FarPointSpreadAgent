using Atomus.Diagnostics;
using FarPoint.Win.Spread;
using FarPoint.Win.Spread.CellType;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Atomus.Control.Grid
{
    public class FarPointSpreadAgent : IDataGridAgent
    {
        readonly string[] MenuListText;
        System.Windows.Forms.Control IDataGridAgent.GridControl { get; set; }

        Dictionary<TextBox, FilterAttribute> FilterControl { get; set; }
        //int headerRowCount;
        //int headerHeight;

        EditAble IDataGridAgent.EditAble
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DefaultStyle.Locked) ? EditAble.False : EditAble.True;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DefaultStyle.Locked = (value == EditAble.True) ? false : true;
            }
        }
        AddRows IDataGridAgent.AddRows
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DataAllowAddNew) ? AddRows.True : AddRows.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DataAllowAddNew = (value == AddRows.True) ? true : false;
            }
        }
        //미구현
        DeleteRows IDataGridAgent.DeleteRows { get; set; }
        ResizeRows IDataGridAgent.ResizeRows
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.Rows.Default.Resizable) ? ResizeRows.True : ResizeRows.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.Rows.Default.Resizable = (value == ResizeRows.True) ? true : false;
            }
        }
        AutoSizeColumns IDataGridAgent.AutoSizeColumns
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DataAutoSizeColumns) ? AutoSizeColumns.True : AutoSizeColumns.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.DataAutoSizeColumns = (value == AutoSizeColumns.True) ? true : false;
            }
        }
        //미구현
        AutoSizeRows IDataGridAgent.AutoSizeRows { get; set; }
        ColumnsHeadersVisible IDataGridAgent.ColumnsHeadersVisible
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.Visible) ? ColumnsHeadersVisible.True : ColumnsHeadersVisible.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.Visible = (value == ColumnsHeadersVisible.True) ? true : false;
            }
        }
        EnableMenu IDataGridAgent.EnableMenu {
            get
            {

                if (((IDataGridAgent)this).GridControl.ContextMenu != null)
                    return EnableMenu.False;
                else
                    return EnableMenu.False;
            }
            set
            {
                if (value == EnableMenu.True && ((IDataGridAgent)this).GridControl.ContextMenu == null)
                    SetContextMenuStrip((FpSpread)((IDataGridAgent)this).GridControl, this.MenuListText);

            }
        }
        //수정 해야 될 수도 있음
        MultiSelect IDataGridAgent.MultiSelect
        {
            get
            {
                return ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode == OperationMode.MultiSelect ? MultiSelect.True : MultiSelect.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode = (value == MultiSelect.True) ? OperationMode.MultiSelect : OperationMode.SingleSelect;
            }
        }
        Alignment IDataGridAgent.HeadAlignment
        {
            get
            {
                switch (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment.ToString() + ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment.ToString())
                {
                    case "BottomCenter":
                        return Alignment.BottomCenter;
                    case "BottomLeft":
                        return Alignment.BottomLeft;
                    case "BottomRight":
                        return Alignment.BottomRight;
                    case "CenterCenter":
                        return Alignment.MiddleCenter;
                    case "CenterLeft":
                        return Alignment.MiddleLeft;
                    case "CenterRight":
                        return Alignment.MiddleRight;
                    case "GeneralGeneral":
                        return Alignment.NotSet;
                    case "TopCenter":
                        return Alignment.TopCenter;
                    case "TopLeft":
                        return Alignment.TopLeft;
                    case "TopRight":
                        return Alignment.TopRight;
                    default:
                        return Alignment.NotSet;
                }
            }
            set
            {
                switch (value)
                {
                    case Alignment.BottomCenter:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Bottom;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Center;
                        break;
                    case Alignment.BottomLeft:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Bottom;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Left;
                        break;
                    case Alignment.BottomRight:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Bottom;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Right;
                        break;
                    case Alignment.MiddleCenter:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Center;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Center;
                        break;
                    case Alignment.MiddleLeft:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Center;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Left;
                        break;
                    case Alignment.MiddleRight:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Center;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Right;
                        break;
                    case Alignment.NotSet:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
                        break;
                    case Alignment.TopCenter:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Top;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Center;
                        break;
                    case Alignment.TopLeft:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Top;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Left;
                        break;
                    case Alignment.TopRight:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.Top;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.Right;
                        break;
                    default:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.VerticalAlignment = CellVerticalAlignment.General;
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.DefaultStyle.HorizontalAlignment = CellHorizontalAlignment.General;
                        break;
                }
            }
        }
        int IDataGridAgent.HeaderHeight
        {
            get
            {
                return (int)((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.Rows.Default.Height;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.Rows.Default.Height = value;
            }
        }
        int IDataGridAgent.HeaderRowCount {
            get
            {
                return ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.RowCount;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.ColumnHeader.RowCount = value;
            }
        }

        int IDataGridAgent.RowHeight
        {
            get
            {
                return (int)((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.Rows.Default.Height;
            }
            set
            {
                if (value >= 0)//'행 높이
                    ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.Rows.Default.Height = value;
            }
        }
        RowHeadersVisible IDataGridAgent.RowHeadersVisible
        {
            get
            {
                return (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.RowHeaderVisible) ? RowHeadersVisible.True : RowHeadersVisible.False;
            }
            set
            {
                ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.RowHeaderVisible = (value == RowHeadersVisible.True) ? true : false;
            }
        }
        Selection IDataGridAgent.Selection
        {
            get
            {
                switch (((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode)
                {
                    case OperationMode.Normal:
                        return Selection.CellSelect;
                    case OperationMode.RowMode:
                        return Selection.FullRowSelect;
                    case OperationMode.SingleSelect:
                        return Selection.RowHeaderSelect;
                    default:
                        throw new AtomusException("Selection Not Support.");
                }

                // (Selection)Enum.Parse(typeof(Selection), ((FpSpread)((IDataGridAgent)this).GridControl).SelectionMode.ToString());
            }
            set
            {
                switch (value)
                {
                    case Selection.CellSelect:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode = OperationMode.Normal;
                        break;
                    case Selection.FullRowSelect:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode = OperationMode.RowMode;
                        break;
                    case Selection.RowHeaderSelect:
                        ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.OperationMode = OperationMode.SingleSelect;
                        break;
                    default:
                        throw new AtomusException("Alignment Not Support.");
                }

                //((FpSpread)((IDataGridAgent)this).GridControl).SelectionMode = (DataGridViewSelectionMode)Enum.Parse(typeof(DataGridViewSelectionMode), value.ToString());
            }
        }

        public FarPointSpreadAgent()
        {
            this.FilterControl = new Dictionary<TextBox, FilterAttribute>();

            try
            {
                this.MenuListText = this.GetAttribute("MenuListText").Split(',').Translate();
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);
            }
        }

        void IDataGridAgent.Init(EditAble _EditAble, AddRows _AllowAddRows, DeleteRows _AllowDeleteRows, ResizeRows _AllowResizeRows, AutoSizeColumns _AutoSizeColumns
            , AutoSizeRows _AutoSizeRows, ColumnsHeadersVisible _ColumnsHeadersVisible, EnableMenu _EnableMenu, MultiSelect _MultiSelect, Alignment _HeadAlign, int _HeaderHeight, int _HeaderRowCount
            , int _RowHeight, RowHeadersVisible _RowHeadersVisible, Selection _SelectionMode)
        {
            FpSpread fpSpread;

            fpSpread = ((FpSpread)((IDataGridAgent)this).GridControl);

            fpSpread.ActiveSheet.Rows.Clear();
            fpSpread.ActiveSheet.Columns.Clear();

            //_DataGridView.DefaultCellStyle.WrapMode = DataGridViewTriState.False; //자동 줄바꿈 금지
            fpSpread.ActiveSheet.ColumnHeader.Rows.Default.Resizable = false;//컬럼헤더 높이 변경 금지

            ((IDataGridAgent)this).EditAble = _EditAble;
            ((IDataGridAgent)this).AddRows = _AllowAddRows;
            ((IDataGridAgent)this).DeleteRows = _AllowDeleteRows;
            ((IDataGridAgent)this).ResizeRows = _AllowResizeRows;
            ((IDataGridAgent)this).AutoSizeColumns = _AutoSizeColumns;
            ((IDataGridAgent)this).AutoSizeRows = _AutoSizeRows;
            ((IDataGridAgent)this).ColumnsHeadersVisible = _ColumnsHeadersVisible;
            ((IDataGridAgent)this).EnableMenu = _EnableMenu;
            ((IDataGridAgent)this).MultiSelect = _MultiSelect;
            ((IDataGridAgent)this).HeadAlignment = _HeadAlign;
            ((IDataGridAgent)this).RowHeight = _RowHeight;

            ((IDataGridAgent)this).HeaderRowCount = _HeaderRowCount;//'컬럼헤더 수
            ((IDataGridAgent)this).HeaderHeight = _HeaderHeight;

            ((IDataGridAgent)this).RowHeadersVisible = _RowHeadersVisible;
            ((IDataGridAgent)this).Selection = _SelectionMode;

            fpSpread.DoubleBuffered(true);

            fpSpread.DataColumnConfigure -= this.DataGridView_DataSourceChanged;
            fpSpread.DataColumnConfigure += this.DataGridView_DataSourceChanged;

            fpSpread.DataColumnConfigure -= this.DataGridView_DataBindingComplete;
            fpSpread.DataColumnConfigure += this.DataGridView_DataBindingComplete;

            this.SetSkin(fpSpread);

        }

        void IDataGridAgent.AddColumn(int width, ColumnVisible visible, EditAble editAble, Filter allowFilter, Merge allowMerge, Sort sortMode, object editControl, Alignment textAlign, string format, string name, params string[] caption)
        {
            Column _DataGridViewColumn;
            FpSpread _DataGridView;
            int columnIndex;

            _DataGridView = ((FpSpread)((IDataGridAgent)this).GridControl);

            _DataGridViewColumn = null;

            foreach (Column column in _DataGridView.ActiveSheet.Columns)
                if (column.DataField == name)
                {
                    _DataGridViewColumn = column;
                    break;
                }

            if (_DataGridViewColumn == null)//'교체할 컬럼이 없으면 생성
            {
                _DataGridView.ActiveSheet.ColumnCount += 1;
                _DataGridViewColumn = _DataGridView.ActiveSheet.Columns[_DataGridView.ActiveSheet.ColumnCount - 1];
                _DataGridViewColumn.DataField = name;//'Data 컬럼명
            }

            columnIndex = _DataGridViewColumn.Index;

            caption = caption.Translate();

            if (((IDataGridAgent)this).HeaderRowCount == 1)//'컬럼 헤더 수가 1개면
                _DataGridView.ActiveSheet.ColumnHeader.Cells[0, columnIndex].Text = caption[caption.Length - 1];
            else
            {
                for (int i = 0; i <= caption.Length - 1; i++)
                {
                    _DataGridView.ActiveSheet.ColumnHeader.Cells[i, columnIndex].Text = caption[i];
                    _DataGridView.ActiveSheet.ColumnHeader.Rows[i].MergePolicy = FarPoint.Win.Spread.Model.MergePolicy.Always;
                }
                _DataGridView.ActiveSheet.ColumnHeader.Columns[columnIndex].MergePolicy = FarPoint.Win.Spread.Model.MergePolicy.Always;
            }

            _DataGridViewColumn.Width = width;
            _DataGridViewColumn.Visible = (visible == ColumnVisible.True);
            _DataGridViewColumn.Locked = !(editAble == EditAble.True);

            _DataGridViewColumn.AllowAutoFilter = (allowFilter == Filter.True);
            _DataGridViewColumn.MergePolicy = (allowMerge == Merge.True) ? FarPoint.Win.Spread.Model.MergePolicy.Always : FarPoint.Win.Spread.Model.MergePolicy.None;
            
            _DataGridViewColumn.AllowAutoSort = (sortMode != Sort.NotSortable) ? true : false;
            switch (textAlign)
            {
                case Alignment.BottomCenter:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Bottom;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Center;
                    break;
                case Alignment.BottomLeft:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Bottom;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Left;
                    break;
                case Alignment.BottomRight:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Bottom;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Right;
                    break;
                case Alignment.MiddleCenter:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Center;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Center;
                    break;
                case Alignment.MiddleLeft:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Center;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Left;
                    break;
                case Alignment.MiddleRight:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Center;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Right;
                    break;
                case Alignment.TopCenter:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Top;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Center;
                    break;
                case Alignment.TopLeft:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Top;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Left;
                    break;
                case Alignment.TopRight:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.Top;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.Right;
                    break;
                case Alignment.NotSet:
                    _DataGridViewColumn.VerticalAlignment = CellVerticalAlignment.General;
                    _DataGridViewColumn.HorizontalAlignment = CellHorizontalAlignment.General;
                    break;
            }


            string tmp;
            int colspan;
            int rowspan;

            for (int col = 0; col < _DataGridView.ActiveSheet.ColumnHeader.Columns.Count; col++)
                for (int row = 0; row < _DataGridView.ActiveSheet.ColumnHeader.RowCount; row++)
                {
                    _DataGridView.ActiveSheet.ColumnHeader.Cells[row, col].ColumnSpan = 1;
                    _DataGridView.ActiveSheet.ColumnHeader.Cells[row, col].RowSpan = 1;
                }


            for (int col = 0; col < _DataGridView.ActiveSheet.ColumnHeader.Columns.Count; col++)
                for (int row = 0; row < _DataGridView.ActiveSheet.ColumnHeader.RowCount; row++)
                {
                    colspan = 1;
                    tmp = _DataGridView.ActiveSheet.ColumnHeader.Cells[row, col].Text;

                    for (int j1 = col + 1; j1 < _DataGridView.ActiveSheet.ColumnHeader.Columns.Count; j1++)
                        if (tmp == _DataGridView.ActiveSheet.ColumnHeader.Cells[row, j1].Text)
                        {
                            colspan += 1;
                        }
                        else
                            break;

                    rowspan = 1;

                    for (int j1 = row + 1; j1 < _DataGridView.ActiveSheet.ColumnHeader.Rows.Count; j1++)
                        if (tmp == _DataGridView.ActiveSheet.ColumnHeader.Cells[j1, col].Text)
                        {
                            rowspan += 1;
                        }
                        else
                            break;

                    _DataGridView.ActiveSheet.ColumnHeader.Cells[row, col].ColumnSpan = colspan;
                    _DataGridView.ActiveSheet.ColumnHeader.Cells[row, col].RowSpan = rowspan;

                    //col += (colspan - 1);
                    row += (rowspan - 1);
                }


            if (editControl != null) // '교체할 컬럼이 있으면 컬럼을 교체
            {
                _DataGridViewColumn.CellType = (ICellType)editControl;

                //if (editControl is CheckBoxCellType && _DataGridView.ActiveSheet.ColumnHeader.Columns[columnIndex].CellType == null)
                if (editControl is CheckBoxCellType && (caption == null || caption.Count() < 1 || caption[0].Trim().Length < 1))
                {
                    CheckBoxCellType checkBoxCellType = new CheckBoxCellType();

                    _DataGridView.ActiveSheet.ColumnHeader.Columns[columnIndex].CellType = checkBoxCellType;

                    for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                    {
                        _DataGridView.ActiveSheet.ColumnHeader.Cells[i, columnIndex].Value = 0;
                        _DataGridView.ActiveSheet.ColumnHeader.Cells[i, columnIndex].Locked = !(editAble == EditAble.True);
                    }

                    if (this.GetAttributeBool("ColumnHeaderCheckBoxClickEventEnable"))
                    {
                        _DataGridView.CellClick -= FpSpread_CellClick;
                        _DataGridView.CellClick += FpSpread_CellClick;
                    }
                }
            }
        }

        private void FpSpread_CellClick(object sender, CellClickEventArgs e)
        {
            FpSpread _DataGridView;
            int intTmp;
            bool boolTmp;

            _DataGridView = (FpSpread)sender;

            if (e.ColumnHeader && _DataGridView.ActiveSheet.ColumnHeader.Rows.Count > 0 && _DataGridView.ActiveSheet.Columns.Count > 0)
            {
                if (_DataGridView.ActiveSheet.ColumnHeader.Columns[e.Column].CellType != null
                    && _DataGridView.ActiveSheet.ColumnHeader.Columns[e.Column].CellType is CheckBoxCellType)
                {
                    if (_DataGridView.ActiveSheet.Columns[e.Column].CellType != null && _DataGridView.ActiveSheet.Columns[e.Column].CellType is CheckBoxCellType)
                    {
                        if (_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value is int)
                            if ((int)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value == 0)
                            {
                                for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                    _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = 1;

                                for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                {
                                    _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 1;
                                }
                            }
                            else
                            {
                                for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                    _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = 0;

                                for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                    _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 0;
                            }

                        if (_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value is bool)
                            if ((bool)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value)
                            {
                                for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                    _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = true;

                                for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                {
                                    _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 0;
                                }
                            }
                            else
                            {
                                for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                    _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = false;

                                for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                    _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 1;
                            }

                        if (_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value is string)
                        {
                            if (int.TryParse((string)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value, out intTmp))
                            {
                                if ((string)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value == "0")
                                {
                                    for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                        _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = 1;

                                    for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                    {
                                        _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 1;
                                    }
                                }
                                else
                                {
                                    for (int i = 0; i < _DataGridView.ActiveSheet.ColumnHeader.RowCount; i++)
                                        _DataGridView.ActiveSheet.ColumnHeader.Cells[i, e.Column].Value = 0;

                                    for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                                        _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 0;
                                }
                            }

                            //if (bool.TryParse((string)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value, out boolTmp))
                            //{
                            //    if (!boolTmp)
                            //    {
                            //        _DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value = char.IsUpper(((string)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value).Substring(0, 1).ToChar()) ? "True" : "true";
                            //        for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                            //        {
                            //            _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 1;
                            //        }
                            //    }
                            //    else
                            //    {
                            //        _DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value = char.IsUpper(((string)_DataGridView.ActiveSheet.ColumnHeader.Cells[0, e.Column].Value).Substring(0, 1).ToChar()) ? "False" : "false";
                            //        for (int row = 0; row < _DataGridView.ActiveSheet.Rows.Count; row++)
                            //            _DataGridView.ActiveSheet.Cells[row, e.Column].Value = 0;
                            //    }
                            //}
                        }
                    }
                }
            }
        }

        void IDataGridAgent.Clear()
        {
            FpSpread fpSpread;

            fpSpread = ((FpSpread)((IDataGridAgent)this).GridControl);

            fpSpread.ActiveSheet.Columns.Clear();
            fpSpread.ActiveSheet.DataSource = null;
        }

        void IDataGridAgent.RemoveColumn(string name)
        {
            FpSpread _DataGridView;
            Column _DataGridViewColumn;

            _DataGridView = ((FpSpread)((IDataGridAgent)this).GridControl);

            _DataGridViewColumn = null;

            foreach (Column column in _DataGridView.ActiveSheet.Columns)
            {
                if (column.DataField == name)
                {
                    _DataGridViewColumn = column;
                    break;
                }
            }

            _DataGridView.ActiveSheet.Columns.Remove(_DataGridViewColumn.Index, 1);
        }

        void IDataGridAgent.RemoveColumn(int index)
        {
            FpSpread fpSpread;

            fpSpread = ((FpSpread)((IDataGridAgent)this).GridControl);

            fpSpread.ActiveSheet.Columns.Remove(index, 1);
        }

        private void SetContextMenuStrip(FpSpread fpSpread, string[] _MenuListText)
        {
            ContextMenuStrip _ContextMenuStrip;
            ToolStripMenuItem _ToolStripMenuItem;

            _ContextMenuStrip = new ContextMenuStrip();

            _ContextMenuStrip.Opening += ContextMenuStrip_Opening;

            _ToolStripMenuItem = null;

            for (int i = 0; i <= this.MenuListText.Length - 1; i++)
            {
                if (i == 2)
                    continue;

                if (this.MenuListText[i] != "")
                {
                    _ToolStripMenuItem = new ToolStripMenuItem(this.MenuListText[i], null, ToolStripMenuItem_Click);
                    _ContextMenuStrip.Items.Add(_ToolStripMenuItem);
                }
                else
                    _ContextMenuStrip.Items.Add(new ToolStripSeparator());

                switch (i)
                {
                    case 5://행추가
                        if (!fpSpread.ActiveSheet.DataAllowAddNew)
                            _ToolStripMenuItem.Enabled = false;
                        break;

                    case 6://행삭제
                        if (((IDataGridAgent)this).DeleteRows != DeleteRows.True )
                            _ToolStripMenuItem.Enabled = false;
                        break;

                    case 7://행복사
                        if (!fpSpread.ActiveSheet.DataAllowAddNew)
                            _ToolStripMenuItem.Enabled = false;
                        break;
                }
            }

            fpSpread.ContextMenuStrip = _ContextMenuStrip;
        }

        private void ContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            ContextMenuStrip contextMenuStrip;

            FpSpread dataGridView;
            decimal sum;
            decimal count;
            decimal avg;

            sum = 0;
            count = 0;
            avg = 0;

            try
            {
                dataGridView = ((FpSpread)((IDataGridAgent)this).GridControl);
                contextMenuStrip = (ContextMenuStrip)sender;

                foreach (FarPoint.Win.Spread.Model.CellRange dataGridViewCell in dataGridView.ActiveSheet.GetSelections())
                {
                    try
                    {
                        for (int i = dataGridViewCell.Column; i < dataGridViewCell.Column + dataGridViewCell.ColumnCount; i++)
                        {
                            for (int j = dataGridViewCell.Row; j < dataGridViewCell.Row + dataGridViewCell.RowCount; j++)
                            {
                                if (!dataGridView.ActiveSheet.Columns[i].Visible)
                                    continue;

                                sum += dataGridView.ActiveSheet.Cells[j, i].Value.ToString().ToDecimal();
                                count += 1;
                            }
                        }

                    }
                    catch (Exception exception)
                    {
                        DiagnosticsTool.MyTrace(exception);
                    }
                }

                avg = sum / count;

                contextMenuStrip.Items[8].Text = string.Format("Sum : {0}", sum);
                contextMenuStrip.Items[9].Text = string.Format("Avg : {0}", avg);

                contextMenuStrip.Items[8].Tag = sum;
                contextMenuStrip.Items[9].Tag = avg;
            }
            catch (Exception exception)
            {
                DiagnosticsTool.MyTrace(exception);

                contextMenuStrip = (ContextMenuStrip)sender;
                contextMenuStrip.Items[8].Text = string.Format("Sum : {0}", 0);
                contextMenuStrip.Items[9].Text = string.Format("Avg : {0}", 0);

                contextMenuStrip.Items[8].Tag = 0;
                contextMenuStrip.Items[9].Tag = 0;
            }

        }

        private void ToolStripMenuItem_Click(Object _sender, EventArgs e)
        {
            ContextMenuStrip _ContextMenuStrip;
            ToolStripMenuItem _ToolStripMenuItem;
            SaveFileDialog _FileDialog;
            FpSpread _DataGridView;
            Process _Process;
            System.Data.DataRowView _DataRowView;

            _ToolStripMenuItem = (ToolStripMenuItem)_sender;
            _ContextMenuStrip = (ContextMenuStrip)_ToolStripMenuItem.Owner;
            _DataGridView = (FpSpread)_ContextMenuStrip.SourceControl;

            //엑셀 저장, 엑셀 저장 & 열기
            if (_ToolStripMenuItem.Text.Equals(this.MenuListText[0]) || _ToolStripMenuItem.Text.Equals(this.MenuListText[1]))
            {
                _FileDialog = new SaveFileDialog()
                {
                    DefaultExt = "*.xls",
                    Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*"
                };

                if (_FileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (!ExportExcel(_FileDialog.FileName, _DataGridView))
                        return;

                    if (_ToolStripMenuItem.Text.Equals(this.MenuListText[1])) //'엑셀 저장 & 열기
                    {
                        _Process = new Process();
                        _Process = Process.Start(_FileDialog.FileName);
                    }
                }

                return;
            }

            //출력
            if (_ToolStripMenuItem.Text.Equals(MenuListText[2]))
            {
                //'_PrintDocument = New Printing.PrintDocument
                //'_PrintDocument.PrinterSettings.DefaultPageSettings.Margins = New Printing.Margins(30, 40, 30, 30)
                //'AddHandler _PrintDocument.BeginPrint, AddressOf Document_BeginPrint
                //'AddHandler _PrintDocument.PrintPage, AddressOf Document_PrintPage
                //
                //'_PrintPreviewDialog = New PrintPreviewDialog
                //'_PrintPreviewDialog.Document = _PrintDocument
                //'_PrintPreviewDialog.ShowDialog(_DataGridView.FindForm)
                return;
            }

            //복사
            if (_ToolStripMenuItem.Text.Equals(MenuListText[3]))
            {
                _DataGridView.ActiveSheet.ClipboardCopy(_DataGridView.ActiveSheet.GetSelection(0));
                return;
            }


            //행추가
            if (_ToolStripMenuItem.Text.Equals(MenuListText[5]))
            {
                if (_DataGridView.DataSource != null && _DataGridView.DataSource is System.Data.DataView)
                {
                    if (((IDataGridAgent)this).AddRows == AddRows.True)
                    {
                        ((System.Data.DataView)_DataGridView.DataSource).AddNew().EndEdit();
                    }
                }

                return;
            }


            //행삭제
            if (_ToolStripMenuItem.Text.Equals(MenuListText[6]))
            {
                if (_DataGridView.DataSource != null && _DataGridView.DataSource is System.Data.DataView)
                    if (((IDataGridAgent)this).DeleteRows == DeleteRows.True)
                        ((System.Data.DataView)_DataGridView.DataSource).Delete(_DataGridView.ActiveSheet.ActiveRowIndex);

                return;
            }

            _DataRowView = null;

            //행복사
            if (_ToolStripMenuItem.Text.Equals(MenuListText[7]))
            {
                if (_DataGridView.DataSource is System.Data.DataView)
                {
                    if (((IDataGridAgent)this).AddRows == AddRows.True)
                    {
                        try
                        {
                            _DataRowView = ((System.Data.DataView)_DataGridView.DataSource).AddNew();

                            for (int i = 0; i < _DataGridView.ActiveSheet.ColumnCount; i++)
                            {
                                _DataRowView.Row[i] = _DataGridView.ActiveSheet.Cells[_DataGridView.ActiveSheet.ActiveRowIndex, i].Value;
                            }
                        }
                        finally
                        {
                            _DataRowView?.EndEdit();
                        }
                    }
                }

                return;
            }
        }

        private static bool ExportExcel(string _Path, FpSpread _DataGridView)
        {
            bool orgProtect;

            orgProtect = _DataGridView.ActiveSheet.Protect;

            try
            {
                _DataGridView.ActiveSheet.Protect = false;
                return _DataGridView.SaveExcel(_Path, FarPoint.Win.Spread.Model.IncludeHeaders.BothCustomOnly);
            }
            catch (Exception ex)
            {
                DiagnosticsTool.MyTrace(ex);
                return false;
            }
            finally
            {
                _DataGridView.ActiveSheet.Protect = orgProtect;
            }
        }

        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith startsWith, AutoComplete autoComplete, string name, params System.Windows.Forms.Control[] controls)
        {
            FpSpread _DataGridView;
            FilterAttribute _FiterAttribute;
            TextBox textBox;

            _DataGridView = ((FpSpread)((IDataGridAgent)this).GridControl);

            _FiterAttribute = new FilterAttribute()
            {
                DataGridView = _DataGridView,
                ColumnName = name,
                IsSearchAll = (searchAll == SearchAll.True),
                IsStartsWith = (startsWith == StartsWith.True),
                AutoCompleteMode = (AutoCompleteMode)Enum.Parse(typeof(AutoCompleteMode), autoComplete.ToString()),
                AutoCompleteStringCollection = new AutoCompleteStringCollection()
            };

            foreach (System.Windows.Forms.Control control in controls)
            {
                if (control is TextBox)
                {
                    textBox = (TextBox)control;

                    this.FilterControl.Add(textBox, _FiterAttribute);

                    if (_FiterAttribute.AutoCompleteMode != AutoCompleteMode.None)
                    {
                        textBox.TextChanged -= this.Control_TextChanged;
                        textBox.TextChanged += this.Control_TextChanged;

                        textBox.AutoCompleteMode = _FiterAttribute.AutoCompleteMode;
                        textBox.AutoCompleteCustomSource = _FiterAttribute.AutoCompleteStringCollection;
                        textBox.AutoCompleteSource = AutoCompleteSource.CustomSource;
                    }

                    textBox.DoubleBuffered(true);
                }
            }
        }
        void IDataGridAgent.AddColumnFiter(SearchAll searchAll, StartsWith _StartsWith, AutoComplete autoComplete, int _Index, params System.Windows.Forms.Control[] controls)
        {
            ((IDataGridAgent)this).AddColumnFiter(searchAll, _StartsWith, autoComplete, ((FpSpread)((IDataGridAgent)this).GridControl).ActiveSheet.Columns[_Index].DataField, controls);
        }

        void IDataGridAgent.RemoveColumnFiter(params System.Windows.Forms.Control[] controls)
        {
            if (controls is TextBox[])
                foreach (TextBox textBox in (TextBox[])controls)
                {
                    if (this.FilterControl.ContainsKey(textBox))
                    {
                        if (this.FilterControl[textBox].AutoCompleteMode != AutoCompleteMode.None)
                        {
                            textBox.TextChanged -= this.Control_TextChanged;
                            textBox.AutoCompleteMode = AutoCompleteMode.None;
                            textBox.AutoCompleteCustomSource = null;
                            textBox.AutoCompleteSource = AutoCompleteSource.None;
                        }

                        this.FilterControl.Remove(textBox);
                    }
                }
        }

        void Control_TextChanged(object sender, EventArgs e)
        {
            FpSpread _DataGridView;
            System.Data.DataView _DataView;
            TextBox _TextBox;
            FilterAttribute _FilterAttribute;
            StringBuilder _StringBuilder;
            //string _Tmp;
            decimal _decimal;

            _TextBox = (TextBox)sender;

            _FilterAttribute = this.FilterControl[_TextBox];

            _DataGridView = _FilterAttribute.DataGridView;

            if (_DataGridView.DataSource == null)
                return;

            _DataView = null;

            if (_DataGridView.DataSource is System.Data.DataView)
                _DataView = (System.Data.DataView)_DataGridView.DataSource;

            if (_DataGridView.DataSource is System.Data.DataSet)
                _DataView = ((System.Data.DataSet)_DataGridView.DataSource).Tables[0].DefaultView;

            if (_DataGridView.DataSource is System.Data.DataTable)
                _DataView = ((System.Data.DataTable)_DataGridView.DataSource).DefaultView;

            if (_DataView == null)
                return;

            _StringBuilder = new StringBuilder();
            if (!_TextBox.Text.Equals(""))
            {
                foreach (FarPoint.Win.Spread.Column _DataGridViewColumn in _DataGridView.ActiveSheet.Columns)
                {
                    if (!_DataGridViewColumn.Visible)//보이는 컬럼만 검색
                        continue;

                    if (_DataGridViewColumn.DataField == _FilterAttribute.ColumnName && !_FilterAttribute.IsSearchAll)
                    {
                        if (_DataGridViewColumn.HorizontalAlignment == CellHorizontalAlignment.Right && _DataGridViewColumn.VerticalAlignment == CellVerticalAlignment.Center)
                        {
                            if (_TextBox.Text.ToTryDecimal(out _decimal))
                            {
                                if (_FilterAttribute.IsStartsWith)
                                    _StringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _TextBox.Text + "%' ";
                                else
                                    _StringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _TextBox.Text + "%' ";
                            }
                            else
                            {
                                if (_FilterAttribute.IsStartsWith)
                                    _StringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _TextBox.Text + "%' ";
                                else
                                    _StringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _TextBox.Text + "%' ";
                            }
                        }
                        else
                        {
                            if (_FilterAttribute.IsStartsWith)
                                _StringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _TextBox.Text + "%' ";
                            else
                                _StringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _TextBox.Text + "%' ";
                        }
                        break;
                    }

                    if (_FilterAttribute.IsSearchAll)
                    {
                        if (_DataGridViewColumn.HorizontalAlignment == CellHorizontalAlignment.Right && _DataGridViewColumn.VerticalAlignment == CellVerticalAlignment.Center)
                        {
                            if (_TextBox.Text.ToTryDecimal(out _decimal))
                            {
                                if (_FilterAttribute.IsStartsWith)
                                    _StringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '" + _TextBox.Text + "%' ";
                                else

                                    _StringBuilder.AppendFormat("OR Convert([{0}], 'System.String') LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR Convert([" + _DataGridViewColumn.DataPropertyName + "], 'System.String') LIKE '%" + _TextBox.Text + "%' ";
                            }
                            else
                            {
                                if (_FilterAttribute.IsStartsWith)
                                    _StringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _TextBox.Text + "%' ";
                                else
                                    _StringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                                //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _TextBox.Text + "%' ";
                            }
                        }
                        else
                        {
                            if (_FilterAttribute.IsStartsWith)
                                _StringBuilder.AppendFormat("OR [{0}] LIKE '{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '" + _TextBox.Text + "%' ";
                            else
                                _StringBuilder.AppendFormat("OR [{0}] LIKE '%{1}%' ", _DataGridViewColumn.DataField, _TextBox.Text);
                            //_Tmp += "OR [" + _DataGridViewColumn.DataPropertyName + "] LIKE '%" + _TextBox.Text + "%' ";
                        }
                    }
                }
            }

            try
            {
                if (_StringBuilder.ToString().StartsWith("OR "))
                    _DataView.RowFilter = _StringBuilder.ToString(3, _StringBuilder.Length - 3);
                else
                    _DataView.RowFilter = _StringBuilder.ToString();
            }
            catch (Exception _Exception)
            {
                DiagnosticsTool.MyTrace(_Exception);
                _DataView.RowFilter = "";
            }
        }

        private void DataGridView_DataSourceChanged(object sender, EventArgs e)
        {
            FpSpread _DataGridView;
            System.Data.DataView _DataView;

            _DataGridView = (FpSpread)sender;

            if (_DataGridView.DataSource == null)
                return;

            if (this.FilterControl == null)
                return;

            _DataView = null;

            if (_DataGridView.DataSource is System.Data.DataView)
                _DataView = (System.Data.DataView)_DataGridView.DataSource;

            if (_DataGridView.DataSource is System.Data.DataSet)
                _DataView = ((System.Data.DataSet)_DataGridView.DataSource).Tables[0].DefaultView;

            if (_DataGridView.DataSource is System.Data.DataTable)
                _DataView = ((System.Data.DataTable)_DataGridView.DataSource).DefaultView;

            if (_DataView == null)
                return;

            foreach (FilterAttribute _FilterAttribute in this.FilterControl.Values)
            {
                //_FilterAttribute.AutoCompleteStringCollection.Clear();

                for (int i = 0; i < _DataView.Table.Rows.Count; i++)
                {
                    if (_DataView.Table.Columns.Contains(_FilterAttribute.ColumnName)
                        && !_FilterAttribute.AutoCompleteStringCollection.Contains(_DataView.Table.Rows[i][_FilterAttribute.ColumnName].ToString()))
                        _FilterAttribute.AutoCompleteStringCollection.Add(_DataView.Table.Rows[i][_FilterAttribute.ColumnName].ToString());
                }
            }

        }

        private void DataGridView_DataBindingComplete(object sender, DataColumnConfigureEventArgs e)
        {
            FpSpread _DataGridView;
            System.Data.DataView _DataView;

            _DataGridView = (FpSpread)sender;

            if (_DataGridView.DataSource == null)
                return;

            _DataView = null;

            if (_DataGridView.DataSource is System.Data.DataView)
                _DataView = (System.Data.DataView)_DataGridView.DataSource;

            if (_DataGridView.DataSource is System.Data.DataSet)
                _DataView = ((System.Data.DataSet)_DataGridView.DataSource).Tables[0].DefaultView;

            if (_DataGridView.DataSource is System.Data.DataTable)
                _DataView = ((System.Data.DataTable)_DataGridView.DataSource).DefaultView;

            if (_DataView == null)
                return;

            //foreach (Row _Rows in _DataGridView.ActiveSheet.Rows)
            //{
            //    //if (_Rows.IsNewRow) continue;
            //    if (_DataGridView.ActiveSheet.SheetCorner.Cells[0, 0].Value != null)
            //        break;
            //
            //    _DataGridView.ActiveSheet.SheetCorner.Cells[0, 0].Value = (_Rows.Index + 1).ToString();
            //}

            _DataGridView.ActiveSheet.SheetCorner.Cells[0, 0].Value = _DataGridView.ActiveSheet.Rows.Count.ToString() + "/" + _DataView.Table.Rows.Count.ToString();

            //_DataGridView.ActiveSheet.SheetCorner = 30 + (_DataGridView.ActiveSheet.Rows.Count.ToString().Length * 10);
            //_DataGridView.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        }

        private void SetSkin(FpSpread _DataGridView)
        {
            string _SkinName;
            //string _Tmp;
            //System.Drawing.Color _Color;
            //System.Drawing.Font _Font;
            //string[] _Tmps;

            _SkinName = (string)Config.Client.GetAttribute("SkinName");

            if (_SkinName == null)
                _SkinName = this.GetAttribute("SkinName");

            if (_SkinName == null)
                return;

            //_DataGridView.EnableHeadersVisualStyles = false;

            //_Color = this.GetAttributeColor(string.Format("{0}.BackgroundColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.BackgroundColor = _Color;

            //_Color = this.GetAttributeColor(string.Format("{0}.ColumnHeadersDefaultCellStyle.BackColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.ColumnHeadersDefaultCellStyle.BackColor = _Color;

            //_Color = this.GetAttributeColor(string.Format("{0}.ColumnHeadersDefaultCellStyle.ForeColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.ColumnHeadersDefaultCellStyle.ForeColor = _Color;

            //_Font = this.GetAttributeFont(_DataGridView.Font, string.Format("{0}.ColumnHeadersDefaultCellStyle.Font", _SkinName));
            //if (_Font != null)
            //    _DataGridView.ColumnHeadersDefaultCellStyle.Font = _Font;

            //_DataGridView.RowHeadersDefaultCellStyle.Font = _Font;

            //_Tmp = this.GetAttribute(string.Format("{0}.ColumnHeadersBorderStyle", _SkinName));
            //_DataGridView.ColumnHeadersBorderStyle = (DataGridViewHeaderBorderStyle)Enum.Parse(typeof(DataGridViewHeaderBorderStyle), _Tmp);

            //_Color = this.GetAttributeColor(string.Format("{0}.RowHeadersDefaultCellStyle.BackColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.RowHeadersDefaultCellStyle.BackColor = _Color;

            //_Color = this.GetAttributeColor(string.Format("{0}.RowHeadersDefaultCellStyle.ForeColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.RowHeadersDefaultCellStyle.ForeColor = _Color;

            //_Tmp = this.GetAttribute(string.Format("{0}.RowHeadersBorderStyle", _SkinName));
            //_DataGridView.RowHeadersBorderStyle = (DataGridViewHeaderBorderStyle)Enum.Parse(typeof(DataGridViewHeaderBorderStyle), _Tmp);

            //_Color = this.GetAttributeColor(string.Format("{0}.RowsDefaultCellStyle.SelectionBackColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.RowsDefaultCellStyle.SelectionBackColor = _Color;

            //_Color = this.GetAttributeColor(string.Format("{0}.RowsDefaultCellStyle.SelectionForeColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.RowsDefaultCellStyle.SelectionForeColor = _Color;

            //_Color = this.GetAttributeColor(string.Format("{0}.AlternatingRowsDefaultCellStyle.BackColor", _SkinName));
            //if (_Color != System.Drawing.Color.Empty)
            //    _DataGridView.AlternatingRowsDefaultCellStyle.BackColor = _Color;

            ////_DataGridView.AutoResizeRowHeadersWidth(DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders);
            ////_DataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            //_DataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.EnableResizing;
        }

        protected struct FilterAttribute
        {
            public FpSpread DataGridView { get; set; }
            public string ColumnName { get; set; }
            public bool IsSearchAll { get; set; }
            public bool IsStartsWith { get; set; }
            public AutoCompleteMode AutoCompleteMode { get; set; }
            public AutoCompleteStringCollection AutoCompleteStringCollection { get; set; }
        }
    }
}