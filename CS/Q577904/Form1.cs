using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using DevExpress.XtraRichEdit;
using DevExpress.XtraRichEdit.Commands;
using DevExpress.XtraGrid.Views.Grid.ViewInfo;
using DevExpress.XtraGrid.Views.Base;
using DevExpress.XtraRichEdit.API.Native;
using DevExpress.Office.Utils;
using DevExpress.XtraGrid.Columns;
using DevExpress.XtraRichEdit.Commands.Internal;
using DevExpress.XtraRichEdit.Export.Html;

namespace Q577904 {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
            gridControl1.DataSource = CreateTable(20);
        }

        private DataTable CreateTable(int rowCount) {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add("String", typeof(string));
            dataTable.Columns.Add("Int", typeof(int));
            dataTable.Columns.Add("Date", typeof(DateTime));
            for(int i = 0; i < rowCount; i++) {
                dataTable.Rows.Add("Row" + i, i, DateTime.Today.AddDays(i));
            }
            return dataTable;
        }

        private void gridControl1_KeyDown(object sender, KeyEventArgs e) {
            if(e.KeyData == (Keys.C | Keys.Control) || e.KeyData == (Keys.Insert | Keys.Control)) {
                CopyToClipboard();
                e.Handled = true;
            }
        }

        private void CopyColumns(IOrderedEnumerable<IGrouping<GridColumn, GridCell>> columns, RichEditDocumentServer srv, Table table)
        {
            int i = 0;
            foreach(var column in columns) {
                TableCell cell = GetCell(table, 0, i);
                srv.Document.InsertText(cell.Range.Start, column.Key.GetTextCaption());
                cell.BackgroundColor = Color.Gray;
                cell.PreferredWidthType = WidthType.Fixed;
                using(Graphics g = this.CreateGraphics()) {
                    cell.PreferredWidth = Units.PixelsToDocumentsF(column.Key.VisibleWidth, g.DpiX);
                }
                i++;
            }
        }

        private static TableCell GetCell(Table table, int i, int j) {
            TableCell cell = table.Cell(i, j);
            return cell;
        }

        private void CopyCells(IOrderedEnumerable<IGrouping<Int32, GridCell>> rows, RichEditDocumentServer srv, Table table)
        {
            GridViewInfo viewInfo = (GridViewInfo)gridView1.GetViewInfo();
            int i = 1;
            foreach(var row in rows) {
                row.OrderBy(z => z.Column.VisibleIndex);
                int j = 0;
                foreach(var c in row) {
                    TableCell cell = GetCell(table, i, j);
                    srv.Document.InsertText(cell.Range.Start, gridView1.GetRowCellDisplayText(c.RowHandle, c.Column));
                    GridCellInfo gridCellInfo = GetGridCellInfo(viewInfo, c);
                    cell.BackgroundColor = gridCellInfo.Appearance.BackColor;
                    j++;
                }
                i++;
            }
        }

        private void SetBorders(Table table) {
            float thickness = 1f;
            table.Borders.Left.LineColor = table.Borders.Right.LineColor = table.Borders.InsideVerticalBorder.LineColor = gridView1.PaintAppearance.VertLine.BackColor;
            table.Borders.Left.LineThickness = table.Borders.Right.LineThickness = table.Borders.InsideVerticalBorder.LineThickness = thickness;
            table.Borders.Left.LineStyle = table.Borders.Right.LineStyle = table.Borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single;
            table.Borders.Bottom.LineColor = table.Borders.Top.LineColor = table.Borders.InsideHorizontalBorder.LineColor = gridView1.PaintAppearance.HorzLine.BackColor;
            table.Borders.Bottom.LineThickness = table.Borders.Top.LineThickness = table.Borders.InsideHorizontalBorder.LineThickness = thickness;
            table.Borders.Bottom.LineStyle = table.Borders.Top.LineStyle = table.Borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single;
        }

        private void CopyToClipboard() {
            GridCell[] cells = gridView1.GetSelectedCells();
            var rows = from c in cells
                       group c by c.RowHandle
                        into gr
                       orderby gr.Key
                       select gr;
            var columns = from c in cells
                          group c by c.Column
                           into gr
                          orderby gr.Key.VisibleIndex
                          select gr;
            RichEditDocumentServer srv = new RichEditDocumentServer();
            srv.CreateNewDocument();                       
            Table table = srv.Document.Tables.Create(srv.Document.CaretPosition, rows.Count() + 1, columns.Count());
            SetBorders(table);
            CopyColumns(columns, srv, table);
            CopyCells(rows, srv, table);

            DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions options = srv.Options.Export.Html;
            options.ExportRootTag = ExportRootTag.Html;
            options.Encoding  = Encoding.UTF8;
            options.CssPropertiesExportType = CssPropertiesExportType.Inline;
            options.UriExportType = UriExportType.Absolute;
            options.EmbedImages = false;
            
            string htmlContent = srv.HtmlText;
            string cfHtml = CF_HTMLHelper.GetHtmlClipboardFormat(htmlContent);

            IDataObject dataObject = new DataObject();
            dataObject.SetData(DataFormats.Text, srv.Text);
            dataObject.SetData(DataFormats.Html, cfHtml);
            Clipboard.SetDataObject(dataObject, true);

        }

        private GridCellInfo GetGridCellInfo(GridViewInfo viewInfo, GridCell cell) {
            GridCellInfo gridCellInfo = viewInfo.GetGridCellInfo(cell.RowHandle, cell.Column);
            gridCellInfo.State &= ~(GridRowCellState.Focused | GridRowCellState.FocusedCell | GridRowCellState.Selected);
            System.Reflection.MethodInfo method = viewInfo.GetType().GetMethod("UpdateCellAppearanceCore", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic);
            method.Invoke(viewInfo, new object[] { gridCellInfo, true, true, null });
            return gridCellInfo;
        }
    }

}
