Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Linq
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit
Imports DevExpress.XtraRichEdit.Commands
Imports DevExpress.XtraGrid.Views.Grid.ViewInfo
Imports DevExpress.XtraGrid.Views.Base
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.Office.Utils
Imports DevExpress.XtraGrid.Columns
Imports DevExpress.XtraRichEdit.Commands.Internal
Imports DevExpress.XtraRichEdit.Export.Html

Namespace Q577904
    Partial Public Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            gridControl1.DataSource = CreateTable(20)
        End Sub

        Private Function CreateTable(ByVal rowCount As Integer) As DataTable
            Dim dataTable As New DataTable()
            dataTable.Columns.Add("String", GetType(String))
            dataTable.Columns.Add("Int", GetType(Integer))
            dataTable.Columns.Add("Date", GetType(Date))
            For i As Integer = 0 To rowCount - 1
                dataTable.Rows.Add("Row" & i, i, Date.Today.AddDays(i))
            Next i
            Return dataTable
        End Function

        Private Sub gridControl1_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles gridControl1.KeyDown
            If e.KeyData = (Keys.C Or Keys.Control) OrElse e.KeyData = (Keys.Insert Or Keys.Control) Then
                CopyToClipboard()
                e.Handled = True
            End If
        End Sub

        Private Sub CopyColumns(ByVal columns As IOrderedEnumerable(Of IGrouping(Of GridColumn, GridCell)), ByVal srv As RichEditDocumentServer, ByVal table As Table)
            Dim i As Integer = 0
            For Each column In columns
                Dim cell As TableCell = GetCell(table, 0, i)
                srv.Document.InsertText(cell.Range.Start, column.Key.GetTextCaption())
                cell.BackgroundColor = Color.Gray
                cell.PreferredWidthType = WidthType.Fixed
                Using g As Graphics = Me.CreateGraphics()
                    cell.PreferredWidth = Units.PixelsToDocumentsF(column.Key.VisibleWidth, g.DpiX)
                End Using
                i += 1
            Next column
        End Sub

        Private Shared Function GetCell(ByVal table As Table, ByVal i As Integer, ByVal j As Integer) As TableCell
            Dim cell As TableCell = table.Cell(i, j)
            Return cell
        End Function

        Private Sub CopyCells(ByVal rows As IOrderedEnumerable(Of IGrouping(Of Int32, GridCell)), ByVal srv As RichEditDocumentServer, ByVal table As Table)
            Dim viewInfo As GridViewInfo = CType(gridView1.GetViewInfo(), GridViewInfo)
            Dim i As Integer = 1
            For Each row In rows
                row.OrderBy(Function(z) z.Column.VisibleIndex)
                Dim j As Integer = 0
                For Each c In row
                    Dim cell As TableCell = GetCell(table, i, j)
                    srv.Document.InsertText(cell.Range.Start, gridView1.GetRowCellDisplayText(c.RowHandle, c.Column))
                    Dim gridCellInfo As GridCellInfo = GetGridCellInfo(viewInfo, c)
                    cell.BackgroundColor = gridCellInfo.Appearance.BackColor
                    j += 1
                Next c
                i += 1
            Next row
        End Sub

        Private Sub SetBorders(ByVal table As Table)
            Dim thickness As Single = 1F
            table.Borders.InsideVerticalBorder.LineColor = gridView1.PaintAppearance.VertLine.BackColor
            table.Borders.Right.LineColor = table.Borders.InsideVerticalBorder.LineColor
            table.Borders.Left.LineColor = table.Borders.Right.LineColor
            table.Borders.InsideVerticalBorder.LineThickness = thickness
            table.Borders.Right.LineThickness = table.Borders.InsideVerticalBorder.LineThickness
            table.Borders.Left.LineThickness = table.Borders.Right.LineThickness
            table.Borders.InsideVerticalBorder.LineStyle = TableBorderLineStyle.Single
            table.Borders.Right.LineStyle = table.Borders.InsideVerticalBorder.LineStyle
            table.Borders.Left.LineStyle = table.Borders.Right.LineStyle
            table.Borders.InsideHorizontalBorder.LineColor = gridView1.PaintAppearance.HorzLine.BackColor
            table.Borders.Top.LineColor = table.Borders.InsideHorizontalBorder.LineColor
            table.Borders.Bottom.LineColor = table.Borders.Top.LineColor
            table.Borders.InsideHorizontalBorder.LineThickness = thickness
            table.Borders.Top.LineThickness = table.Borders.InsideHorizontalBorder.LineThickness
            table.Borders.Bottom.LineThickness = table.Borders.Top.LineThickness
            table.Borders.InsideHorizontalBorder.LineStyle = TableBorderLineStyle.Single
            table.Borders.Top.LineStyle = table.Borders.InsideHorizontalBorder.LineStyle
            table.Borders.Bottom.LineStyle = table.Borders.Top.LineStyle
        End Sub

        Private Sub CopyToClipboard()
            Dim cells() As GridCell = gridView1.GetSelectedCells()
            Dim rows = From c In cells _
                       Group c By c.RowHandle Into gr = Group _
                       Order By RowHandle _
                       Select gr
            Dim columns = From c In cells _
                          Group c By c.Column Into gr = Group _
                          Order By Column.VisibleIndex _
                          Select gr
            Dim srv As New RichEditDocumentServer()
            srv.CreateNewDocument()
            Dim table As Table = srv.Document.InsertTable(srv.Document.CaretPosition, rows.Count() + 1, columns.Count())
            SetBorders(table)
            CopyColumns(columns, srv, table)
            CopyCells(rows, srv, table)

            Dim options As DevExpress.XtraRichEdit.Export.HtmlDocumentExporterOptions = srv.Options.Export.Html
            options.ExportRootTag = ExportRootTag.Html
            options.Encoding = Encoding.UTF8
            options.CssPropertiesExportType = CssPropertiesExportType.Inline
            options.UriExportType = UriExportType.Absolute
            options.EmbedImages = False

            Dim htmlContent As String = srv.HtmlText
            Dim cfHtml As String = CF_HTMLHelper.GetHtmlClipboardFormat(htmlContent)

            Dim dataObject As IDataObject = New DataObject()
            dataObject.SetData(DataFormats.Text, srv.Text)
            dataObject.SetData(DataFormats.Html, cfHtml)
            Clipboard.SetDataObject(dataObject, True)

        End Sub

        Private Function GetGridCellInfo(ByVal viewInfo As GridViewInfo, ByVal cell As GridCell) As GridCellInfo
            Dim gridCellInfo As GridCellInfo = viewInfo.GetGridCellInfo(cell.RowHandle, cell.Column)
            gridCellInfo.State = gridCellInfo.State And Not(GridRowCellState.Focused Or GridRowCellState.FocusedCell Or GridRowCellState.Selected)
            Dim method As System.Reflection.MethodInfo = viewInfo.GetType().GetMethod("UpdateCellAppearanceCore", System.Reflection.BindingFlags.Instance Or System.Reflection.BindingFlags.NonPublic)
            method.Invoke(viewInfo, New Object() { gridCellInfo })
            Return gridCellInfo
        End Function
    End Class

End Namespace
