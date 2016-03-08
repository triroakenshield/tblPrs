Imports MyAcAs = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports RTree

Public Class MyTable
    'Public Shared MinColW As Double = 1
    'Public Shared MinRowH As Double = 1
    Private vert, horz As List(Of Line) 'Список линий формирующих таблицу
    Public Cells(,) As MyCell 'Массив ячеек
    Friend wTree As RTree(Of MyCell) 'Дерево для поиска ячеек

    Public Enum Orent
        Vert ' вертикальна
        Horz ' горизонтална
        None ' неопределенна
    End Enum

    Public Shared Function isOrto(wL As Line) As Orent
        'Определяем орентацию линии - вертикальна или горизонтална
        Dim wValue As Double = wL.Angle / Math.PI
        Dim delta As Double = 0.05
        wValue = wValue - Math.Truncate(wValue + delta / 2)
        If Math.Abs(wValue) <= delta Then
            Return Orent.Horz
        ElseIf (Math.Abs(wValue) < 0.5 + delta) And (Math.Abs(wValue) > 0.5 - delta) Then
            Return Orent.Vert
        Else
            Return Orent.None
        End If
    End Function

    Private Shared Function CompareByX(l1 As Line, l2 As Line) As Integer
        If l1.StartPoint.X > l2.StartPoint.X Then
            Return 1
        ElseIf l1.StartPoint.X = l2.StartPoint.X Then
            Return 0
        Else
            Return -1
        End If
    End Function

    Private Shared Function CompareByY(l1 As Line, l2 As Line) As Integer
        If l1.StartPoint.Y > l2.StartPoint.Y Then
            Return 1
        ElseIf l1.StartPoint.Y = l2.StartPoint.Y Then
            Return 0
        Else
            Return -1
        End If
    End Function

    Private Shared Function GetSelect(ed As Editor) As ObjectId()
        'Получаем от пользователя набор данных для парсинга
        Dim PSResult As PromptSelectionResult
        Dim wTV() As TypedValue = {New TypedValue(DxfCode.Operator, "<or"), _
                                   New TypedValue(DxfCode.Start, "LINE"), _
                                   New TypedValue(DxfCode.Start, "LWPOLYLINE"), _
                                   New TypedValue(DxfCode.Start, "TEXT"), _
                                   New TypedValue(DxfCode.Start, "MTEXT"), _
                                   New TypedValue(DxfCode.Operator, "or>")}
        Dim wSF As New SelectionFilter(wTV)
        PSResult = ed.GetSelection(wSF)
        If PSResult.Status = PromptStatus.OK Then
            Return PSResult.Value.GetObjectIds()
        Else
            Return Nothing
        End If
    End Function

    Private Shared Function PolyToLine(pl As Polyline) As List(Of Line)
        Dim wList As New List(Of Line)
        Dim wL As Line
        For i = 0 To pl.NumberOfVertices - 2
            wL = New Line(pl.GetPoint3dAt(i), pl.GetPoint3dAt(i + 1))
            wList.Add(wL)
        Next
        Return wList
    End Function

    Private Sub New(nvert As List(Of Line), nhorz As List(Of Line))
        'Формируем "пустую" таблицу из линий
        Me.vert = nvert
        Me.horz = nhorz
        Dim CC, RC As Integer
        CC = Me.GetCols()
        RC = Me.GetRows()
        ReDim Me.Cells(CC, RC)
        Me.wTree = New RTree(Of MyCell)()
        Dim wLine As Line
        Dim nCell As MyCell
        For i = 0 To CC - 1
            For j = 0 To RC - 1
                wLine = Me.GetCellBox(i, j)
                nCell = New MyCell(wLine, i, j, "")
                Me.Cells(i, j) = nCell
                Me.wTree.Add(nCell.GetRectangle, nCell)
            Next
        Next
    End Sub

    Public Sub SetValue(wt As DBText)
        'Заполняем таблицу
        If wt.Bounds IsNot Nothing Then
            Dim tExtent As Extents3d = wt.Bounds
            Dim X, Y As Double
            X = (tExtent.MaxPoint.X + tExtent.MinPoint.X) / 2
            Y = (tExtent.MaxPoint.Y + tExtent.MinPoint.Y) / 2
            Dim wP As New Point(X, Y, 0)
            Dim wList As List(Of MyCell) = Me.wTree.Nearest(wP, wt.Height / 2)
            If wList IsNot Nothing Then
                If wList.Count > 0 Then wList(0).Value = wt.TextString
            End If
        End If
    End Sub

    Public Sub SetValue(wt As MText)
        'Заполняем таблицу
        If wt.Bounds IsNot Nothing Then
            Dim tExtent As Extents3d = wt.Bounds
            Dim X, Y As Double
            X = (tExtent.MaxPoint.X + tExtent.MinPoint.X) / 2
            Y = (tExtent.MaxPoint.Y + tExtent.MinPoint.Y) / 2
            Dim wP As New Point(X, Y, 0)
            Dim wList As List(Of MyCell) = Me.wTree.Nearest(wP, 1)
            If wList IsNot Nothing Then
                If wList.Count > 0 Then wList(0).Value = wt.Text
            End If
        End If
    End Sub

    Private Shared Function CrTbl(wList As List(Of Line)) As MyTable
        'Формируем "пустую" таблицу из линий
        Dim nvert, nhorz, overt, ohorz As List(Of Line)
        nvert = wList.FindAll(Function(l) isOrto(l) = Orent.Vert)
        nvert.Sort(AddressOf CompareByX)
        nhorz = wList.FindAll(Function(l) isOrto(l) = Orent.Horz)
        nhorz.Sort(AddressOf CompareByY)
        '
        Dim MinColW, MinRowH As Double
        MinColW = Math.Abs(nvert(0).StartPoint.X - nvert(nvert.Count - 1).StartPoint.X) * 0.01
        MinRowH = Math.Abs(nhorz(0).StartPoint.Y - nhorz(nhorz.Count - 1).StartPoint.Y) * 0.01
        '
        Dim ol As Line = Nothing
        overt = New List(Of Line)
        For Each l In nvert
            If ol Is Nothing Then
                ol = l
                overt.Add(l)
            Else
                If Math.Abs(l.StartPoint.X - ol.StartPoint.X) > MinColW Then
                    ol = l
                    overt.Add(l)
                End If
            End If
        Next
        '
        ohorz = New List(Of Line)
        For Each l In nhorz
            If ol Is Nothing Then
                ol = l
                ohorz.Add(l)
            Else
                If Math.Abs(l.StartPoint.Y - ol.StartPoint.Y) > MinRowH Then
                    ol = l
                    ohorz.Add(l)
                End If
            End If
        Next
        Return New MyTable(overt, ohorz)
    End Function

    Public Shared Function CrTbl(acDoc As MyAcAs.Document) As MyTable
        'Создаём таблицу
        Dim ed As Editor = acDoc.Editor
        Dim objIdArray() As ObjectId = MyTable.GetSelect(ed) 'Получаем от пользователя набор данных для парсинга
        If objIdArray IsNot Nothing Then
            Dim dbObj As DBObject
            Dim wList As New List(Of Line)
            Dim wTList As New List(Of DBText)
            Dim wMTList As New List(Of MText)
            Using tr As Transaction = acDoc.Database.TransactionManager.StartTransaction
                Try
                    For Each objId As ObjectId In objIdArray
                        dbObj = tr.GetObject(objId, OpenMode.ForRead)
                        'Сортируем полученные объекты
                        Select Case True
                            Case TypeOf dbObj Is Line
                                wList.Add(dbObj)
                            Case TypeOf dbObj Is Polyline
                                wList.AddRange(MyTable.PolyToLine(dbObj))
                            Case TypeOf dbObj Is DBText
                                wTList.Add(dbObj)
                            Case TypeOf dbObj Is MText
                                wMTList.Add(dbObj)
                        End Select
                    Next
                    tr.Commit()
                Catch ex As Exception
                    ed.WriteMessage(ex.ToString())
                    tr.Abort()
                End Try
            End Using
            '
            Dim wMTbl As MyTable = MyTable.CrTbl(wList)
            'Заполняем текстом
            For Each wt In wTList
                wMTbl.SetValue(wt)
            Next
            For Each wmt In wMTList
                wMTbl.SetValue(wmt)
            Next
            Return wMTbl
        Else
            Return Nothing
        End If
    End Function

    Public Function GetCols() As Integer
        Return vert.Count - 1
    End Function

    Public Function GetColW(i As Integer) As Double
        Dim res As Double = Math.Abs(vert(i + 1).StartPoint.X - vert(i).StartPoint.X)
        If res = 0 Then res = 1 '?!
        Return res
    End Function

    Public Function GetRows() As Integer
        Return horz.Count - 1
    End Function

    Public Function GetRowH(j As Integer) As Double
        Dim res As Double = Math.Abs(horz(j + 1).StartPoint.Y - horz(j).StartPoint.Y)
        If res = 0 Then res = 1 '?!
        Return res
    End Function

    Public Function GetCellBox(i As Integer, j As Integer) As Line
        'Получаем диагональную линию в нужной ячейке (размер)
        Dim p1, p2 As Point3d
        p1 = New Point3d(vert(i).StartPoint.X, horz(j).StartPoint.Y, 0)
        p2 = New Point3d(vert(i + 1).StartPoint.X, horz(j + 1).StartPoint.Y, 0)
        Return New Line(p1, p2)
    End Function

    Public Function CrTbl(ip As Point3d) As Table
        'Создаём ACAD-таблицу
        Dim res As New Table()
        Dim Rs, Cs As Integer
        Rs = Me.GetRows()
        Cs = Me.GetCols()
        res.SetSize(Rs, Cs)
        res.Position = ip
        For i = 0 To Cs - 1
            res.Columns(i).Width = Me.GetColW(i)
            For j = 0 To Rs - 1
                res.Rows(j).Height = Me.GetRowH(j)
                res.Cells(Rs - j - 1, i).TextString = Me.Cells(i, j).Value
            Next
        Next
        res.GenerateLayout() '!?
        Return res
    End Function

End Class
