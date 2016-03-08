Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports RTree

Public Class MyCell
    Public Box As Line 'Размер ячейки
    Public Col As Integer
    Public Row As Integer
    Public Value As String

    Public Sub New()
        Box = Nothing
        Col = 0
        Row = 0
        Value = ""
    End Sub

    Public Sub New(nBox As Line, wCol As Integer, wRow As Integer, nValue As String)
        Box = nBox
        Col = wCol
        Row = wRow
        Value = nValue
    End Sub

    Public Function GetH() As Double
        Return Box.EndPoint.Y - Box.StartPoint.Y
    End Function

    Public Function GetW() As Double
        Return Box.EndPoint.X - Box.StartPoint.X
    End Function

    Public Function GetRectangle() As Rectangle
        Return New Rectangle(Box.StartPoint.X, Box.StartPoint.Y, Box.EndPoint.X, Box.EndPoint.Y, 0, 0)
    End Function

End Class
