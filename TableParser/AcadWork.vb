Imports Autodesk.AutoCAD.Runtime
Imports MyAcAs = Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.DatabaseServices

Public Class AcadWork

    <CommandMethod("TblParse")> _
    Public Sub TblParse()
        Dim acDoc As MyAcAs.Document = MyAcAs.Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = acDoc.Editor
        Dim wMTbl As MyTable = MyTable.CrTbl(acDoc)
        '
        Dim PPResult As PromptPointResult
        PPResult = ed.GetPoint("Точка вставки")
        If PPResult.Status = PromptStatus.OK Then
            Dim nTbl As Table = wMTbl.CrTbl(PPResult.Value)
            Using tr As Transaction = acDoc.Database.TransactionManager.StartTransaction
                Try
                    Dim bt As BlockTable = tr.GetObject(acDoc.Database.BlockTableId, OpenMode.ForRead)
                    Dim btr As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                    btr.AppendEntity(nTbl)
                    tr.AddNewlyCreatedDBObject(nTbl, True)
                    tr.Commit()
                Catch ex As Exception
                    ed.WriteMessage(ex.ToString())
                    tr.Abort()
                End Try
            End Using
        End If
    End Sub

End Class
