Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Barramento

    Public tamanhoBarramento As String
    Public posInicial

    Public Sub New(_disj As Integer, _posInicial As Point3d)

        posInicial = _posInicial

        If _disj <> Nothing Then

            Select Case (_disj * 1.3)
                Case < 140
                    tamanhoBarramento = "15x2"
                    Exit Select
                Case < 170
                    tamanhoBarramento = "15x3"
                    Exit Select
                Case < 185
                    tamanhoBarramento = "20x2"
                    Exit Select
                Case < 220
                    tamanhoBarramento = "20x3"
                    Exit Select
                Case < 270
                    tamanhoBarramento = "25x3"
                    Exit Select
                Case < 295
                    tamanhoBarramento = "20x5"
                    Exit Select
                Case < 315
                    tamanhoBarramento = "30x3"
                    Exit Select
                Case < 350
                    tamanhoBarramento = "25x5"
                    Exit Select
                Case < 400
                    tamanhoBarramento = "30x5"
                    Exit Select
                Case < 420
                    tamanhoBarramento = "40x3"
                    Exit Select
                Case < 520
                    tamanhoBarramento = "40x5"
                    Exit Select
                Case < 630
                    tamanhoBarramento = "50x5"
                    Exit Select
                Case < 760
                    tamanhoBarramento = "40x10"
                    Exit Select
                Case < 820
                    tamanhoBarramento = "50x10"
                    Exit Select
                Case < 970
                    tamanhoBarramento = "80x5"
                    Exit Select
                Case < 1060
                    tamanhoBarramento = "60x10"
                    Exit Select
                Case < 1200
                    tamanhoBarramento = "100x5"
                    Exit Select
                Case < 1380
                    tamanhoBarramento = "80x10"
                    Exit Select
                Case < 1700
                    tamanhoBarramento = "100x10"
                    Exit Select
                Case < 2000
                    tamanhoBarramento = "120x10"
                    Exit Select
                Case < 2500
                    tamanhoBarramento = "160x10"
                    Exit Select
                Case < 3000
                    tamanhoBarramento = "200x10"
                    Exit Select

            End Select

        End If

    End Sub

    ' Draws the busbar (dimensions for the phase/neutral/ground bars)
    Public Sub DesenhaBarramento()

        Dim corFiacao As Integer = 7 'White

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        ' Starts a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null ' The ID of the DR block

            ' If there's no block called "MD..." in the block table, get it from the block .dwg, else use it from the block table
            If Not bt.Has("MD - DU Barramento") Then
                ' Should add block from a file path
            Else
                blkID = bt("MD - DU Barramento")
            End If

            ' Creates and inserts the new block reference
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                    Using blkRef As New BlockReference(posInicial, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        ' Verify block table record has attribute definitions associated with it
                        If btr.HasAttributeDefinitions Then
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            'Checks If the attribute's tag is one of the below
                                            Select Case attRef.Tag
                                                Case "F"
                                                    attRef.TextString = "F - " + tamanhoBarramento + " mm"
                                                Case "N"
                                                    attRef.TextString = "N - " + tamanhoBarramento + " mm"
                                                Case "PE"
                                                    attRef.TextString = "PE - " + tamanhoBarramento + " mm"
                                            End Select

                                            ' Add DU block to the block table record and to the transaction
                                            blkRef.AttributeCollection.AppendAttribute(attRef)
                                            trans.AddNewlyCreatedDBObject(attRef, True)

                                        End Using
                                    End If
                                End If
                            Next
                        End If
                    End Using
                End Using
            End If

            trans.Commit()

        End Using
    End Sub

End Class
