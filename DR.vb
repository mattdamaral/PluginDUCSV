Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class DR

    Public valorDR As String
    Public tipoDR As String
    Public posDR As Point3d

    Public Sub New(_valorDR As String, _tipoDR As String, _posDR As Point3d)

        valorDR = _valorDR
        tipoDR = _tipoDR
        posDR = _posDR

    End Sub

    Public Sub DesenhaDR()

        Dim corFiacao As Integer = 7

        If valorDR.Contains("-") Then
            DesenhaLinha(posDR, New Point3d(posDR.X - 55, posDR.Y, posDR.Z), corFiacao)
        Else

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database

            Using trans As Transaction = db.TransactionManager.StartTransaction()

                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

                Dim blkID As ObjectId = ObjectId.Null

                If Not bt.Has("MD - DU DR") Then

                    'Adicionar o Bloco caso não faça parte do Projeto

                Else
                    blkID = bt("MD - DU DR")
                End If

                If blkID <> ObjectId.Null Then

                    Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)

                        Using blkRef As New BlockReference(New Point3d(posDR.X - 45, posDR.Y, posDR.Z), btr.Id)

                            Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                            curBtr.AppendEntity(blkRef)
                            trans.AddNewlyCreatedDBObject(blkRef, True)

                            'Verifica se o Bloco possui Atributos
                            If btr.HasAttributeDefinitions Then

                                'Checa Atributo por Atributo
                                For Each objID As ObjectId In btr

                                    Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                    If TypeOf dbObj Is AttributeDefinition Then

                                        Dim attDef As AttributeDefinition = dbObj

                                        If Not attDef.Constant Then

                                            Using attRef As New AttributeReference

                                                'Desenha a fiação entre a Entrada, o DR e o Barramento
                                                DesenhaLinha(posDR, New Point3d(posDR.X - 10, posDR.Y, posDR.Z), corFiacao)
                                                DesenhaLinha(New Point3d(posDR.X - 45, posDR.Y, posDR.Z), New Point3d(posDR.X - 55, posDR.Y, posDR.Z), corFiacao)

                                                attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                                attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                                attRef.TextString = valorDR.ToString & " A"

                                                'Adiciona o Bloco ao Model Space e Confirma
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

                'Confirma a Transação
                trans.Commit()

            End Using
        End If
    End Sub

End Class
