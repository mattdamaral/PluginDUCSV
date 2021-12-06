Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class QuadroDeDistribuicao

    Public circuitos As New List(Of Circuito)()
    Public drs As New List(Of DR)()
    Public desc As String
    Public potTotal As Double
    Public sec As Double
    Public disj As Integer

    Public valoresDR As List(Of String)
    Public tipoDR As String

    Public Sub New(_circuitos As List(Of Circuito), _desc As String)

        circuitos = _circuitos
        desc = _desc
        'CalculaPotTotal()
        'CalculaSecDisj()

    End Sub

    Public Sub New(_circuitos As List(Of Circuito), _valoresDR As List(Of String), _tipoDR As String, _desc As String, _potTotal As Double, _sec As Double, _disj As Integer)

        circuitos = _circuitos
        desc = _desc
        potTotal = _potTotal
        sec = _sec
        disj = _disj

        valoresDR = _valoresDR
        tipoDR = _tipoDR

    End Sub

    'Desenha o Diagrama Unifilar a partir do .csv
    Public Sub DesenhaDUCSV()

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim pr As PromptPointResult = ed.GetPoint("Enter the Diagram insertion point: ")  'Pede ao Usuário o ponto de inserção do Diagrama Unifilar

        If pr.Status = PromptStatus.OK Then 'Se a seleção está OK, continua

            Call TrocaLayer("MD - Diagrama Unifilar") 'Muda a Layer atual

            Dim posUsuario As Point3d = pr.Value 'Point3D da posição inicial, definida pelo Usuário

            For i = 0 To circuitos.Count() - 1

                Dim posAtual As Point3d = New Point3d(pr.Value.X, pr.Value.Y + (i * (-60)), pr.Value.Z) '60 = Distância entre Blocos
                circuitos(i).DesenhaCircuito(posAtual)

            Next

            DesenhaTexto(desc, New Point3d((posUsuario.X - 200), (posUsuario.Y + 65), 0), 7, 15) 'Texto da Descrição
            DesenhaTexto("(" & potTotal.ToString & " W)", New Point3d((posUsuario.X - 195), (posUsuario.Y + 45), 0), 7, 10) 'Texto da Potência Total

            'Desenha o Retângulo que envolve o Diagrama Unifilar
            Dim frame As New List(Of Point2d)
            frame.Add(New Point2d(posUsuario.X - 200, posUsuario.Y + 60))
            frame.Add(New Point2d(posUsuario.X + 125, posUsuario.Y + 60))
            frame.Add(New Point2d(posUsuario.X + 125, posUsuario.Y - (60 * (circuitos.Count - 1) + 200)))
            frame.Add(New Point2d(posUsuario.X - 200, posUsuario.Y - (60 * (circuitos.Count - 1) + 200)))
            frame.Add(New Point2d(posUsuario.X - 200, posUsuario.Y + 60))
            DesenhaPolyline(frame, 0, 0, 0, True, "DASHED")

            DesenhaTerra(New Point3d(posUsuario.X + 75, posUsuario.Y - (60 * (circuitos.Count - 1) + 200), 0)) 'Desenha o Terra do Diagrama Unifilar

            SeparaDRs(posUsuario)
            For i = 0 To drs.Count() - 1
                drs(i).DesenhaDR()
            Next

            Dim barramento As New Barramento(disj, New Point3d(posUsuario.X, posUsuario.Y - (60 * (circuitos.Count() - 1)) - 70, posUsuario.Z))
            barramento.DesenhaBarramento()

        End If

end_DUCSV:

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub DesenhaTerra(posInicial As Point3d)

        DesenhaLinha(posInicial, New Point3d(posInicial.X, posInicial.Y - 20, 0), 3)
        DesenhaLinha(New Point3d(posInicial.X - 15, posInicial.Y - 20, 0), New Point3d(posInicial.X + 15, posInicial.Y - 20, 0), 3)
        DesenhaLinha(New Point3d(posInicial.X - 9, posInicial.Y - 27.5, 0), New Point3d(posInicial.X + 9, posInicial.Y - 27.5, 0), 3)
        DesenhaLinha(New Point3d(posInicial.X - 3, posInicial.Y - 35, 0), New Point3d(posInicial.X + 3, posInicial.Y - 35, 0), 3)

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub SeparaDRs(posUsuario As Point3d)

        Dim drQtd = valoresDR.Count()
        Dim posInicialBarramento As Point3d ' Where the circuits' buses start
        Dim posFinalBarramento As Point3d ' Where the circuits' buses finish

        Dim posInicialEntrada As Point3d
        Dim posFinalEntrada As Point3d

        Dim currentDRValue As String

        ' Defines the circuits that each DR protects
        Dim drProx As String = "-"
        Dim isDrProxNovo As Boolean = True 'Boolean para checar se o próximo DR é diferente do último
        For i = 0 To (drQtd - 1)
            Dim valorDRAtual As String = valoresDR(i)

            If i = 0 Then

                currentDRValue = valorDRAtual
                posInicialBarramento = New Point3d(posUsuario.X, posUsuario.Y + 20, posUsuario.Z)
                isDrProxNovo = False
                If (drQtd - 1) > i Then
                    drProx = valoresDR(i + 1)

                    If valorDRAtual IsNot drProx And drProx IsNot "" Then
                        posFinalBarramento = New Point3d(posUsuario.X, posUsuario.Y - 20, posUsuario.Z)
                        DesenhaLinha(posInicialBarramento, posFinalBarramento, 6)
                        drs.Add(New DR(currentDRValue, tipoDR, New Point3d(posInicialBarramento.X, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)))
                        isDrProxNovo = True

                        If posInicialEntrada = Nothing Then
                            posInicialEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                        Else
                            posFinalEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                        End If

                    End If
                End If
            Else
                If isDrProxNovo = True Then
                    currentDRValue = valorDRAtual
                    posInicialBarramento = New Point3d(posUsuario.X, posUsuario.Y + (i * (-60)) + 20, posUsuario.Z)
                    isDrProxNovo = False
                End If

                If (drQtd - 1) > i Then
                    drProx = valoresDR(i + 1)

                    If valorDRAtual IsNot drProx And drProx IsNot "" Then
                        posFinalBarramento = New Point3d(posUsuario.X, posUsuario.Y + (i * (-60)) - 20, posUsuario.Z)
                        DesenhaLinha(posInicialBarramento, posFinalBarramento, 6)
                        drs.Add(New DR(currentDRValue, tipoDR, New Point3d(posInicialBarramento.X, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)))
                        isDrProxNovo = True

                        If posInicialEntrada = Nothing Then
                            posInicialEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                        Else
                            posFinalEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                        End If

                    Else
                        isDrProxNovo = False
                    End If
                Else
                    posFinalBarramento = New Point3d(posUsuario.X, posUsuario.Y + (i * (-60)) - 20, posUsuario.Z)
                    DesenhaLinha(posInicialBarramento, posFinalBarramento, 6)
                    drs.Add(New DR(currentDRValue, tipoDR, New Point3d(posInicialBarramento.X, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)))

                    If posInicialEntrada = Nothing Then
                        posInicialEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                    Else
                        posFinalEntrada = New Point3d(posInicialBarramento.X - 55, (posInicialBarramento.Y + posFinalBarramento.Y) / 2, posInicialBarramento.Z)
                    End If

                End If
            End If
        Next

        DesenhaLinha(posInicialEntrada, posFinalEntrada, 7)
        DesenhaEntrada(New Point3d(posInicialEntrada.X, (posInicialEntrada.Y + posFinalEntrada.Y) / 2, posInicialEntrada.Z), "??")

    End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub DesenhaEntrada(posEntrada As Point3d, nomeEntrada As String)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            ' Opens the Block table for read
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null

            If Not bt.Has("MD - DU Proteção Entrada") Then
                ' Add block from file
            Else
                blkID = bt("MD - DU Proteção Entrada")
            End If

            ' Create and insert the new block reference
            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead)
                    Using blkRef As New BlockReference(posEntrada, btr.Id)

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef)
                        trans.AddNewlyCreatedDBObject(blkRef, True)

                        ' Verify block table record has attribute definitions associated with it
                        If btr.HasAttributeDefinitions Then

                            ' Add attributes from the block table record
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            Select Case attRef.Tag
                                                Case "DISJUNTOR"
                                                    attRef.TextString = disj & " A - C"
                                                    Exit Select
                                                Case "SEÇÃO"
                                                    attRef.TextString = "[3#" & sec & "(" & sec & ")" & sec & "] mm²"
                                                    Exit Select
                                                Case "NOME_ENTRADA"
                                                    attRef.TextString = nomeEntrada.ToString
                                                    Exit Select
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
