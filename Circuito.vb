Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class Circuito

    Public circ As String 'Número do Circuito
    Public desc As String 'Descrição do Circuito (i.e. Tomadas Cozinha)
    Public esquema As String 'Esquemas do Circuito (i.e. F+N+T)
    Public tensao As String 'Tensão do Circuito - 220 ou 220/380
    Public potTotal As Double 'Potência total do Circuito
    Public fases As String 'Fases do Circuito (i.e. R+S+T)
    Public potR As Double 'Potência da fase R
    Public potS As Double 'Potência da fase S
    Public potT As Double 'Potência da fase T
    Public sec As Double 'Seção do cabo
    Public disj As Integer 'Valor do Disjuntor
    Public valorDR As Integer 'Valor do DR
    Public tipoDR As String 'Tipo de DR - DDR ou IDR

    Public Sub New(_circ As String, _desc As String, _potTotal As Double,
                   _fases As String, _sec As Double, _disj As Integer)

        circ = _circ
        desc = _desc
        potTotal = _potTotal
        fases = _fases
        sec = _sec
        disj = _disj

    End Sub

    'Public Sub New(_circ As String, _desc As String, _potTotal As Double,
    '               _fases As String, _sec As Double, _disj As Integer,
    '               _valorDr As Integer, _tipoDr As String)

    '    circ = _circ
    '    desc = _desc
    '    potTotal = _potTotal
    '    fases = _fases
    '    sec = _sec
    '    disj = _disj
    '    valorDr = _valorDr
    '    tipoDr = _tipoDr

    'End Sub

    'Public Sub New(_circ As String, _desc As String, _esquema As String,
    '               _tensao As String, _potTotal As Double, _fases As String,
    '               _potR As Double, _potS As Double, _potT As Double,
    '               _sec As Double, _disj As Integer, _valorDr As Integer, _tipoDr As String)

    '    circ = _circ
    '    desc = _desc
    '    esquema = _esquema
    '    tensao = _tensao
    '    potTotal = _potTotal
    '    fases = _fases
    '    potR = _potR
    '    potS = _potS
    '    potT = _potT
    '    sec = _sec
    '    disj = _disj
    '    valorDr = _valorDr
    '    tipoDr = _tipoDr

    'End Sub

    '------------------------------------------------------------------------------------------------------------------------------------------------

    Public Sub DesenhaCircuito(posicao As Point3d)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()

            'Abre o Block Table para Leitura 
            Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)

            Dim blkID As ObjectId = ObjectId.Null

            Dim blkName As String = "MD - DU Circuito"
            'Checa se o bloco do circuito faz parte do projeto ou não
            If Not bt.Has(blkName) Then 'Se o Bloco não faz parte do projeto, adiciona ele

                'Adicionar Bloco pelo .dwg caso já não faça parte do projeto

            Else 'Se o Bloco faz parte do projeto, utiliza ele

                blkID = bt(blkName)

            End If

            If blkID <> ObjectId.Null Then

                Using btr As BlockTableRecord = trans.GetObject(blkID, OpenMode.ForRead) 'Block Table Record do Bloco do Circuito

                    Using blkRef As New BlockReference(posicao, btr.Id) 'Cria o Block Reference do Bloco do Circuito - Leva em consideração sua posição de inserção

                        Dim curBtr As BlockTableRecord = trans.GetObject(db.CurrentSpaceId, OpenMode.ForWrite)

                        curBtr.AppendEntity(blkRef) 'Insere o Block Reference ao Model Space
                        trans.AddNewlyCreatedDBObject(blkRef, True) 'Confirma a inserção

                        'Verifica se o Bloco possui Atributos
                        If btr.HasAttributeDefinitions Then

                            'Checa Atributo por Atributo
                            For Each objID As ObjectId In btr

                                Dim dbObj As DBObject = trans.GetObject(objID, OpenMode.ForRead)

                                If TypeOf dbObj Is AttributeDefinition Then

                                    Dim attDef As AttributeDefinition = dbObj

                                    If Not attDef.Constant Then

                                        Using attRef As New AttributeReference

                                            attRef.SetAttributeFromBlock(attDef, blkRef.BlockTransform)
                                            attRef.Position = attDef.Position.TransformBy(blkRef.BlockTransform)

                                            Select Case attRef.Tag
                                                Case "CIRC_N_CIRCUITO"
                                                    attRef.TextString = circ.ToString
                                                    Exit Select
                                                Case "CIRC_DESCRIÇÃO"
                                                    attRef.TextString = "(" & desc.ToString & ")"
                                                    Exit Select
                                                Case "CIRC_POTÊNCIA"
                                                    attRef.TextString = "(" + potTotal.ToString + " W)"
                                                    Exit Select
                                                Case "CIRC_FASE"
                                                    attRef.TextString = fases
                                                    Exit Select
                                                Case "CIRC_SEÇÃO"
                                                    attRef.TextString = sec.ToString & " mm²"
                                                    Exit Select
                                                Case "CIRC_DISJUNTOR"
                                                    If desc.Contains("Iluminação") Then
                                                        attRef.TextString = disj.ToString & " A - B"
                                                    Else
                                                        attRef.TextString = disj.ToString & " A - C"
                                                    End If
                                                    Exit Select
                                            End Select

                                            blkRef.AttributeCollection.AppendAttribute(attRef) 'Adiciona o Bloco do Circuito ao Canvas
                                            trans.AddNewlyCreatedDBObject(attRef, True) 'Confirma a Transação

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
