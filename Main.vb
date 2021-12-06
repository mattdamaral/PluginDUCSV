Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Windows
Imports System.IO
Imports System
Imports System.Windows

Public Module Main

    'Desenha o Diagrama Unifilar a partir de um Quadro de Cargas armazenado em um arquivo .CSV
    <CommandMethod("DUCSV")>
    Public Sub DUCSV()

        Dim ofd As New OpenFileDialog("Selecione o Arquivo .csv Contendo o Quadro de Cargas", "", "csv", ".CSV to AutoCAD",
                                      OpenFileDialog.OpenFileDialogFlags.DoNotTransferRemoteFiles) 'Configura o Form do Explorador de Arquivos
        Dim dialogResult As System.Windows.Forms.DialogResult = ofd.ShowDialog() 'Abre o Explorador de Arquivos
        Dim csvPath As String = ofd.Filename
        'MsgBox(csvPath) 'Mostra o endereço do Arquivo
        Dim csvContent As String = File.ReadAllText(ofd.Filename) 'Abre o Arquivo e mostra seu conteúdo
        'MsgBox(csvContent) 'Mostra o conteúdo do Arquivo

        Dim inputRecord As String = Nothing
        Dim inReader As StreamReader = File.OpenText(csvPath)
        inputRecord = inReader.ReadLine
        Dim lastRow As Integer = inputRecord.Length
        Dim currentRow As Integer = 0

        Dim circuitos As New List(Of Circuito)() 'Lista de Circuitos que compõem o Quadro de Distribuição
        Dim valoresDR As New List(Of String)() 'Lista de Circuitos que compõem o Quadro de Distribuição
        Dim descQD As String = "" 'Descrição do Quadro de Distribuição (i.e. QD54 - Apartamento Tipo Final 01)
        Dim tipoDR As String 'DDR ou IDR

        While inputRecord IsNot Nothing

            If inputRecord.Contains(",") Then 'Checa se o arquivo .csv contém vírgulas

                currentRow += 1
                Dim nColumns As Integer = 12 'Número de colunas da tabela
                Dim currentRowData(nColumns) As String
                currentRowData = inputRecord.Split(CChar(","))

                If currentRowData(0).Contains("Total") Or currentRowData(0).Contains("TOTAL") Then 'Total

                    Dim potTotal As Double
                    Double.TryParse(currentRowData(4), potTotal)
                    Dim sec As Double
                    Double.TryParse(currentRowData(9), sec)
                    Dim disj As Integer
                    Integer.TryParse(currentRowData(10), disj)
                    Dim qd As New QuadroDeDistribuicao(circuitos, valoresDR, tipoDR, descQD, potTotal, sec, disj)
                    qd.DesenhaDUCSV()

                ElseIf currentRow = 1 Then 'Título do Quadro de Cargas - Descrição

                    descQD = currentRowData(0).ToString

                ElseIf currentRow = 2 Then 'Descrição das Colunas - Tipo de DR

                    If currentRowData(11).Contains("DDR") Then
                        tipoDR = "DDR"
                    ElseIf currentRowData(11).Contains("IDR") Then
                        tipoDR = "IDR"
                    Else
                        tipoDR = "DDR"
                    End If

                ElseIf currentRow > 2 Then 'Circuitos

                    Dim circ As String = currentRowData(0)
                    Dim desc As String = currentRowData(1)
                    Dim potTotal As Double
                    Double.TryParse(currentRowData(4), potTotal)
                    Dim fases As String = currentRowData(5)
                    Dim sec As Double
                    Double.TryParse(currentRowData(9), sec)
                    Dim disj As Integer
                    Integer.TryParse(currentRowData(10), disj)
                    circuitos.Add(New Circuito(circ, desc, potTotal, fases, sec, disj))

                    Dim valorDR As String = currentRowData(11)
                    valoresDR.Add(valorDR)

                End If

                inputRecord = inReader.ReadLine() 'Lê a próxima linha

            End If
        End While
    End Sub

End Module
