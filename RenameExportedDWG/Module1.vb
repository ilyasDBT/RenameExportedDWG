Imports System.Windows.Forms
Imports Inventor
Imports Microsoft.WindowsAPICodePack.Dialogs

Module Module1

    Sub Main()
        'Browse for exported files
        Console.WriteLine("RenameExportedDWG Start")

        Dim OpenFileDlg As New OpenFileDialog

        OpenFileDlg.Title = "Exported DWG"
        OpenFileDlg.FileName = "" ' Default file name
        OpenFileDlg.DefaultExt = ".dwg" ' Default file extension
        OpenFileDlg.Filter = "DWG Files (*.dwg)|*.dwg"
        OpenFileDlg.Multiselect = True
        OpenFileDlg.RestoreDirectory = True
        OpenFileDlg.ShowDialog()
        If OpenFileDlg.FileName = "" Then End

        Dim sourceFolderDlg As New CommonOpenFileDialog
        sourceFolderDlg.IsFolderPicker = True
        sourceFolderDlg.Title = "Source Inventor Drawing Folder"
        If Not sourceFolderDlg.ShowDialog() = CommonFileDialogResult.Ok Then End
        Console.WriteLine("Source Folder: " & sourceFolderDlg.FileName)

        Dim countSuccess As Integer = 0
        Dim countFailed As Integer = 0
        For Each filePath In OpenFileDlg.FileNames
            If RenameExportedDWG(filePath, sourceFolderDlg.FileName) Then
                countSuccess += 1
            Else
                countFailed += 1
            End If
        Next filePath

        WriteDivider()
        Console.WriteLine("Success: " & countSuccess)
        Console.WriteLine("Failed: " & countFailed)
        Console.WriteLine("RenameExportedDWG Finish")
        Console.WriteLine("Press Any Key to Exit")
        Console.ReadKey()
    End Sub

    Function RenameExportedDWG(filePath As String, sourceDirectory As String) As Boolean
        WriteDivider()
        Console.WriteLine("FilePath: " & filePath)
        Dim sourcePath As String = sourceDirectory & "\" & System.IO.Path.GetFileNameWithoutExtension(filePath)
        Console.WriteLine("Source File: " & sourcePath)

        If Not My.Computer.FileSystem.FileExists(sourcePath) Then
            Console.WriteLine("Source File Not Found")
            Return False
        End If

        Dim apprentice As ApprenticeServerComponent = New ApprenticeServerComponent
        Dim apprenticeDoc As ApprenticeServerDocument = apprentice.Open(sourcePath)
        Dim revision As String = apprenticeDoc.PropertySets.Item("Inventor Summary Information").Item("Revision Number").Value
        Console.WriteLine("Revision: " & revision)
        apprenticeDoc.Close()
        apprentice.Close()

        Dim fileName As String = System.IO.Path.GetFileNameWithoutExtension(filePath)
        Dim fileExtension As String = System.IO.Path.GetExtension(filePath)
        Dim newName As String = fileName & "_rev_" & revision & fileExtension

        Try
            My.Computer.FileSystem.RenameFile(filePath, newName)
        Catch
            Console.WriteLine("Failed: " & newName)
            Return False
        End Try
        Console.WriteLine("Success: " & newName)
        Return True
    End Function
    Sub WriteDivider()
        Console.WriteLine("----------------------------------------------------------------------------------------------")
    End Sub


End Module
