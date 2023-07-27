Imports System.IO
'Imports Microsoft.Office.Interop

Module SAPPayment

    Dim objLogCls As New ClsErrLog
    Dim Objbase As ClsBase
    Dim objGetSetINI As ClsShared

    Public Function GenerateSAPOutPutFile(ByRef dtSAP As DataTable, ByVal strFileName As String) As Boolean

        Dim objstrWriter As StreamWriter
        Dim strOutputStream As String

     
        Dim FileCounter As String
        'Dim xlWorkBook As Excel.Workbook
        'Dim xlWorkSheet As Excel.Worksheet
        'Dim misValue As Object = System.Reflection.Missing.Value
        'Dim xlApp As Excel.Application = New Microsoft.Office.Interop.Excel.Application()
        Try
            Objbase = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            objGetSetINI = New ClsShared


            'gstrOutputFile = strFileName

            gstrOutputFile = Path.GetFileNameWithoutExtension(strFileName) & ".csv"
            ' gstrOutputFile = Path.GetFileNameWithoutExtension(strFileName) & ".xlsx"

            objstrWriter = New StreamWriter(strOutputFolderPath & "\" & gstrOutputFile)

            Objbase.LogEntry("[Output File Name] = " & gstrOutputFile)

            Dim RowNo As Integer = 0
            Dim ColNo As Integer = 0
            Dim DrRow As DataRow() = Nothing



            'If xlApp Is Nothing Then
            '    'MessageBox.Show("Excel is not properly installed!!")
            '    'Return
            'End If



            'xlWorkBook = xlApp.Workbooks.Add(misValue)
            'xlWorkSheet = xlWorkBook.Sheets("sheet1")
            strOutputStream = ""

            'For Each column As DataColumn In _dtInput.Columns
            For col As Integer = 0 To dtSAP.Columns.Count - 3
                'ColNo += 1
                'RowNo = 1
                strOutputStream = strOutputStream & Check_Comma(dtSAP.Columns(col).ColumnName)

                '  xlWorkSheet.Cells(RowNo, ColNo) = dtSAP.Columns(col).ColumnName  'column.ColumnName
            Next
            ''For Last Value Contain (,)
            strOutputStream = strOutputStream.Substring(0, strOutputStream.Length - 1)
            objstrWriter.WriteLine(strOutputStream, gstrOutputFile)
           

            'xlWorkSheet.Range("A1", "J1").Font.Bold = True
            strOutputStream = ""
            For Rowsi As Integer = 0 To dtSAP.Rows.Count - 1
                'RowNo += 1
                'ColNo = 1
                strOutputStream = ""
                For Colsi As Integer = 0 To dtSAP.Columns.Count - 3

                    strOutputStream = strOutputStream & Check_Comma(dtSAP.Rows(Rowsi)(Colsi))

                    'xlWorkSheet.Cells(RowNo, ColNo) = "'" & dtSAP.Rows(Rowsi)(Colsi)
                    'ColNo += 1
                Next
                ''For Last Value Contain (,)
                strOutputStream = strOutputStream.Substring(0, strOutputStream.Length - 1)
                objstrWriter.WriteLine(strOutputStream, gstrOutputFile)

            Next
            'xlWorkSheet.Columns.AutoFit()
            'xlWorkBook.SaveAs(strOutputFolderPath & "\" & gstrOutputFile, Excel.XlFileFormat.xlWorkbookDefault, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue)
            'xlWorkBook.Close(True, misValue, misValue)
            'xlApp.Quit()

            'releaseObject(xlWorkSheet)
            'releaseObject(xlWorkBook)
            'releaseObject(xlApp)
            If Not objstrWriter Is Nothing Then
                objstrWriter.Close()
                objstrWriter.Dispose()
            End If



            GenerateSAPOutPutFile = False
        Catch ex As Exception
            GenerateSAPOutPutFile = True
            objLogCls.WriteErrorToTxtFile(Err.Number, Err.Description, "Payment", "GenerateOutPutFile")

        Finally
            'releaseObject(xlWorkSheet)
            'releaseObject(xlWorkBook)
            'releaseObject(xlApp)
          
            If Not objstrWriter Is Nothing Then
                objstrWriter.Close()
                objstrWriter.Dispose()
            End If

        End Try

    End Function

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
    Private Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then
                Check_Comma = Chr(34) & strTemp & Chr(34) & ","
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            blnErrorLog = True  '-Added by Jaiwant dtd 31-03-2011

            objGetSetINI.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Check_Comma")

        End Try
    End Function
 

    Private Function Pad_Length(ByVal strtemp As String, ByVal intLen As Integer) As String
        Try
            Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen).Trim()
            'Pad_Length = Microsoft.VisualBasic.Left(strtemp & StrDup(intLen, " "), intLen)

        Catch ex As Exception
            Call objLogCls.Handle_Error(ex, "SAPPayment", Err.Number, "Pad_Length")

        End Try
    End Function

End Module