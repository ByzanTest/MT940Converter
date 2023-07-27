Imports System
Imports System.Data
Imports System.IO
Imports System.Text

'Imports Microsoft.Office.Interop

Public Class ClsValidation

    Implements IDisposable

    Private ObjBaseClass As ClsBase         ''need to be dispose 
    Private DtValidation As DataTable       ''need to be dispose

    Private DtTemp As DataTable             ''need to be dispose
    Public DtInput As DataTable             ''need to be dispose
    Public DtUnSucInput As DataTable        ''need to be dispose

    Public DtTemp_Reverse As DataTable             ''need to be dispose
    Public DtReverse As DataTable
    Public DtUnSucReverse As DataTable

    Private StrFilePath As String
    Private ValidationPath As String
    Public ErrorMessage As String
    Private DtReverseValidation As DataTable

    Public Sub New(ByVal _strFilePath As String, ByVal _SettINIPath As String)

        StrFilePath = _strFilePath
        'strAdviceFileName = strAdviceFile

        Try
            ObjBaseClass = New ClsBase(_SettINIPath)
            'ValidationPath = ObjBaseClass.GetINISettings("General", "Validation", _SettINIPath)
            'SpCharValidationPath = ObjBaseClass.GetINISettings("General", "Special Character Validation", _SettINIPath)   ''Changed on 20-04-12

            DtInput = New DataTable("SAPInput")
            DefineColumnForSAP(DtInput)
            DtUnSucInput = New DataTable("SAPUnSucInput")
            DefineColumnForSAP(DtUnSucInput)

            'DtReverse = New DataTable("ReverseInput")
            'DefineColumnForReverse(DtReverse)
            'DtUnSucReverse = New DataTable("ReverseUnSucInput")
            'DefineColumnForReverse(DtUnSucReverse)

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "Constructor")

        End Try

    End Sub
  
    Private Sub DefineColumnForSAP(ByRef DtInput As DataTable)

        DtInput.Columns.Add(New DataColumn("Transaction Reference Number"))    '0
        DtInput.Columns.Add(New DataColumn("Related Reference")) '1

        DtInput.Columns.Add(New DataColumn("Account Identification")) '2
        DtInput.Columns.Add(New DataColumn("Statement_Sequence Number"))  '3
        DtInput.Columns.Add(New DataColumn("Intermediate Balance for the date requested")) '4


        DtInput.Columns.Add(New DataColumn("Opening Balance"))  '5
        DtInput.Columns.Add(New DataColumn("Value Date")) '6
        DtInput.Columns.Add(New DataColumn("Entry Date"))    '7
        DtInput.Columns.Add(New DataColumn("Credit_Debit Indicator"))   '8

        DtInput.Columns.Add(New DataColumn("Funds Distribution"))   '9

        DtInput.Columns.Add(New DataColumn("Transaction Amount")) '10

        DtInput.Columns.Add(New DataColumn("Transaction Type")) '11
        DtInput.Columns.Add(New DataColumn("YOUR REF")) '12

        DtInput.Columns.Add(New DataColumn("Bank reference number"))  '13
        DtInput.Columns.Add(New DataColumn("YOUR REF1"))  '14

        DtInput.Columns.Add(New DataColumn("Information to Account Owner"))    '15

        DtInput.Columns.Add(New DataColumn("Intermediate Closing Balance"))    '16


        DtInput.Columns.Add(New DataColumn("Closing Balance"))    '17

        DtInput.Columns.Add(New DataColumn("Closing Available Balance"))    '18
        DtInput.Columns.Add(New DataColumn("Forward Available Balance"))    '19
        DtInput.Columns.Add(New DataColumn("additional information"))    '20



        DtInput.Columns.Add(New DataColumn("TXN_NO"))    '21 
        DtInput.Columns.Add(New DataColumn("SUBTXN_NO"))   '22

        'DtInput.Columns.Add(New DataColumn("TXN_NO")) '24
        'DtInput.Columns.Add(New DataColumn("SUBTXN_NO")) '25


    End Sub

    Public Function CheckValidateFile() As Boolean

        Try
            If Not File.Exists(StrFilePath) Then
                Call ObjBaseClass.Handle_Error(New ApplicationException("Input file path is incorrect or not file found. [" & StrFilePath & "]"), "ClsValidation", -123, "CheckValidateFile")
                CheckValidateFile = False
                Exit Function
            End If

            CheckValidateFile = Validate()
            'If File.Exists(ValidationPath) Then
            '    CheckValidateFile = Validate()
            'Else
            '    'Call ObjBaseClass.Handle_Error(New ApplicationException("Validation file path is incorrect. [" & ValidationPath & "]"), "ClsValidation", -123, "CheckValidateFile")
            'End If

        Catch ex As Exception
            CheckValidateFile = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "CheckValidateFile")
        End Try

    End Function


    Private Function RemoveBlankRow(ByRef _DtTemp As DataTable)
        'To Remove Blank Row Exists in DataTable
        Dim blnRowBlank As Boolean

        Try
            For Each vRow As DataRow In _DtTemp.Rows
                blnRowBlank = True

                For intCol As Int32 = 0 To _DtTemp.Columns.Count - 1
                    If vRow.Item(intCol).ToString().Trim() <> "" Then
                        blnRowBlank = False
                        Exit For
                    End If
                Next

                If blnRowBlank = True Then
                    _DtTemp.Rows(vRow.Table.Rows.IndexOf(vRow)).Delete()
                End If

            Next
            _DtTemp.AcceptChanges()

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "RemoveBlankRow")

        End Try

    End Function

    Private Function Validate() As Boolean

        Validate = False

        Dim DrValidOutputColumn() As DataRow = Nothing

        Dim StrDataRow(22) As String
        Dim ArrDataRow As Object
        Dim InputLineNumber As Int32
        Dim ArrLineData() As Object
        Dim ss As StringBuilder = New StringBuilder()
        Dim ss1 As StringBuilder = New StringBuilder()

        Dim TXN_NO As Int32
        Dim SUBTXN_NO As Int32
        Dim NumberLen As Integer
        Dim TransAmount As String = String.Empty

        Dim TransCodeNAv As String = String.Empty
        Dim TransCode As String = String.Empty
        Dim strFullData As String = String.Empty
        Dim TransactionReferenceNumber As String = String.Empty
        Dim RelatedReference As String = String.Empty
        Dim AccountIdentification As String = String.Empty
        Dim Statement_SequenceNumber As String = String.Empty
        Dim IntermediateBalance As String = String.Empty
        Dim OpeningBalance As String = String.Empty
        Try
            ErrorMessage = ""


            DtTemp = ObjBaseClass.GetDatatable_Text(gstrInputFolder & "\" & gstrInputFile)
            RemoveBlankRow(DtTemp)




            For Each ROW As DataRow In DtTemp.Select("[RecordType]='H'")
                ArrLineData = ROW.ItemArray


                If (ArrLineData(0).ToString().StartsWith(":20:") = True) Then
                    StrDataRow(0) = ArrLineData(0).ToString().Split(":")(2)
                    TransactionReferenceNumber = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":21:") = True) Then
                    StrDataRow(1) = ArrLineData(0).ToString().Split(":")(2)
                    RelatedReference = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":25:") = True) Then
                    StrDataRow(2) = ArrLineData(0).ToString().Split(":")(2)
                    AccountIdentification = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":28C:") = True) Then
                    StrDataRow(3) = ArrLineData(0).ToString().Split(":")(2)
                    Statement_SequenceNumber = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":60M:") = True) Then
                    StrDataRow(4) = ArrLineData(0).ToString().Split(":")(2)
                    IntermediateBalance = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":60F:") = True) Then
                    StrDataRow(5) = ArrLineData(0).ToString().Split(":")(2)
                    OpeningBalance = ArrLineData(0).ToString().Split(":")(2)
                End If
            Next
            'DtInput.Rows.Add(StrDataRow)

            ClearArray(StrDataRow)

            'For Each ROW As DataRow In DtTemp.Select("[RecordType]='P'")
            '    ArrLineData = ROW.ItemArray

            'Next
            ArrLineData = Nothing
            For Each dtRowA As DataRow In DtTemp.Select("SUBTXN_NO='" & 1 & "' And [RecordType]='P' ")
                ClearArray(StrDataRow)
                For Each row As DataRow In DtTemp.Select("TXN_NO='" & dtRowA("TXN_NO") & "'", "SUBTXN_NO")
                    ArrLineData = row.ItemArray

                    If (ArrLineData(0).ToString().StartsWith(":61:") = True) Then

                        strFullData = String.Empty
                        TransAmount = String.Empty
                        TransCodeNAv = String.Empty
                        ss.Clear()
                        ss1.Clear()

                        StrDataRow(0) = TransactionReferenceNumber.ToString()
                        StrDataRow(1) = RelatedReference.ToString()
                        StrDataRow(2) = AccountIdentification.ToString()
                        StrDataRow(3) = Statement_SequenceNumber.ToString()
                        StrDataRow(4) = IntermediateBalance.ToString()
                        StrDataRow(5) = OpeningBalance.ToString()


                        Dim strvalue() As String = ArrLineData(0).ToString().Split(":")

                        StrDataRow(6) = Left(strvalue(2).ToString(), 6)
                        StrDataRow(7) = strvalue(2).ToString().Substring(6, 4)
                        StrDataRow(8) = strvalue(2).ToString().Substring(10, 1)
                        StrDataRow(9) = strvalue(2).ToString().Substring(11, 1)

                        strFullData = strvalue(2).ToString().Substring(12, strvalue(2).ToString().Length - 12)
                        Dim flag As Boolean = False
                        For Each c As Char In strFullData

                            If Char.IsLetter(c) = True Or Char.IsWhiteSpace(c) = True Or (c = Chr(34) And flag = False) Then
                                NumberLen = ss1.Length
                                TransAmount = strFullData.ToString().Substring(0, NumberLen)
                                TransCodeNAv = strFullData.ToString().Substring(NumberLen, strFullData.ToString().Length - NumberLen)
                                Exit For
                                ss = ss.Append(c)
                                flag = True
                            Else
                                ss1 = ss1.Append(c)
                            End If

                        Next
                        StrDataRow(10) = TransAmount.ToString().Replace(",", ".")

                        TransCode = Left(TransCodeNAv, 4)

                        StrDataRow(11) = TransCode.ToString()

                        Dim strdetails As String = String.Empty
                        strdetails = TransCodeNAv.ToString().Substring(4, TransCodeNAv.ToString().Length - 4)
                        StrDataRow(12) = strdetails.ToString().Trim().Split("//")(0)
                        StrDataRow(13) = strdetails.ToString().Trim().Split("//")(2)

                    ElseIf (ArrLineData(0).ToString().StartsWith(":86:") = True) Then

                        Dim strvalue() As String = ArrLineData(0).ToString().Split(":")
                        StrDataRow(15) = strvalue(2).ToString()
                    Else
                        StrDataRow(15) = StrDataRow(15) & " " & ArrLineData(0).ToString()
                    End If


                Next

                DtInput.Rows.Add(StrDataRow)
                ClearArray(StrDataRow)

            Next
            ClearArray(StrDataRow)
            ArrLineData = Nothing
            For Each ROW As DataRow In DtTemp.Select("[RecordType]='F'")
                ArrLineData = ROW.ItemArray
                If (ArrLineData(0).ToString().StartsWith(":62M:") = True) Then
                    StrDataRow(16) = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":62F:") = True) Then
                    StrDataRow(17) = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":64:") = True) Then
                    StrDataRow(18) = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":65:") = True) Then
                    StrDataRow(19) = ArrLineData(0).ToString().Split(":")(2)
                ElseIf (ArrLineData(0).ToString().StartsWith(":86:") = True) Then

                    Dim strvalue() As String = ArrLineData(0).ToString().Split(":")
                    StrDataRow(20) = strvalue(2).ToString()
                Else
                    StrDataRow(20) = StrDataRow(20) & " " & ArrLineData(0).ToString()
                End If
            Next
            DtInput.Rows.Add(StrDataRow)
            Validate = True



        Catch ex As Exception
            Validate = False
            ErrorMessage = ex.Message
            Call ObjBaseClass.Handle_Error(ex, "ClsRBIValidation", Err.Number, "Validate")
        Finally
            DrValidOutputColumn = Nothing
        End Try

    End Function



    Public Function SpCharValidation(ByVal StringValue As String, ByRef _dtSpChar As DataTable) As String

        Dim ArrSpChar(0) As String
        Dim intSpCharRow As Integer
        ''---
        ClearArray(ArrSpChar)
        Array.Resize(ArrSpChar, _dtSpChar.Select.Length)
        intSpCharRow = 0

        For Each SVRow As DataRow In _dtSpChar.Rows
            ArrSpChar(intSpCharRow) = SVRow(0).ToString
            intSpCharRow += 1
        Next

        Dim StrOriginalValue As String = ""
        'Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&", " "} ''Commented by Lakshmi dtd 22-03-2012
        Dim arrSpecialChar() As String = {"'", ";", ".", ",", "<", ">", ":", "?", """", "/", "{", "[", "}", "]", "`", "~", "!", "@", "#", "$", "%", "^", "*", "(", ")", "_", "-", "+", "=", "|", "\", "&"} ''Added by Lakshmi dtd 22-03-2012

        Try
            ''To remove special chars from array which need to ignore.
            For iIChar As Int16 = 0 To ArrSpChar.Length - 1
                For iSChar As Int16 = 0 To arrSpecialChar.Length - 1
                    If ArrSpChar(iIChar) = arrSpecialChar(iSChar) Then
                        arrSpecialChar(iSChar) = Nothing
                    End If
                Next
            Next
            SpCharValidation = ""
            Dim i As Integer
            For i = 0 To arrSpecialChar.Length - 1
                If InStr(StringValue, arrSpecialChar(i), CompareMethod.Binary) <> 0 Then
                    SpCharValidation = SpCharValidation & arrSpecialChar(i)
                End If
            Next

            Return SpCharValidation

        Catch ex As Exception
            blnErrorLog = True

            Call ObjBaseClass.Handle_Error(ex, "ClsRBIValidation", "SpCharValidation")

        End Try
    End Function

    Private Sub AddRowsToDataTable(ByVal pNotValid As Boolean, ByVal Data() As String)
        Try
            If Data Is Nothing Then Exit Sub

            If pNotValid = True Then
                DtUnSucInput.Rows.Add(Data)
            Else
                DtInput.Rows.Add(Data)
            End If


        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "AddRowsToDataTable")
        End Try
    End Sub

    Private Function GetSubstring(ByVal pStrValue As String, ByVal pStartPos As Int16, ByVal pEndPos As Int16) As String

        Try
            If pStartPos = 0 And pEndPos = 0 Then
                GetSubstring = ""
            Else
                pStartPos = pStartPos - 1
                If pStartPos >= pEndPos Then
                    GetSubstring = "~ERROR~"
                Else
                    ''GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    If Len(Mid(pStrValue, pStartPos + 1, Len(pStrValue))) < (pEndPos - pStartPos) Then
                        GetSubstring = Mid(pStrValue, pStartPos + 1, pEndPos - pStartPos)
                    Else
                        GetSubstring = pStrValue.Substring(pStartPos, pEndPos - pStartPos)
                    End If
                End If
            End If

        Catch ex As Exception
            GetSubstring = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetSubstring")

        End Try

    End Function

    Private Function GetValueFormArray(ByRef pArray() As Object, ByVal pPosition As Int16) As String

        Try

            If pArray.Length >= pPosition Then
                GetValueFormArray = pArray(pPosition - 1).ToString()
                'GetValueFormArray = pArray(pPosition).ToString()
            Else
                ErrorMessage = "~ERROR~"
            End If

        Catch ex As Exception
            ErrorMessage = "~ERROR~"
            Call ObjBaseClass.Handle_Error(ex, "ClsValidation", Err.Number, "GetValueFormArray")

        End Try

    End Function



    Private Sub ClearArray(ByRef ArrRow() As String)

        Try
            For i As Integer = 0 To ArrRow.Length - 1
                ArrRow(i) = ""
            Next

        Catch ex As Exception

        End Try

    End Sub

    Public Function GetColumValue(ByVal strString As String, ByVal intStart As Integer, ByVal intEnd As Integer)

        Try

            intStart = intStart - 1
            GetColumValue = strString.Substring(intStart, intEnd - intStart).Trim()

        Catch ex As Exception
            GetColumValue = ""
        End Try

    End Function

    Public Function RemoveJunk(ByVal sText As String) As String
        ''-To remove Junk Characters-
        Try
            ''PURPOSE: To return only the alpha chars A-Z or a-z or 0-9 and special chars in a string and ignore junk chars.
            Dim iTextLen As Integer = Len(sText)
            Dim n As Integer
            Dim sChar As String = ""

            If sText <> "" Then
                For n = 1 To iTextLen
                    sChar = Mid(sText, n, 1)
                    If IsAlpha(sChar) Then
                        RemoveJunk = RemoveJunk + sChar
                    End If
                Next
            End If

        Catch ex As Exception
            Call ObjBaseClass.Handle_Error(ex, "ClsRBIValidation", "RemoveJunk")

        End Try

    End Function

    Private Function IsAlpha(ByVal sChr As String) As Boolean
        '-To remove Junk Characters-

        IsAlpha = sChr Like "[A-Z]" Or sChr Like "[a-z]" Or sChr Like "[0-9]" _
        Or sChr Like "[.]" Or sChr Like "[,]" Or sChr Like "[;]" Or sChr Like "[:]" _
        Or sChr Like "[<]" Or sChr Like "[>]" Or sChr Like "[?]" Or sChr Like "[/]" _
        Or sChr Like "[']" Or sChr Like "[""]" Or sChr Like "[|]" Or sChr Like "[\]" _
        Or sChr Like "[{]" Or sChr Like "[[]" Or sChr Like "[}]" Or sChr Like "[]]" _
        Or sChr Like "[+]" Or sChr Like "[=]" Or sChr Like "[_]" Or sChr Like "[-]" _
        Or sChr Like "[(]" Or sChr Like "[)]" Or sChr Like "[*]" Or sChr Like "[&]" _
        Or sChr Like "[^]" Or sChr Like "[%]" Or sChr Like "[$]" Or sChr Like "[#]" _
        Or sChr Like "[@]" Or sChr Like "[!]" Or sChr Like "[`]" Or sChr Like "[~]" Or sChr Like "[ ]"

    End Function


    'Public Sub SaveFileinexcelformat(ByVal sfilepath As String)
    '    Dim xl As New Excel.Application
    '    Dim strFile As String
    '    Try

    '        xl.DisplayAlerts = False
    '        xl.Workbooks.Open(sfilepath)
    '        Application.DoEvents()

    '        strFile = Path.GetExtension(sfilepath)
    '        If strFile.ToUpper() = ".XLS" Then
    '            xl.Workbooks(1).SaveAs(sfilepath, Excel.XlFileFormat.xlWorkbookNormal)
    '        ElseIf strFile.ToUpper() = ".XLSX" Then
    '            xl.Workbooks(1).SaveAs(sfilepath, Excel.XlFileFormat.xlOpenXMLWorkbook)
    '        Else
    '            xl.Workbooks(1).SaveAs(sfilepath, Excel.XlFileFormat.xlWorkbookNormal)
    '        End If
    '        'Application.DoEvents()
    '        xl.Workbooks.Close()
    '        'Application.DoEvents()
    '        xl.Quit()
    '        xl = Nothing
    '    Catch ex As Exception
    '        xl.Workbooks.Close()
    '        xl.Quit()
    '        xl = Nothing
    '        Call ObjBaseClass.Handle_Error(ex, "ClsValidation", "SaveFileinexcelformat")
    '    End Try
    'End Sub

#Region " IDisposable Support "
    Public Sub Dispose() Implements IDisposable.Dispose

        If Not ObjBaseClass Is Nothing Then ObjBaseClass.Dispose()
        If Not DtValidation Is Nothing Then DtValidation.Dispose()
        If Not DtInput Is Nothing Then DtInput.Dispose()
        If Not DtUnSucInput Is Nothing Then DtUnSucInput.Dispose()
        If Not DtTemp Is Nothing Then DtTemp.Dispose()

        ObjBaseClass = Nothing
        DtValidation = Nothing
        DtInput = Nothing
        DtUnSucInput = Nothing
        DtTemp = Nothing

        GC.SuppressFinalize(Me)

    End Sub
#End Region

End Class
