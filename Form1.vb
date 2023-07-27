Imports System
Imports System.IO
Imports System.Text
Public Class Form1
    Dim objBaseClass As ClsBase
    Dim objFileValidate As ClsValidation
    Dim objGetSetINI As ClsShared

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            'Dim ss As StringBuilder = New StringBuilder()
            'Dim ss1 As StringBuilder = New StringBuilder()
            'Dim strdd = "20,00NTRF1111RT131445//MUHABIR MASRAFI"
            'Dim flag As Boolean = False
            'For Each c As Char In strdd

            '    If Char.IsLetter(c) = True Or Char.IsWhiteSpace(c) = True Or (c = Chr(34) And flag = False) Then

            '        Dim ssle As Integer = ss1.Length
            '        Dim traamout As String = strdd.ToString().Substring(0, ssle)
            '        Dim tracode As String = strdd.ToString().Substring(ssle, strdd.ToString().Length - ssle)
            '        Exit For
            '        ss = ss.Append(c)
            '        flag = True
            '    Else
            '        ss1 = ss1.Append(c)
            '    End If


            'Next

            Timer1.Enabled = True
            Timer1.Interval = 1000

            Generate_Setting_File()

        Catch ex As Exception
            Call objBaseClass.Handle_Error(ex, "Form", Err.Number, "Form_Load")

        End Try
    End Sub

    Private Sub Generate_Setting_File()

        Dim strConverterCaption As String = ""
        Dim strSettingsFilePath As String = My.Application.Info.DirectoryPath & "\settings.ini"

        Try
            objGetSetINI = New ClsShared

            '-Genereate Settings.ini File-
            If Not File.Exists(strSettingsFilePath) Then
                '-General Section-
                Call objGetSetINI.SetINISettings("General", "Date", Format(Now, "dd/MM/yyyy"), strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Audit Log", My.Application.Info.DirectoryPath & "\Audit", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Error Log", My.Application.Info.DirectoryPath & "\Error", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Input Folder", My.Application.Info.DirectoryPath & "\Input", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Output Folder", My.Application.Info.DirectoryPath & "\Output", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Reverse Folder", My.Application.Info.DirectoryPath & "\Reverse", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Reverse Output Folder", My.Application.Info.DirectoryPath & "\Reverse Output", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderSuc", My.Application.Info.DirectoryPath & "\Archive\Success", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Archived FolderUnSuc", My.Application.Info.DirectoryPath & "\Archive\Unsuccess", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Temp Folder", My.Application.Info.DirectoryPath & "\Temp", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Report Folder", My.Application.Info.DirectoryPath & "\Report", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Validation", My.Application.Info.DirectoryPath & "\Validation\Validation.xlsx", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Master", My.Application.Info.DirectoryPath & "\Master\Master File Vendor.xls", strSettingsFilePath)
                Call objGetSetINI.SetINISettings("General", "Converter Caption", "MT-940 To Excel File Converter", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "Process Output File Ignoring Invalid Transactions", "N", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("General", "File Counter", "0", strSettingsFilePath)

                Call objGetSetINI.SetINISettings("General", "==", "==", strSettingsFilePath) 'Separator

                '-Client Details Section-

                ''Call objGetSetINI.SetINISettings("Client Details", "Naming Convention With SFTP (Y/N)", "Y", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Name", "Test", strSettingsFilePath)
                ''Call objGetSetINI.SetINISettings("Client Details", "Domain Name", "TEST", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Client Code", "Test", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Client Details", "Input Date Format", "DD/MM/YYYY", strSettingsFilePath)    'Blank By Default
                'Call objGetSetINI.SetINISettings("Client Details", "==", "==", strSettingsFilePath) 'Separator

                'Call objGetSetINI.SetINISettings("Instruction Details", "Run Identification", "UVR3", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Instruction Details", "Paying Company Code", "7000", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Instruction Details", "House Bank", "HDFC", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Instruction Details", "HB Account Id", "CA2", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Instruction Details", "==", "==", strSettingsFilePath) 'Separator

                ' ''-Annexure Details Section-
                ''Call objGetSetINI.SetINISettings("Annexure Details", "No of Annexure Record Line", "1", strSettingsFilePath)
                ''Call objGetSetINI.SetINISettings("Annexure Details", "==", "==", strSettingsFilePath) 'Separator
                'Call objGetSetINI.SetINISettings("Annexure Details", "Annexure Text Link Ref", "0", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Annexure Details", "Address", "Mumbai", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Annexure Details", "Bank Name", "ABC Bank Ltd", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Annexure Details", "Account Number", "123456", strSettingsFilePath)


                'Call objGetSetINI.SetINISettings("Annexure Details", "==", "==", strSettingsFilePath)

                ''-Encryption Section-
                'Call objGetSetINI.SetINISettings("Encryption", "Encryption Required (Y/N)", "N", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Encryption", "Batch File Path", "C:\GenericEncryption_Client\encryptdaemon.bat", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Encryption", "PICKDIR Path", "C:\GenericEncryption_Client\datafiles\clearfiles", strSettingsFilePath)
                'Call objGetSetINI.SetINISettings("Encryption", "DROPDIR Path", "C:\GenericEncryption_Client\datafiles\encfiles", strSettingsFilePath)

            End If

            '-Get Converter Caption from Settings-
            If File.Exists(strSettingsFilePath) Then
                strConverterCaption = objGetSetINI.GetINISettings("General", "Converter Caption", strSettingsFilePath)
                If strConverterCaption <> "" Then
                    Text = strConverterCaption.ToString() & " - Version " & Mid(Application.ProductVersion.ToString(), 1, 3)
                Else
                    MsgBox("Either settings.ini file does not contains the key as [ Converter Caption ] or the key value is blank" & vbCrLf & "Please refer to " & strSettingsFilePath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End
                End If
            End If

        Catch ex As Exception
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error while Generating Settings File")
            End

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Sub

    Private Function GetDateFormatExtensionWise(ByVal strFileExtension As String)

        Dim SysDateFormat As String = System.Globalization.DateTimeFormatInfo.CurrentInfo.ShortDatePattern.ToUpper()
        Dim TmpInputDateFormat As String = ""

        Try
            objGetSetINI = New ClsShared

            If File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                TmpInputDateFormat = objGetSetINI.GetINISettings("Client Details", "Input Date Format", My.Application.Info.DirectoryPath & "\settings.ini")
                If TmpInputDateFormat = "" Then
                    If strFileExtension.ToString().Trim().ToUpper() = ".XLS" Or strFileExtension.ToString().Trim().ToUpper() = ".CSV" Then
                        objGetSetINI.SetINISettings("Client Details", "Input Date Format", SysDateFormat, My.Application.Info.DirectoryPath & "\settings.ini")
                    Else
                        objGetSetINI.SetINISettings("Client Details", "Input Date Format", "DD/MM/YYYY", My.Application.Info.DirectoryPath & "\settings.ini")
                    End If
                End If
            End If

        Catch ex As Exception
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Input Date Format")

        Finally
            If Not objGetSetINI Is Nothing Then
                objGetSetINI.Dispose()
                objGetSetINI = Nothing
            End If

        End Try

    End Function

    Private Sub Conversion_Process()
        Dim objFolderAll As DirectoryInfo

        Try

            If objBaseClass Is Nothing Then
                objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")
            End If

            'objBaseClass = New ClsBase(My.Application.Info.DirectoryPath & "\settings.ini")

            '-Get Settings-
            If GetAllSettings() = True Then
                MsgBox("Either file path is invalid or any key value is left blank in settings.ini file", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error In Settings")
                Exit Sub
            End If

            '-Process Input-
            'Process_Each(TxtFilePath.Text)

            objFolderAll = New DirectoryInfo(strInputFolderPath)
            If objFolderAll.GetFiles.Length = 0 Then
                objFolderAll = Nothing
            Else
                objBaseClass.LogEntry("", False)
                objBaseClass.LogEntry("Process Started for INPUT Files")

                For Each objFileOne As FileInfo In objFolderAll.GetFiles()
                    objBaseClass.isCompleteFileAvailable(objFileOne.FullName)
                    If Mid(objFileOne.FullName, objFileOne.FullName.Length - 3, 4).ToString().ToUpper() <> ".BAK" Then
                        objBaseClass.LogEntry("", False)
                        objBaseClass.LogEntry("INPUT File [ " & objFileOne.Name & " ] -- Started At -- " & Format(Date.Now, "hh:mm:ss"), False)

                        Process_Each(objFileOne.FullName)

                        objFolderAll.Refresh()
                    End If
                Next
            End If


        Catch ex As Exception
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Conversion_Process")

        Finally
            If Not objBaseClass Is Nothing Then
                objBaseClass.Dispose()
                objBaseClass = Nothing
            End If

        End Try

    End Sub

    Private Function GetAllSettings() As Boolean

        Try
            GetAllSettings = False

            If Not File.Exists(My.Application.Info.DirectoryPath & "\settings.ini") Then
                GetAllSettings = True
                MsgBox("Either settings.ini file does not exists or invalid file path" & vbCrLf & My.Application.Info.DirectoryPath & "\settings.ini", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            End If

            '-Audit Folder Path-
            If strAuditFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strAuditFolderPath) Then
                    Directory.CreateDirectory(strAuditFolderPath)
                    If Not Directory.Exists(strAuditFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Audit Log folder. Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", True)
                        End If
                        MsgBox("Invalid path for Audit Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Audit Log ] contains invalid path specification", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                        Exit Function
                    End If
                End If
            End If

            '-Error Folder Path-
            If strErrorFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Error Log folder" & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strErrorFolderPath) Then
                    Directory.CreateDirectory(strErrorFolderPath)
                    If Not Directory.Exists(strErrorFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Error Log folder. Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Error Log folder." & vbCrLf & "Please check settings.ini file, the key as [ Error Log ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Input Folder Path-
            If strInputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Input folder" & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strInputFolderPath) Then
                    Directory.CreateDirectory(strInputFolderPath)
                    If Not Directory.Exists(strInputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Input Folder. Please check [ settings.ini ] file, the key as [ Input Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Input Folder." & vbCrLf & "Please check settings.ini file, the key as [ Input Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            '-Output Folder Path-
            If strOutputFolderPath = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Output folder" & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strOutputFolderPath) Then
                    Directory.CreateDirectory(strOutputFolderPath)
                    If Not Directory.Exists(strOutputFolderPath) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Output Folder. Please check [ settings.ini ] file, the key as [ Output Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                    End If
                End If
            End If

            ' ''-Temp Folder Path-
            'If strTempFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Temp folder" & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strTempFolderPath) Then
            '        Directory.CreateDirectory(strTempFolderPath)
            '        If Not Directory.Exists(strTempFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Temp Folder. Please check [ settings.ini ] file, the key as [ Output Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Temp Folder." & vbCrLf & "Please check settings.ini file, the key as [ Temp Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            ''-Report Folder Path-
            'If strReportFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Report folder" & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strReportFolderPath) Then
            '        Directory.CreateDirectory(strReportFolderPath)
            '        If Not Directory.Exists(strReportFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Report Folder. Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Report Folder." & vbCrLf & "Please check settings.ini file, the key as [ Report Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            ' ''-Reverse Folder Path-
            'If strReverseFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Reverse folder" & vbCrLf & "Please check settings.ini file, the key as [ Reverse Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strReverseFolderPath) Then
            '        Directory.CreateDirectory(strReverseFolderPath)
            '        If Not Directory.Exists(strReverseFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Reverse Folder. Please check settings.ini file, the key as [ Reverse Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Reverse Folder." & vbCrLf & "Please check settings.ini file, the key as [ Reverse Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            ' ''-Reverse Output Folder Path-
            'If strReverseOutputFolderPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Reverse Output folder" & vbCrLf & "Please check settings.ini file, the key as [ Reverse Output Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not Directory.Exists(strReverseOutputFolderPath) Then
            '        Directory.CreateDirectory(strReverseOutputFolderPath)
            '        If Not Directory.Exists(strReverseOutputFolderPath) Then
            '            GetAllSettings = True
            '            If Not objBaseClass Is Nothing Then
            '                objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Reverse Output Folder. Please check settings.ini file, the key as [ Reverse Output Folder ] contains invalid path specification.", True)
            '            End If
            '            MsgBox("Invalid path for Reverse Output Folder." & vbCrLf & "Please check settings.ini file, the key as [ Reverse Output Folder ] contains invalid path specification.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '        End If
            '    End If
            'End If

            '-Archive Successful-
            If strArchivedFolderSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive Suc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive Suc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderSuc) Then
                    Directory.CreateDirectory(strArchivedFolderSuc)
                    If Not Directory.Exists(strArchivedFolderSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive Suc Folder. Please check settings.ini file, the key as [ Archive Suc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive Suc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            '-Archive Unsuccessful-
            If strArchivedFolderUnSuc = "" Then
                GetAllSettings = True
                MsgBox("Path is blank for Archive UnSuc folder" & vbCrLf & "Please check settings.ini file, the key as [ Archive UnSuc Folder ] is either does not exist or left blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
                Exit Function
            Else
                If Not Directory.Exists(strArchivedFolderUnSuc) Then
                    Directory.CreateDirectory(strArchivedFolderUnSuc)
                    If Not Directory.Exists(strArchivedFolderUnSuc) Then
                        GetAllSettings = True
                        If Not objBaseClass Is Nothing Then
                            objBaseClass.LogEntry("Error in settings.ini file, Invalid path for Archive UnSuc Folder. Please check settings.ini file, the key as [ Archive UnSuc Folder ] contains invalid path specification.", True)
                        End If
                        MsgBox("Invalid path for Archive UnSuc folder", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
                    End If
                End If
            End If

            ''-Validation File Path-
            'If strValidationPath = "" Then
            '    GetAllSettings = True
            '    MsgBox("Path is blank for Validation file." & vbCrLf & "Please check settings.ini file, the key as [ Validation ] is either does not exist or left blank.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    Exit Function
            'Else
            '    If Not File.Exists(strValidationPath) Then
            '        GetAllSettings = True
            '        If Not objBaseClass Is Nothing Then
            '            objBaseClass.LogEntry("Error in settings.ini file, Validation file does not exist or invalid file path", True)
            '        End If
            '        MsgBox("Validation file does not exist or invalid file path" & vbCrLf & strValidationPath, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error in settings.ini file")
            '    End If
            'End If

        Catch ex As Exception
            GetAllSettings = True
            MsgBox("Error-" & vbCrLf & Err.Description & "[" & Err.Number & "]", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Error While Getting Log Path from Settings.ini File")

        End Try

    End Function

    Private Sub Process_Each(ByVal StrInputFileName As String)

        Dim StrAns As Int32

        Try
            gstrInputFolder = StrInputFileName.Substring(0, StrInputFileName.LastIndexOf("\"))
            'gstrInputFile = StrInputFileName.Substring(StrInputFileName.LastIndexOf("\"))
            gstrInputFile = Path.GetFileName(StrInputFileName)

            ''-Client Code-
            'If strClientCode.Trim().Length = 0 Then
            '    objBaseClass.LogEntry("Client Code cannot be blank", True)
            '    MsgBox("Client Code cannot be blank", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Settings Error")
            '    Exit Sub
            'End If

            '-Conversion Process-
            'CmdConvert.Enabled = False
            objBaseClass.LogEntry("", False)
            objBaseClass.LogEntry("Process Started")
            objBaseClass.LogEntry("Reading Input File " & gstrInputFile, False)
            'LblStatus.Text = "Reading Input File " & gstrInputFile
            'System.Windows.Forms.Application.DoEvents()

            objFileValidate = New ClsValidation(StrInputFileName, objBaseClass.gstrIniPath)

            If objFileValidate.CheckValidateFile() = True Then
                objBaseClass.LogEntry("Input File Reading Completed Successfully")
                objBaseClass.LogEntry("Input File Validated Successfully")

                'If StrAns = 6 Then
                If objFileValidate.DtInput.Rows.Count > 0 Then
                    objBaseClass.LogEntry("Output File Generation Process Started")

                    If GenerateSAPOutPutFile(objFileValidate.DtInput, gstrInputFile) = True Then       ''Generating Output
                        objBaseClass.LogEntry("Output File Generation process failed due to Error", True)
                        objBaseClass.LogEntry("Output File Generation process failed due to Error", False)
                        objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))
                        'MessageBox.Show("Output File Generation process failed due to Error", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else

                        objBaseClass.FileMove(strInputFolderPath & "\" & gstrInputFile, strArchivedFolderSuc & "\" & gstrInputFile)

                        objBaseClass.LogEntry("Output File" & gstrOutputFile & " is Generated Successfully", False)
                    End If

                Else
                    objBaseClass.LogEntry("No Valid Record present in Input File")
                    'MessageBox.Show("No Valid Record present in Input File", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))

                    'LinkAudit.Visible = True
                End If

            Else
                objBaseClass.LogEntry(gstrInputFile & " is not Valid Input File", False)
                'MessageBox.Show(gstrInputFile & " is not Valid Input File", strClientName, MessageBoxButtons.OK, MessageBoxIcon.Error)
                objBaseClass.FileMove(StrInputFileName, strArchivedFolderUnSuc & "\" & Path.GetFileName(StrInputFileName))

                'LinkAudit.Visible = True
            End If

            '-Process Status-
            If StrAns <> 7 Then
                objBaseClass.LogEntry("Process Completed")
            End If

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, strClientName, "CmdProcess_Click")

        Finally
            '-Error Log Link-
            If blnErrorLog = True Then
                'LinkError.Visible = True
            Else
                'LinkError.Visible = False
            End If

            objBaseClass.ObjectDispose(objFileValidate.DtInput)
            objBaseClass.ObjectDispose(objFileValidate.DtUnSucInput)

            If Not objFileValidate Is Nothing Then
                objFileValidate.Dispose()
                objFileValidate = Nothing
            End If

        End Try

    End Sub
  
    Private Sub Summary_Report()
        Dim strSumFileName As String
        Dim strSumRepName As String

        Try
            strSumFileName = "Summary_" & Path.GetFileNameWithoutExtension(gstrInputFile) & ".txt"
            strSumRepName = strSumFileName

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Summary Report As On [" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")
            objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))

            '-Summary of Input File-
            objBaseClass.WriteSummaryTxt(strSumFileName, "Input File Details")
            objBaseClass.WriteSummaryTxt(strSumFileName, ("Input File Name ").ToString.PadRight(25, " ") & ":  " & gstrInputFile)

            '-Summary of Output File-
            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Output File Details")
            objBaseClass.WriteSummaryTxt(strSumFileName, ("Output File Name ").ToString.PadRight(25, " ") & ":  " & gstrOutputFile)

            '-Summary of Payment Transaction-
            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "Payments Summary")

            objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of Cheque Records").PadRight(40, " ") & ":" & objFileValidate.DtInput.Select("[Payment method]= 'C'").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(True, "C").ToString().PadLeft(20, " "))
            objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of RTGS Records").PadRight(40, " ") & ":" & objFileValidate.DtInput.Select("[Payment Method]= 'R'").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(True, "R").ToString().PadLeft(20, " "))
            objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of NEFT Records").PadRight(40, " ") & ":" & objFileValidate.DtInput.Select("[Payment Method]= 'N'").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(True, "N").ToString().PadLeft(20, " "))

            'objBaseClass.WriteSummaryTxt(strSumFileName, ("Total No of Records").PadRight(40, " ") & ":" & objFileValidate.DtInput.Select("[TXN_NO]<> 0").Length().ToString().PadLeft(15, " ") & StrDup(8, " ") & ("Total Amount").PadRight(25, " ") & ":" & GetPaymentAmount(True, "").ToString().PadLeft(20, " "))
            objBaseClass.WriteSummaryTxt(strSumFileName, StrDup(105, "-"))

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Summary_Report")

        End Try

    End Sub

    Private Sub Payment_Report()
        Dim strSumFileName As String
        Dim strTranRepName As String

        Try
            strSumFileName = "Transaction_Report_" & Path.GetFileNameWithoutExtension(gstrInputFile) & ".csv"
            strTranRepName = strSumFileName

            objBaseClass.WriteSummaryTxt(strSumFileName, "")
            objBaseClass.WriteSummaryTxt(strSumFileName, "[" & Format(Now, "dd-MM-yyyy hh:mm:ss") & "]")

            objBaseClass.WriteSummaryTxt(strSumFileName, "Transaction Report for Input File " & gstrInputFile)
            objBaseClass.WriteSummaryTxt(strSumFileName, "Beneficiary Name,Transaction Type,Instrument Amount,Status,Reason")

            For Each row As DataRow In objFileValidate.DtInput.Select("[SUBTXN_NO]=0")
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name1").ToString, ",", " ") & "," & row("Payment Method").ToString & "," & row("Amount") & ",Successful," & row("REASON").ToString)
            Next

            For Each row As DataRow In objFileValidate.DtUnSucInput.Select("[SUBTXN_NO]=0")
                objBaseClass.WriteSummaryTxt(strSumFileName, Replace(row("Beneficiary Name1").ToString, ",", " ") & "," & row("Payment Method").ToString & "," & row("Amount") & ",UnSuccessful," & row("REASON").ToString)
            Next

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Payment_Report")

        End Try

    End Sub


    Private Function GetPaymentAmount(ByVal IsSuccess As Boolean, ByVal PayType As String) As Double

        Dim DblAmount As Double = 0

        Try
            If IsSuccess = True Then
                'For Each Row As DataRow In objFileValidate.DtInput.Select("[Transaction Type]='" & PayTpye & "'")
                For Each Row As DataRow In objFileValidate.DtInput.Select("[SUBTXN_NO]=0 and [Payment Method]='" & PayType & "'")
                    DblAmount += Val(Row("Amount").ToString())
                Next
            Else
                'For Each Row As DataRow In objFileValidate.DtUnSucInput.Select("[Transaction Type]='" & PayTpye & "'")
                For Each Row As DataRow In objFileValidate.DtUnSucInput.Select("[SUBTXN_NO]=0 and [Payment Method]='" & PayType & "'")
                    DblAmount += Val(Row("Amount").ToString())
                Next
            End If

            GetPaymentAmount = DblAmount
            ''Added by Lakshmi dtd 29-08-12
            GetPaymentAmount = Convert.ToDecimal(DblAmount).ToString("0.00")
            ''-

        Catch ex As Exception
            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "GetPaymentAmount")

        End Try

    End Function

    Private Function Check_Comma(ByVal strTemp) As String
        Try
            If InStr(strTemp, ",") > 0 Then
                Check_Comma = Chr(34) & strTemp & Chr(34) & ","
            Else
                Check_Comma = strTemp & ","
            End If

        Catch ex As Exception
            blnErrorLog = True

            objBaseClass.WriteErrorToTxtFile(Err.Number, Err.Description, "Form", "Check_Comma")

        End Try
    End Function

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Interval = 1000
        Timer1.Enabled = False

        Conversion_Process()

        Timer1.Enabled = True
    End Sub
End Class
