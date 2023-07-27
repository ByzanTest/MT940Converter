Option Explicit On

Module ModGen

    Public gstrInputFolder As String
    Public gstrInputFile As String
    Public gstrOutputFile As String

    '-Standard-
    Public blnErrorLog As Boolean
    Public strSettingClientCode As String
    Public strSettingClientName As String

    Public gstrReverseOutputFile As String

    '-General-
    Public strAuditFolderPath As String
    Public strErrorFolderPath As String
    Public strInputFolderPath As String
    Public strAdviceFolderPath As String
    Public strOutputFolderPath As String
    Public strTempFolderPath As String
    Public strReportFolderPath As String
    Public strValidationPath As String
    Public strMasterFilePath As String
    Public strArchivedFolderSuc As String
    Public strArchivedFolderUnSuc As String
    Public strProceed As String
    Public strBeneReportFolderPath As String

    Public strReverseFolderPath As String
    Public strReverseOutputFolderPath As String

    Public gstrReverseFile As String
    Public gstrReverseFolder As String

    '-Client Details-
    Public strClientCode As String
    Public strClientName As String
    Public strInputDateFormat As String
    Public strDomainID As String
    Public nmConventionSFTP As String = ""

    '-Instruction Details-
    Public strTranType As String
    Public strPrintLoc As String
    Public strBeneMailAddr As String

    Public strRunIdentification As String
    Public strPayingCompanyCode As String
    Public strHouseBank As String
    Public strHBAccount_Id As String

    '-Encryption-
    Public strEncrypt As String
    Public strBatchFilePath As String
    Public strPICKDIRpath As String
    Public strDROPDIRPath As String

    '-Additional Settings Keys-
    Public strPaymentNo As String
    Public No_Annex_Rec As String
    Public intAnnex As Decimal
    Public strAddress As String
    Public strBankName As String
    Public strAccountNo As String

    'Public strRunIdentification As String
    'Public strPayCompanyCode As String
    'Public strHouseBank As String


    Public strDBpath As String
    Public DBCon As New OleDb.OleDbConnection

    Public gstrFileName As String
    Public gstrJSWAccountNo As String



End Module
