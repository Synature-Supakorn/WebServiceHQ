Imports System
Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Web.Configuration
Imports POSMySQL.POSControl
Imports System.Data
Imports System.Drawing
Imports System.IO
Imports MySql.Data
Imports ClsWebserviceInfo
Imports System.Collections.Generic
Imports System.Net
Imports System.Text

<WebService(Namespace:="pRoMiSeSystem")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    Private objCnn As New MySqlClient.MySqlConnection()
    Private getCnn As New CDBUtil()
    Private DBServer As String
    Private DBName As String
    Private DBIPServer As String
    Private DBIPName As String
    Private RegionID As Integer
    Private ManageDataDri As String
    Private CurrentDri As String
    Private ExchangeDirectory As String
    Private FOLDER_ERROR As String = "LogFileWebService"
    Private xLogFileName As String = "ErrorWebService"
    Private strManageDataDirectory As String = WebConfigurationManager.AppSettings("ErrorLogDirectory")

    <WebMethod()> _
    Private Sub GetconnectionAndWebConfiguration()
        objCnn = getCnn.EstablishConnection()

        DBIPServer = WebConfigurationManager.AppSettings("DBName")
        DBIPName = WebConfigurationManager.AppSettings("DBServer")

        RegionID = WebConfigurationManager.AppSettings("RegiouID")
        ExchangeDirectory = WebConfigurationManager.AppSettings("ExchangeDirectory")
        CurrentDri = WebConfigurationManager.AppSettings("ExchangeDirectory")
        ManageDataDri = WebConfigurationManager.AppSettings("ExchangeDirectory")
    End Sub
    <WebMethod()> _
       Private Function CheckStringForInsertInto(ByVal Str As String) As String
        Str = Str.Replace("'", "''")
        Return Str
    End Function
    <WebMethod()> _
    Public Function ExportDataSetToBranch(ByRef dsResultList() As DataSet, ByVal RegionID As Integer, ByRef xResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            Dim ManageData As New pRoMiSe_ManageData_Class.pRoMiSeExportImportDataProcess(DBServer, DBName, RegionID, 1, ManageDataDri, CurrentDri, pRoMiSe_ManageData_Class.ProgramFor.HeadQuarter)
            If ManageData.AutoExportDataSetToBranch(objCnn, RegionID, dsResultList, xResultText) = False Then
                'Write Error Log file
                WriteErrorLogFile("AutoExportDataSetToShopID_" & RegionID, xResultText, xLogFileName)
                Return False
            Else
                Return True
                WriteErrorLogFile("AutoExportDataSetToShopID_" & RegionID, "Successfully.", xLogFileName)
            End If
            objCnn.Close()
        Catch ex As Exception
            'Write Error Log file
            xResultText = ex.ToString()
            objCnn.Close()
            WriteErrorLogFile("AutoExportDataSetToShopID_" & RegionID, ex.ToString, xLogFileName)
            Return False
        End Try

    End Function
    <WebMethod()> _
    Public Function AutoUpdateDataSetToHQ(ByRef dsResult As DataSet, ByVal xResultText As String) As Boolean
        Dim fromShopID As Integer
        Dim dtTransfer As DataTable

        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()

            dtTransfer = dsResult.Tables("transferlogdetailforexport")
            fromShopID = dtTransfer.Rows(0)("fromShopID")

            Dim ManageData As New pRoMiSe_ManageData_Class.pRoMiSeExportImportDataProcess(DBServer, DBName, RegionID, 1, ManageDataDri, CurrentDri, pRoMiSe_ManageData_Class.ProgramFor.HeadQuarter)
            If ManageData.AutoSetDataInDataSetExportToBranchAtHQ(objCnn, dsResult, xResultText) = False Then
                'Write Error Log file
                WriteErrorLogFile("AutoUpdateDataSetToHQFromShopID_" & fromShopID, xResultText, xLogFileName)
                Return False
            Else
                WriteErrorLogFile("AutoUpdateDataSetToHQFromShopID_" & fromShopID, "Successfully.", xLogFileName)
                Return True
            End If
            objCnn.Close()
        Catch ex As Exception
            'Write Error Log file
            xResultText = ex.ToString()
            objCnn.Close()
            WriteErrorLogFile("AutoUpdateDataSetToHQFromShopID_" & fromShopID, ex.ToString, xLogFileName)
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function ImportRewardPointSummaryAtHQ(ByRef dsData As DataSet, ByRef xResultText As String) As Boolean
        Dim fromShopID As Integer
        Dim dtTransfer As DataTable
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()

            dtTransfer = dsData.Tables("transferlogdetailforexport")
            fromShopID = dtTransfer.Rows(0)("fromShopID")

            Dim ManageData As New pRoMiSe_ManageData_Class.pRoMiSeExportImportDataProcess(DBServer, DBName, RegionID, 1, ManageDataDri, CurrentDri, pRoMiSe_ManageData_Class.ProgramFor.HeadQuarter)
            If ManageData.AutoImportDataSetToHQ(objCnn, dsData, xResultText) = False Then
                'Write Error Log file
                WriteErrorLogFile("ImportRewardPointSummaryAtHQFromShopID_" & fromShopID, xResultText, xLogFileName)
                Return False
            Else
                WriteErrorLogFile("ImportRewardPointSummaryAtHQFromShopID_" & fromShopID, "Successfully.", xLogFileName)
                Return True
            End If
            objCnn.Close()
        Catch ex As Exception
            xResultText = ex.ToString
            'Write Error Log file
            objCnn.Close()
            WriteErrorLogFile("ImportRewardPointSummaryAtHQFromShopID_" & fromShopID, ex.ToString, xLogFileName)
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function ImportSummarySaleByDateToHQ(ByRef dsData As DataSet, ByRef xResultText As String) As Boolean
        Dim fromShopID As Integer
        Dim dtTransfer As DataTable
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            dtTransfer = dsData.Tables("transferlogdetailforexport")
            fromShopID = dtTransfer.Rows(0)("fromShopID")

            Dim ManageData As New pRoMiSe_ManageData_Class.pRoMiSeExportImportDataProcess(DBServer, DBName, RegionID, 1, ManageDataDri, CurrentDri, pRoMiSe_ManageData_Class.ProgramFor.HeadQuarter)
            If ManageData.AutoImportDataSetToHQ(objCnn, dsData, xResultText) = False Then
                'Write Error Log file
                WriteErrorLogFile("AutoImportDataSetToHQFromShopID_" & fromShopID, xResultText, xLogFileName)
                Return False
            Else
                WriteErrorLogFile("AutoImportDataSetToHQFromShopID_" & fromShopID, "Successfully.", xLogFileName)
                Return True
            End If
            objCnn.Close()
        Catch ex As Exception
            xResultText = ex.ToString
            objCnn.Close()
            'Write Error Log file
            WriteErrorLogFile("AutoImportDataSetToHQFromShopID", ex.ToString, xLogFileName)
            Return False
        End Try
    End Function
    <WebMethod()> _
  Public Function SearchMember(ByVal searchBy As SearchMemberBy, ByVal paramSearch As String, ByRef memberData As List(Of Member_Data), ByRef dsMemberData As DataSet, ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            Dim starttime As String
            Dim endtime As String
            starttime = Date.Now
            GetconnectionAndWebConfiguration()
            If MembersList(getCnn, objCnn, searchBy, paramSearch, memberData, dsMemberData, strResultText) = True Then
                endtime = Date.Now
                WriteErrorLogFile("SearchMember", "Request:" & starttime & " Response:" & endtime, "LogSearchMember")
                Return True
            Else
                WriteErrorLogFile("SearchMember", strResultText, xLogFileName)
                Return False

            End If
            objCnn.Close()
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            objCnn.Close()
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function GetMember(ByVal searchBy As Integer, ByVal memberCode As String, ByVal memberMobile As String, ByRef memberData As Member_Data, ByRef dsMemberData As DataSet, ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            Dim starttime As String
            Dim endtime As String
            starttime = Date.Now
            GetconnectionAndWebConfiguration()
            If Members(getCnn, objCnn, searchBy, memberCode, memberMobile, memberData, dsMemberData, strResultText) = True Then
                endtime = Date.Now
                WriteErrorLogFile("SearchMember", "Request:" & starttime & " Response:" & endtime, "LogSearchMember")
                Return True
            Else
                WriteErrorLogFile("GetMember", "MemberCode :" & memberCode & " " & strResultText, "GetMember")
                Return False
            End If
            objCnn.Close()
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            WriteErrorLogFile("GetMember", "MemberCode :" & memberCode & " " & ex.Message.ToString(), "GetMember")
            objCnn.Close()
            Return False
        End Try
    End Function
    <WebMethod()> _
   Public Function AddUpdateMembersAtQH(ByVal fromShopID As Integer, ByVal destinationShopID As Integer, ByVal dsData As DataSet, ByRef xResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            Dim ManageData As New pRoMiSe_ManageData_Class.pRoMiSeExportImportDataProcess(DBServer, DBName, RegionID, 1, ManageDataDri, CurrentDri, pRoMiSe_ManageData_Class.ProgramFor.HeadQuarter)
            If ManageData.DirectImportDataSetToDatabase(objCnn, fromShopID, destinationShopID, dsData, xResultText) = False Then
                'Write Error Log file
                WriteErrorLogFile("AddUpdateMembersAtQHFromShopID_" & fromShopID, xResultText, xLogFileName)
                Return False
            Else
                WriteErrorLogFile("AddUpdateMembersAtQHFromShopID_" & fromShopID, "Successfully.", xLogFileName)
                Return True
            End If
            objCnn.Close()
        Catch ex As Exception
            xResultText = ex.ToString
            'Write Error Log file
            objCnn.Close()
            WriteErrorLogFile("AddUpdateMembersAtQHFromShopID_" & fromShopID, ex.ToString, xLogFileName)
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function SearchSummaryPoint(ByVal memberID As Integer, ByRef memberData As SummaryPoint_Data, ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            If SummaryPoint(getCnn, objCnn, memberID, memberData, strResultText) = True Then
                Return True
            Else
                WriteErrorLogFile("SearchSummaryPoint", strResultText, "SearchSummaryPoint")
                Return False
            End If
            objCnn.Close()
        Catch ex As Exception
            strResultText = ex.Message.ToString
            WriteErrorLogFile("SearchSummaryPoint", ex.Message.ToString(), "SearchSummaryPoint")
            objCnn.Close()
            Return False
        End Try
    End Function
    <WebMethod()> _
   Public Function UpdateSoftwareVersion(ByVal ComputerID As Integer, ByVal ProductLevelID As Integer, ByVal IPAddress As String, ByVal FrontVersion As String, ByVal FrontFileDate As String, ByVal FrontUpdateDate As String, ByVal backOfficeVersion As String, ByVal backOfficeFileDate As String, ByVal backOfficeUpdateDate As String, ByVal InvVersion As String, ByVal InvFileDate As String, ByVal InvUpdateDate As String, ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            If ClsWebServiceData.UpdateSoftwareVersionAtHQ(getCnn, objCnn, ComputerID, ProductLevelID, IPAddress, FrontVersion, FrontFileDate, FrontUpdateDate, backOfficeVersion, backOfficeFileDate, backOfficeUpdateDate, InvVersion, InvFileDate, InvUpdateDate, strResultText) = True Then
                Return True
            Else
                WriteErrorLogFile("UpdateSoftwareVersionFromShopID_" & ProductLevelID, strResultText, "UpdateSoftwareVersion")
                Return False
            End If
            objCnn.Close()
        Catch ex As Exception
            strResultText = ex.ToString()
            objCnn.Close()
            WriteErrorLogFile("CheckVersionSoftwareFromShopID_" & ProductLevelID, ex.ToString, "CheckVersionSoftware")
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function GetSoftwareVersion(ByVal programTypeID As Integer, ByRef softwareData As Softwareversion_Data, ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            If SoftwareVersion(getCnn, objCnn, programTypeID, softwareData, strResultText) = True Then
                Return True
            Else
                WriteErrorLogFile("CheckVersionSoftware", strResultText, "CheckVersionSoftware")
                Return False
            End If
            objCnn.Close()
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            objCnn.Close()
            WriteErrorLogFile("CheckVersionSoftware", ex.ToString, "CheckVersionSoftware")
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function ContentLastUpdate(ByVal shopID As Integer, ByVal sectionID As Integer, ByVal limitContent As Integer, ByRef contentData As List(Of News_CategoryData), ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            contentData = News_ListCategory(getCnn, objCnn, shopID, sectionID, limitContent)
            strResultText = ""
            objCnn.Close()
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            objCnn.Close()
            WriteErrorLogFile("ContentLastUpdate", ex.ToString, "ContentLastUpdate")
            Return False
        End Try
    End Function
    <WebMethod()> _
   Public Function ContentSection(ByVal sectionData As List(Of News_SectionData), ByRef strResultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            sectionData = News_ListSection(getCnn, objCnn, -1)
            strResultText = ""
            objCnn.Close()
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            objCnn.Close()
            WriteErrorLogFile("ContentLastUpdate", ex.ToString, "ContentLastUpdate")
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function Member_UpdatePackage(ByVal objPackage As List(Of Packagehistory), ByRef resultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()
            Dim strSQL As String = ""
            strSQL = ClsWebServiceData.Member_GenerateScriptUpdatePackage(objPackage)

            If strSQL <> "" Then
                getCnn.sqlExecute(strSQL, objCnn)
            End If
            objCnn.Close()
            Return True
        Catch ex As Exception
            resultText = ex.Message.ToString()
            WriteErrorLogFile("UpdatePackage", ex.Message.ToString(), "UpdatePackage")
            objCnn.Close()
            Return False
        End Try
    End Function
    <WebMethod()> _
  Public Function Payment_Paybyvoucher(ByRef dsData As DataSet, ByRef resultText As String) As Boolean
        Try
            'Create Connecttion
            GetconnectionAndWebConfiguration()

            Dim objRequest
            Dim URL
            Dim DtData As New DataTable
            Dim StrBD As New StringBuilder
            DtData = dsData.Tables("DtPayByVoucher")
            If DtData.Rows.Count > 0 Then
                For i As Integer = 0 To DtData.Rows.Count - 1
                    If Not IsDBNull(DtData.Rows(i)("couponvoucherno")) Then
                        If DtData.Rows(i)("couponvoucherno") <> "" Then
                            StrBD.Append("UPDATE HistoryOfCreateCouponVoucher SET CouponVoucherAlreadyUse=1 WHERE CreateVoucherNo='" & DtData.Rows(i)("couponvoucherno") & "'")
                        End If
                    End If
                Next
                If StrBD.ToString <> "" Then
                    getCnn.sqlExecute(StrBD.ToString, objCnn)
                End If
            End If
            objRequest = CreateObject("Microsoft.XMLHTTP")
            'Put together the URL link appending the Variables.
            URL = "http://127.0.0.1/Loyaltyservice/SyncVoucher.aspx?action=1"

            'Open the HTTP request and pass the URL to the objRequest object
            objRequest.open("POST", URL, False)

            'Send the HTML Request
            objRequest.Send()

            'Set the object to nothing
            objRequest = Nothing

            objCnn.Close()
            Return True
        Catch ex As Exception
            resultText = ex.Message.ToString()
            WriteErrorLogFile("Payment_Paybyvoucher", ex.Message.ToString(), "Payment_Paybyvoucher")
            objCnn.Close()
            Return False
        End Try
    End Function
    <WebMethod()> _
    Public Function WriteErrorLogFile(ByVal errorFrom As String, ByVal errorString As String, _
        ByVal logFileName As String) As Boolean
        Dim strErrorLogFileName As String
        If System.IO.Directory.Exists(strManageDataDirectory & FOLDER_ERROR & System.IO.Path.DirectorySeparatorChar) = False Then
            System.IO.Directory.CreateDirectory(strManageDataDirectory & FOLDER_ERROR & System.IO.Path.DirectorySeparatorChar)
        End If
        strErrorLogFileName = strManageDataDirectory & FOLDER_ERROR & System.IO.Path.DirectorySeparatorChar & _
                                logFileName & "_" & Format(Now, "yyyyMMdd")
        strErrorLogFileName &= ".txt"
        Try
            'File stream and text stream
            Dim fsWrite As System.IO.FileStream = New System.IO.FileStream(strErrorLogFileName, IO.FileMode.Append, _
                                    IO.FileAccess.Write, IO.FileShare.Write)
            Dim wr As System.IO.StreamWriter = New System.IO.StreamWriter(fsWrite)
            wr.WriteLine("Error Message At " & Format(Now, "dd-MM-yyyy hh:mm:ss"))
            wr.WriteLine(errorFrom & ": ")
            wr.WriteLine(errorString)
            wr.WriteLine("")
            wr.WriteLine("--------------------------------------------------------------------------------")
            wr.Close()
            fsWrite.Close()
        Catch e As Exception
            Return False
        End Try
        Return True
    End Function

End Class
