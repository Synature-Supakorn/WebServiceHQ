Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data
Imports System.IO
Imports MySql.Data.MySqlClient
Imports POSMySQL.POSControl
Imports System.Web.Configuration
Imports POSTypeClass
Imports System.Xml
Imports FrontUtil
Imports pRoMiSeUtil
<WebService(Namespace:="pRoMiSeService")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.None)> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class Service
    Inherits System.Web.Services.WebService
    Dim objCnn As New MySqlConnection()
    Dim objDB As New CDBUtil()

    Dim ShopID As Integer
    Dim ComputerID As Integer
    Dim noOfCustomer As Integer

    Private ExchangeDirectory As String
    Private FOLDER_ERROR As String = "LogFileWebService"
    Private xLogFileName As String = "msgWebService"
    Private strManageDataDirectory As String = WebConfigurationManager.AppSettings("ErrorLogDirectory")
    Dim i, j, n, m As Integer

    Private Sub EstablisConnection()
        objCnn = objDB.EstablishConnection()
    End Sub
    Private Sub CloseConnention()
        objCnn.Close()
    End Sub
    Private Sub LoadConfiguration()
        '**************************************************************************************************** 
        ' Sub Function Name : LoadConfiguration 
        ' Author            : Supakorn Aonwichean 
        ' Description       : Load configuration 
        '**************************************************************************************************** 
        ShopID = WebConfigurationManager.AppSettings("ShopID")
        ComputerID = WebConfigurationManager.AppSettings("ComputerID")
        noOfCustomer = WebConfigurationManager.AppSettings("noOfCustomer")
    End Sub
    '   <WebMethod()> _
    ' Public Function AddOrderForSpotOn(ByVal dtOrder As String, ByVal MemberCode As String) As DataSet

    '       Dim dt1 As New DataTable
    '       Dim dt2 As New DataTable
    '       dt1.TableName = "Queuing"
    '       dt2.TableName = "Member"
    '       Dim strResult As String = ""
    '       Dim dr1 As DataRow
    '       Dim dr2 As DataRow
    '       Dim ds As New DataSet
    '       Dim dtOrderData As New DataTable
    '       Dim Path As String = Context.Request.PhysicalApplicationPath

    '       Dim strOrder As String = ""
    '       Dim strMember As String = MemberCode
    '       Dim strLog As String = ""

    '       dtOrderData.Columns.Add("ProductCode")
    '       dtOrderData.Columns.Add("Amount")


    '       dt1.Columns.Add("QueueNumber")
    '       dt2.Columns.Add("MemberCode")
    '       dt2.Columns.Add("MemberName")
    '       dt2.Columns.Add("PreviousPrepaidAmount")
    '       dt2.Columns.Add("CurrentPrepaidAmount")

    '       Try

    '           'Load XML OrderData

    '           Dim strXML As String = "<NewDataSet>" + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + "</NewDataSet>"
    '           Dim xmlDoc As New XmlDocument()
    '           xmlDoc.LoadXml(strXML)

    '           Dim root As XmlNode = xmlDoc.FirstChild.ChildNodes(0)
    '           For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("Order")
    '               Dim Dr As DataRow = dtOrderData.NewRow()
    '               For i As Integer = 0 To xmlNode.ChildNodes.Count - 1
    '                   Dr(i) = xmlNode.ChildNodes(i).InnerText
    '                   strOrder += "," + xmlNode.ChildNodes(i).InnerText
    '               Next
    '               dtOrderData.Rows.Add(Dr)
    '           Next

    '           'LogOrder
    '           strLog = "Order : " + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + " MemberCode : " + MemberCode
    '           WriteErrorLogFile("Order", strLog, "KDS")


    '           'Add Transaction Order

    '           GetconnectionAndWebConfiguration()
    '           Path += "pRoMiSeFrontRes.xml"
    '           FrontUtil.clsUtilData.SetAllVariableForWebService(Path, objDB, objCnn, ShopID, ComputerID, 1)
    '           Dim KDSValues As New Data_KDSDetailForWebService
    '           Dim KDS As New POSMainManageTransaction.POSMainManageTransaction(objDB, objCnn)
    '           If KDS.KDSForWebService_AddTransactionAndOrder(ShopID, ComputerID, MemberCode, noOfCustomer, dtOrderData, KDSValues, strResult) = True Then
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = KDSValues.QueueNo
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = KDSValues.MemberID
    '               dr2("MemberName") = KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName
    '               dr2("PreviousPrepaidAmount") = Format(KDSValues.PreviousPrepaidAmount, "#,##0.00")
    '               dr2("CurrentPrepaidAmount") = Format(KDSValues.CurrentPrepaidAmount, "#,##0.00")
    '               dt2.Rows.Add(dr2)


    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)

    '               WriteErrorLogFile("msgKDS", "Succeede", "KDS")
    '               WriteErrorLogFile("msgKDSReturnData", "QueueNumber : " & KDSValues.QueueNo & " Member :" & KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName, "KDS")

    '           Else
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = -1
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = 0
    '               dr2("MemberName") = strResult
    '               dr2("PreviousPrepaidAmount") = 0.0
    '               dr2("CurrentPrepaidAmount") = 0.0
    '               dt2.Rows.Add(dr2)

    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)
    '               WriteErrorLogFile("msgKDS", strResult, "KDS")
    '           End If
    '           objCnn.Close()
    '       Catch ex As Exception
    '           dr1 = dt1.NewRow
    '           dr1("QueueNumber") = -1
    '           dt1.Rows.Add(dr1)

    '           dr2 = dt2.NewRow
    '           dr2("MemberCode") = 0
    '           dr2("MemberName") = ex.ToString
    '           dr2("PreviousPrepaidAmount") = 0.0
    '           dr2("CurrentPrepaidAmount") = 0.0
    '           dt2.Rows.Add(dr2)
    '           objCnn.Close()
    '           ds.Tables.Add(dt1)
    '           ds.Tables.Add(dt2)
    '           WriteErrorLogFile("msgKDS", ex.ToString, "KDS")
    '       End Try

    '       Return ds
    '   End Function
    '   <WebMethod()> _
    '   Public Function AddOrder(ByVal dtOrder As String, ByVal MemberCode As String) As DataSet

    '       Dim dt1 As New DataTable
    '       Dim dt2 As New DataTable
    '       dt1.TableName = "Queuing"
    '       dt2.TableName = "Member"
    '       Dim strResult As String = ""
    '       Dim dr1 As DataRow
    '       Dim dr2 As DataRow
    '       Dim ds As New DataSet
    '       Dim dtOrderData As New DataTable
    '       Dim Path As String = Context.Request.PhysicalApplicationPath

    '       Dim strOrder As String = ""
    '       Dim strMember As String = MemberCode
    '       Dim strLog As String = ""

    '       dtOrderData.Columns.Add("ProductCode")
    '       dtOrderData.Columns.Add("Amount")


    '       dt1.Columns.Add("QueueNumber")
    '       dt2.Columns.Add("MemberCode")
    '       dt2.Columns.Add("MemberName")
    '       dt2.Columns.Add("PreviousPrepaidAmount")
    '       dt2.Columns.Add("CurrentPrepaidAmount")

    '       Try

    '           'Load XML OrderData

    '           Dim strXML As String = "<NewDataSet>" + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + "</NewDataSet>"
    '           Dim xmlDoc As New XmlDocument()
    '           xmlDoc.LoadXml(strXML)

    '           Dim root As XmlNode = xmlDoc.FirstChild.ChildNodes(0)
    '           For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("Order")
    '               Dim Dr As DataRow = dtOrderData.NewRow()
    '               For i As Integer = 0 To xmlNode.ChildNodes.Count - 1
    '                   Dr(i) = xmlNode.ChildNodes(i).InnerText
    '                   strOrder += "," + xmlNode.ChildNodes(i).InnerText
    '               Next
    '               dtOrderData.Rows.Add(Dr)
    '           Next

    '           'LogOrder
    '           strLog = "Order : " + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + " MemberCode : " + MemberCode
    '           WriteErrorLogFile("Order", strLog, "KDS")


    '           'Add Transaction Order

    '           GetconnectionAndWebConfiguration()
    '           Path += "pRoMiSeFrontRes.xml"
    '           FrontUtil.clsUtilData.SetAllVariableForWebService(Path, objDB, objCnn, ShopID, ComputerID, 1)
    '           Dim KDSValues As New Data_KDSDetailForWebService
    '           Dim KDS As New POSMainManageTransaction.POSMainManageTransaction(objDB, objCnn)
    '           If KDS.KDSForWebService_AddTransactionAndOrder(ShopID, ComputerID, MemberCode, noOfCustomer, dtOrderData, KDSValues, strResult) = True Then
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = KDSValues.QueueNo
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = KDSValues.MemberID
    '               dr2("MemberName") = KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName
    '               dr2("PreviousPrepaidAmount") = KDSValues.PreviousPrepaidAmount
    '               dr2("CurrentPrepaidAmount") = KDSValues.CurrentPrepaidAmount
    '               dt2.Rows.Add(dr2)


    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)

    '               WriteErrorLogFile("msgKDS", "Succeede", "KDS")
    '               WriteErrorLogFile("msgKDSReturnData", "QueueNumber : " & KDSValues.QueueNo & " Member :" & KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName, "KDS")

    '           Else
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = -1
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = 0
    '               dr2("MemberName") = strResult
    '               dr2("PreviousPrepaidAmount") = 0
    '               dr2("CurrentPrepaidAmount") = 0
    '               dt2.Rows.Add(dr2)

    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)
    '               WriteErrorLogFile("msgKDS", strResult, "KDS")
    '           End If
    '           objCnn.Close()
    '       Catch ex As Exception
    '           dr1 = dt1.NewRow
    '           dr1("QueueNumber") = -1
    '           dt1.Rows.Add(dr1)

    '           dr2 = dt2.NewRow
    '           dr2("MemberCode") = 0
    '           dr2("MemberName") = ex.ToString
    '           dr2("PreviousPrepaidAmount") = 0
    '           dr2("CurrentPrepaidAmount") = 0
    '           dt2.Rows.Add(dr2)
    '           objCnn.Close()
    '           ds.Tables.Add(dt1)
    '           ds.Tables.Add(dt2)
    '           WriteErrorLogFile("msgKDS", ex.ToString, "KDS")
    '       End Try

    '       Return ds
    '   End Function

    '   <WebMethod(MessageName:="AddOrderDelivery")> _
    'Public Function AddOrder(ByVal dtOrder As String, ByVal MemberCode As String, ByVal saleMode As Integer) As DataSet

    '       Dim dt1 As New DataTable
    '       Dim dt2 As New DataTable
    '       dt1.TableName = "Queuing"
    '       dt2.TableName = "Member"
    '       Dim strResult As String = ""
    '       Dim dr1 As DataRow
    '       Dim dr2 As DataRow
    '       Dim ds As New DataSet
    '       Dim dtOrderData As New DataTable
    '       Dim Path As String = Context.Request.PhysicalApplicationPath

    '       Dim strOrder As String = ""
    '       Dim strMember As String = MemberCode
    '       Dim strLog As String = ""

    '       dtOrderData.Columns.Add("ProductCode")
    '       dtOrderData.Columns.Add("Amount")


    '       dt1.Columns.Add("QueueNumber")
    '       dt2.Columns.Add("MemberCode")
    '       dt2.Columns.Add("MemberName")
    '       dt2.Columns.Add("PreviousPrepaidAmount")
    '       dt2.Columns.Add("CurrentPrepaidAmount")
    '       dt2.Columns.Add("NumberReference")

    '       Try

    '           'Load XML OrderData

    '           Dim strXML As String = "<NewDataSet>" + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + "</NewDataSet>"
    '           Dim xmlDoc As New XmlDocument()
    '           xmlDoc.LoadXml(strXML)

    '           Dim root As XmlNode = xmlDoc.FirstChild.ChildNodes(0)
    '           For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("Order")
    '               Dim Dr As DataRow = dtOrderData.NewRow()
    '               For i As Integer = 0 To xmlNode.ChildNodes.Count - 1
    '                   Dr(i) = xmlNode.ChildNodes(i).InnerText
    '                   strOrder += "," + xmlNode.ChildNodes(i).InnerText
    '               Next
    '               dtOrderData.Rows.Add(Dr)
    '           Next

    '           'LogOrder
    '           strLog = "Order : " + dtOrder.Replace("&lt;", "<").Replace("&gt;", ">") + " MemberCode : " + MemberCode
    '           WriteErrorLogFile("Order", strLog, "KDS_SaleMode")


    '           'Add Transaction Order

    '           GetconnectionAndWebConfiguration()
    '           Path += "pRoMiSeFrontRes.xml"
    '           FrontUtil.clsUtilData.SetAllVariableForWebService(Path, objDB, objCnn, ShopID, ComputerID, 1)
    '           Dim KDSValues As New Data_KDSDetailForWebService
    '           Dim KDS As New POSMainManageTransaction.POSMainManageTransaction(objDB, objCnn)
    '           If KDS.KDSForWebService_AddTransactionAndOrder(ShopID, ComputerID, MemberCode, noOfCustomer, saleMode, dtOrderData, KDSValues, strResult) = True Then
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = KDSValues.QueueNo
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = KDSValues.MemberID
    '               dr2("MemberName") = KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName
    '               dr2("PreviousPrepaidAmount") = KDSValues.PreviousPrepaidAmount
    '               dr2("CurrentPrepaidAmount") = KDSValues.CurrentPrepaidAmount
    '               dr2("NumberReference") = KDSValues.ReferenceNo
    '               dt2.Rows.Add(dr2)

    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)

    '               WriteErrorLogFile("msgKDS", "Succeede", "KDS_SaleMode")
    '               WriteErrorLogFile("msgKDSReturnData", "QueueNumber : " & KDSValues.QueueNo & " Member :" & KDSValues.MemberFirstName & "  " & KDSValues.MemberLastName, "KDS")

    '           Else
    '               dr1 = dt1.NewRow
    '               dr1("QueueNumber") = -1
    '               dt1.Rows.Add(dr1)

    '               dr2 = dt2.NewRow
    '               dr2("MemberCode") = 0
    '               dr2("MemberName") = strResult
    '               dr2("PreviousPrepaidAmount") = 0
    '               dr2("CurrentPrepaidAmount") = 0
    '               dr2("CurrentPrepaidAmount") = 0
    '               dr2("NumberReference") = 0
    '               dt2.Rows.Add(dr2)

    '               ds.Tables.Add(dt1)
    '               ds.Tables.Add(dt2)
    '               WriteErrorLogFile("msgKDS", strResult, "KDS_SaleMode")
    '           End If
    '           objCnn.Close()
    '       Catch ex As Exception
    '           dr1 = dt1.NewRow
    '           dr1("QueueNumber") = -1
    '           dt1.Rows.Add(dr1)

    '           dr2 = dt2.NewRow
    '           dr2("MemberCode") = 0
    '           dr2("MemberName") = ex.ToString
    '           dr2("PreviousPrepaidAmount") = 0
    '           dr2("CurrentPrepaidAmount") = 0
    '           dr2("CurrentPrepaidAmount") = 0
    '           dr2("NumberReference") = 0
    '           dt2.Rows.Add(dr2)
    '           objCnn.Close()
    '           ds.Tables.Add(dt1)
    '           ds.Tables.Add(dt2)
    '           WriteErrorLogFile("msgKDS", ex.ToString, "KDS_SaleMode")
    '       End Try

    '       Return ds
    '   End Function

    '   <WebMethod()> _
    '   Public Function AddSurvey(ByVal ListSurver() As String, ByRef strResultText As String) As Boolean
    '       Try
    '           Dim strVale() As String
    '           Dim i As Integer
    '           Dim Syntax As String = ""
    '           Dim svDatatime As String = pRoMiSeUtil.pRoMiSeUtil.FormatDateTimeForMySQL(Date.Now())
    '           GetconnectionAndWebConfiguration()

    '           If ListSurver.Length >= 0 Then
    '               For i = 0 To ListSurver.Length - 1
    '                   strVale = ListSurver(i).Split(",")
    '                   Syntax = "INSERT INTO QuestionSurveyDetail(QDDID,OptionID,SurveyDate)" & _
    '                            "VALUES(" & strVale(0) & "," & strVale(1) & "," & svDatatime & ")"
    '                   objDB.sqlExecute(Syntax, objCnn)
    '               Next
    '           End If
    '           objCnn.Close()
    '           strResultText = ""
    '           Return True
    '       Catch ex As Exception
    '           strResultText = ex.ToString
    '           Return False
    '       End Try

    '   End Function

    '   <WebMethod()> _
    '   Public Function QuestionSurvey() As DataSet
    '       Dim Syntax As String = ""
    '       Dim ds As New DataSet
    '       Dim dt1 As New DataTable
    '       Dim dt As New DataTable

    '       GetconnectionAndWebConfiguration()
    '       dt1.TableName = "Question"
    '       dt.TableName = "QuestionSurvey"
    '       Syntax = "SELECT a.QDDID,a.QDDName,b.OptionID,b.OptionName FROM QuestionDefineData a,QuestionDefineOption b" & _
    '                " WHERE a.QDDID = b.QDDID And a.QDDSurvey = 1 And a.Deleted = 0 And a.Activated = 1" & _
    '                " ORDER BY a.QDDOrdering"
    '       dt = objDB.List(Syntax, objCnn)
    '       Syntax = "SELECT a.QDDID,a.QDDName FROM QuestionDefineData a" & _
    '              " WHERE a.QDDSurvey = 1 And a.Deleted = 0 And a.Activated = 1" & _
    '              " ORDER BY a.QDDOrdering"
    '       dt1 = objDB.List(Syntax, objCnn)
    '       objCnn.Close()
    '       ds.Tables.Add(dt1)
    '       ds.Tables.Add(dt)
    '       Return ds
    '   End Function

    '   <WebMethod()> _
    '   Public Function CheckOrder(ByVal ReferenceNo As String, ByRef strResultText As String) As DataSet
    '       Dim Syntax As String = ""
    '       Dim dt As New DataTable
    '       Dim ds As New DataSet
    '       Syntax = "select a.TransactionID,a.computerid,c.ProductCode,c.ProductName,c.ProductName1 ,e.KDSStatusID,e.KDSStatus" & _
    '               " from ordertransactionfront a,orderdetailfront b,products c ,kdstransactiondetailfront d,kdsstatus e" & _
    '               " where a.transactionid = b.TransactionID And a.computerid = b.computerid" & _
    '               " and a.transactionid=d.TransactionID and a.computerid=d.computerid" & _
    '               " and d.KDSStatus=e.KDSStatusID" & _
    '               " and a.referenceno ='" + ReferenceNo.ToString + "'" & _
    '               " and b.ProductID=c.ProductID" & _
    '               " union" & _
    '               " select a.TransactionID,a.computerid,c.ProductCode,c.ProductName,c.ProductName1 ,e.KDSStatusID,e.KDSStatus" & _
    '               " from ordertransaction a,orderdetail b,products c ,kdstransactiondetailfront d,kdsstatus e" & _
    '               " where a.transactionid = b.TransactionID And a.computerid = b.computerid" & _
    '               " and a.transactionid=d.TransactionID and a.computerid=d.computerid" & _
    '               " and d.KDSStatus=e.KDSStatusID" & _
    '               " and a.referenceno ='" + ReferenceNo.ToString + "'" & _
    '               " and b.ProductID=c.ProductID" & _
    '               " union" & _
    '               " select a.TransactionID,a.computerid,c.ProductCode,c.ProductName,c.ProductName1 ,IF(e.KDSStatusID is NULL,0,e.KDSStatusID) AS KDSStatusID,e.KDSStatus" & _
    '               " from ordertransactionfront a inner join orderdetailfront b on a.transactionid = b.TransactionID And a.computerid = b.computerid" & _
    '               " inner join products c on b.ProductID=c.ProductID" & _
    '               " left outer join kdstransactiondetailfront d on a.transactionid=d.TransactionID and a.computerid=d.computerid " & _
    '               " left outer join kdsstatus e on d.KDSStatus=e.KDSStatusID" & _
    '               " where a.referenceno ='" + ReferenceNo.ToString + "'"
    '       Try
    '           GetconnectionAndWebConfiguration()
    '           dt = objDB.List(Syntax, objCnn)
    '           objCnn.Close()
    '           strResultText = ""
    '           WriteErrorLogFile("msgKDS", "Succeede", "KDS_SaleMode_CheckStatus")
    '       Catch ex As Exception
    '           strResultText = ex.ToString
    '           WriteErrorLogFile("msgKDS", ex.ToString, "KDS_SaleMode_CheckStatus")
    '       End Try
    '       ds.Tables.Add(dt)
    '       Return ds
    '   End Function
    <WebMethod()> _
    Public Function Menu_GenerateXMLMenu(ByVal sXml As String) As String
        '**************************************************************************************************** 
        ' Function Name : GenerateXMLMenu 
        ' Author        : Supakorn Aonwichean 
        ' Description   : Generate the menu in an XML string with the required format 
        '**************************************************************************************************** 
        Try
            '**************************************************************************************************** 
            ' EstablishConnection
            '**************************************************************************************************** 
            EstablisConnection()
            '**************************************************************************************************** 
            ' Use a XmlDocument to load Xml file. 
            '**************************************************************************************************** 
            Dim strXML As String = ""
            Dim sXMLData As String = ""
            Dim xmlDoc As New XmlDocument()
            xmlDoc.LoadXml(sXml)

            Dim params As New Param_Data
            Dim validateShop As Boolean = False
            Dim validateKeyCode As Boolean = False
            For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("params")
                params.ShopID = xmlNode.ChildNodes(0).InnerText
                params.KeyCode = xmlNode.ChildNodes(1).InnerText
            Next

            validateShop = PromiseData.CheckShop(objDB, objCnn, params.ShopID)
            validateKeyCode = PromiseData.CheckKeyCode(params.KeyCode)

            If validateShop = True And validateKeyCode = True Then

                Dim dsMenu As New DataSet
                dsMenu = PromiseData.MenuData(objDB, objCnn, 3)
                '**************************************************************************************************** 
                ' Generate the munu in an XML string  
                '**************************************************************************************************** 
                Dim myDataMemorystream As New MemoryStream
                dsMenu.WriteXml(myDataMemorystream)

                Dim stream_reader As New StreamReader(myDataMemorystream, System.Text.Encoding.UTF8)
                myDataMemorystream.Seek(0, SeekOrigin.Begin)
                sXMLData = stream_reader.ReadToEnd()
                'sXMLData = WriteReponse(1, stream_reader.ReadToEnd())
            Else
                If validateKeyCode = False Then
                    sXMLData = WriteReponse(0, "!Invalid keycode. Please contact technical support Synature Technology Co., Ltd.")
                Else
                    If validateShop = False Then
                        sXMLData = WriteReponse(0, "!Invalid ShopId.")
                    End If
                End If
            End If
            '**************************************************************************************************** 
            ' Close connection
            '****************************************************************************************************
            CloseConnention()
            '**************************************************************************************************** 
            ' return the XML to the calling function 
            '**************************************************************************************************** 
            Return sXMLData
        Catch ex As Exception
            Throw
        End Try
    End Function
    <WebMethod()> _
   Public Function Order_AddOrders(ByVal sPOSData As String) As String
        '**************************************************************************************************** 
        ' Function Name : GenerateXMLMenu 
        ' Author        : Supakorn Aonwichean 
        ' Description   : Generate the menu in an XML string with the required format 
        '****************************************************************************************************        
        Try
            '**************************************************************************************************** 
            ' Use a XmlDocument to load Xml file. 
            '**************************************************************************************************** 
            Dim strXML As String = ""
            Dim xmlDoc As New XmlDocument()
            xmlDoc.LoadXml(sPOSData)

            '**************************************************************************************************** 
            ' Get staffcode and staffpassword for check staff permission.
            '**************************************************************************************************** 
            Dim staffCode As String = ""
            Dim staffPassword As String = ""
            For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("staff")
                staffCode = xmlNode.ChildNodes(0).InnerText
                staffPassword = xmlNode.ChildNodes(1).InnerText
            Next

            '**************************************************************************************************** 
            ' Get table name for check table take order and check table status.
            '**************************************************************************************************** 
            Dim tableName As String = ""
            For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("tableno")
                tableName = xmlNode.ChildNodes(0).InnerText
            Next

            Dim orderList As New List(Of Order_Data)
            Dim orderData As Order_Data
            Dim commentList As New List(Of Comment_Data)
            Dim commentData As Comment_Data
            For Each xmlNode As XmlNode In xmlDoc.GetElementsByTagName("orders")
                orderData = New Order_Data
                orderData.ProductCode = xmlNode.ChildNodes(0).InnerText
                orderData.ProductAmount = xmlNode.ChildNodes(1).InnerText
                commentData = New Comment_Data
                For Each xmlNodeComment As XmlNode In xmlDoc.GetElementsByTagName("comments")
                    commentData.CommentCode = xmlNodeComment.ChildNodes(0).InnerText
                    commentData.CommentAmount = xmlNodeComment.ChildNodes(1).InnerText
                    commentList.Add(commentData)
                Next
                orderList.Add(orderData)
            Next
            Return strXML

        Catch ex As Exception
            Throw
        End Try
    End Function
    Private Function WriteErrorLogFile(ByVal errorFrom As String, ByVal errorString As String, _
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
            wr.WriteLine("Message At " & Format(Now, "dd-MM-yyyy hh:mm"))
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
    Private Function WriteReponse(ByVal status As Integer, ByVal strData As String) As String
        Dim strXML As String = ""
        Try
            Dim myMemorystream As New MemoryStream
            Dim writer As New XmlTextWriter(myMemorystream, System.Text.Encoding.UTF8)

            '**************************************************************************************************** 
            ' Use indentation to make the result easier to debug 
            '**************************************************************************************************** 
            writer.Formatting = Formatting.Indented
            writer.Indentation = 4

            '**************************************************************************************************** 
            ' Write the XML declaration 
            '**************************************************************************************************** 
            writer.WriteStartDocument()

            '**************************************************************************************************** 
            ' Start the methodCall node 
            '**************************************************************************************************** 
            writer.WriteStartElement("methodResponse")

            '**************************************************************************************************** 
            ' Start the params node 
            '**************************************************************************************************** 
            writer.WriteStartElement("params")

            '****************************************************************************************** 
            ' Start the param node 
            '**************************************************************************************************** 
            writer.WriteStartElement("param")
            '**************************************************************************************************** 
            ' Start the value node 
            '**************************************************************************************************** 
            writer.WriteStartElement("value")
            '**************************************************************************************************** 
            ' Write the data 
            '**************************************************************************************************** 
            writer.WriteElementString("string", status)
            '**************************************************************************************************** 
            ' End the value node 
            '**************************************************************************************************** 
            writer.WriteEndElement()
            '**************************************************************************************************** 
            ' End the param node 
            '**************************************************************************************************** 
            writer.WriteEndElement()

            '****************************************************************************************** 

            '**************************************************************************************************** 
            ' Start the param node 
            '**************************************************************************************************** 
            writer.WriteStartElement("param")
            '**************************************************************************************************** 
            ' Start the value node 
            '**************************************************************************************************** 
            writer.WriteStartElement("value")
            '**************************************************************************************************** 
            ' Write the data 
            '**************************************************************************************************** 
            writer.WriteElementString("string", strData)
            '**************************************************************************************************** 
            ' End the value node 
            '**************************************************************************************************** 
            writer.WriteEndElement()
            '**************************************************************************************************** 
            ' End the param node 
            '**************************************************************************************************** 
            writer.WriteEndElement()

            '****************************************************************************************** 

            '**************************************************************************************************** 
            ' End the params node 
            '**************************************************************************************************** 
            writer.WriteEndElement()

            '**************************************************************************************************** 
            ' End the methodCall node 
            '**************************************************************************************************** 
            writer.WriteEndElement()

            '**************************************************************************************************** 
            ' End the document 
            '**************************************************************************************************** 
            writer.WriteEndDocument()

            writer.Flush()

            '**************************************************************************************************** 
            ' Use a StreamReader to display the result. 
            '**************************************************************************************************** 
            Dim stream_reader As New StreamReader(myMemorystream)

            myMemorystream.Seek(0, SeekOrigin.Begin)

            Dim sXML As String = stream_reader.ReadToEnd()

            writer.Close()

            '**************************************************************************************************** 
            ' return the XML to the calling function 
            '**************************************************************************************************** 
            Return sXML
        Catch ex As Exception
            Throw
        End Try

    End Function
End Class

Friend Class PromiseData
    Shared Function MenuData(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal shopId As Integer) As DataSet

        Dim strSQL As String = ""
        Dim dtproductgroup As New DataTable
        Dim dtproductdept As New DataTable
        Dim dtproducts As New DataTable
        Dim dtproductprice As New DataTable
        Dim dtfavoriteproducts As New DataTable
        Dim dtfavoriteproductpageindex As New DataTable
        Dim dtproductcomponent As New DataTable
        Dim dtpcomponentgroup As New DataTable
        Dim dtproducttype As New DataTable
        Dim dtsalemode As New DataTable
        Dim dsMenu As New DataSet("Menu")

        dtproductgroup = objDB.List("select * from productgroup where productlevelid=" & shopId & " and deleted=0", objCnn)
        dtproductgroup.TableName = "productgroup"

        strSQL = "select b.* from productgroup a, productdept b where a.productgroupid=b.productgroupid" & _
                 " and a.productlevelid=" & shopId & " and a.Deleted=0 and b.deleted=0"
        dtproductdept = objDB.List(strSQL, objCnn)
        dtproductdept.TableName = "productdept"

        strSQL = "select c.* from productgroup a, productdept b, products c" & _
                 " where a.productgroupid = b.productgroupid And b.productdeptid = c.productdeptid" & _
                 " and a.productlevelid=" & shopId & " and a.deleted=0 and b.deleted=0 and c.deleted=0"
        dtproducts = objDB.List(strSQL, objCnn)
        dtproducts.TableName = "products"

        strSQL = "select d.* from productgroup a, productdept b, products c, productprice d" & _
                 " where a.productgroupid = b.productgroupid And b.productdeptid = c.productdeptid" & _
                 " and c.productid=d.productid" & _
                 " and a.productlevelid=" & shopId & " and a.deleted=0 and b.deleted=0 and c.deleted=0"
        dtproductprice = objDB.List(strSQL, objCnn)
        dtproductprice.TableName = "productprice"

        dtfavoriteproducts = objDB.List("select * from favoriteproducts where productlevelid=" & shopId, objCnn)
        dtfavoriteproducts.TableName = "favoriteproducts"

        dtfavoriteproductpageindex = objDB.List("select * from favoriteproductpageindex where productlevelid=" & shopId, objCnn)
        dtfavoriteproductpageindex.TableName = "favoriteproductpageindex"

        dtproductcomponent = objDB.List("select * from productcomponent", objCnn)
        dtproductcomponent.TableName = "productcomponent"

        dtpcomponentgroup = objDB.List("select * from pcomponentgroup", objCnn)
        dtpcomponentgroup.TableName = "pcomponentgroup"

        dtproducttype = objDB.List("select * from producttype", objCnn)
        dtproducttype.TableName = "producttype"

        dtsalemode = objDB.List("select * from salemode", objCnn)
        dtsalemode.TableName = "salemode"

        dsMenu.Tables.Add(dtfavoriteproductpageindex)
        dsMenu.Tables.Add(dtfavoriteproducts)
        dsMenu.Tables.Add(dtpcomponentgroup)
        dsMenu.Tables.Add(dtproductcomponent)
        dsMenu.Tables.Add(dtproductdept)
        dsMenu.Tables.Add(dtproductgroup)
        dsMenu.Tables.Add(dtproductprice)
        dsMenu.Tables.Add(dtproducts)
        dsMenu.Tables.Add(dtproducttype)
        dsMenu.Tables.Add(dtsalemode)

        Return dsMenu
    End Function
    Shared Function CheckStaff(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal staffCode As String, ByVal staffPassword As String) As Boolean

        'create the SHA1 hash of the password provided by the user
        Dim strHash As String
        strHash = FormsAuthentication.HashPasswordForStoringInConfigFile(staffPassword, "SHA1")

        Dim dtData As New DataTable
        Try
            dtData = objDB.List("select * from staffs where deleted=0 and staffcode='" & staffCode & "' and staffpassword='" & strHash & "'", objCnn)
            If dtData.Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw
        End Try
    End Function
    Shared Function CheckTable(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal tableName As String) As DataTable
        Dim strSQL As String = ""
        strSQL = "select b.* from tablezone a,tableno b where a.zoneid=b.zoneid and a.shopid=8 and a.deleted=0 and tablename='" & tableName & "'"
        Try
            Return objDB.List(strSQL, objCnn)
        Catch ex As Exception
            Throw
        End Try
    End Function
    Shared Function ConvertXMLToBase64(ByVal strXML As String) As String
        Dim CodePage1252 As System.Text.Encoding = System.Text.Encoding.GetEncoding(1252)
        Return System.Convert.ToBase64String(CodePage1252.GetBytes(strXML.ToCharArray))
    End Function
    Shared Function CheckKeyCode(ByVal keyCode As String) As Boolean
        Dim sKeyCode As String = "123456789"
        If keyCode = sKeyCode Then
            Return True
        Else
            Return False
        End If
    End Function
    Shared Function CheckShop(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal shopId As Integer) As Boolean
        Dim dtData As New DataTable
        dtData = objDB.List("select * from productlevel where productlevelId=" & shopId, objCnn)
        If dtData.Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function
End Class
Public Class Order_Data
    Protected _ProductCode As String = ""
    Protected _ProductAmount As Integer
    Property ProductCode() As String
        Get
            Return _ProductCode
        End Get
        Set(ByVal value As String)
            _ProductCode = value
        End Set
    End Property
    Property ProductAmount() As Integer
        Get
            Return _ProductAmount
        End Get
        Set(ByVal value As Integer)
            _ProductAmount = value
        End Set
    End Property
End Class
Public Class Comment_Data
    Protected _CommentCode As String = ""
    Protected _CommentAmount As Integer
    Property CommentCode() As String
        Get
            Return _CommentCode
        End Get
        Set(ByVal value As String)
            _CommentCode = value
        End Set
    End Property
    Property CommentAmount() As Integer
        Get
            Return _CommentAmount
        End Get
        Set(ByVal value As Integer)
            _CommentAmount = value
        End Set
    End Property
End Class
Public Class Param_Data
    Protected _ShopID As Integer
    Protected _KeyCode As String
    Property ShopID() As Integer
        Get
            Return _ShopID
        End Get
        Set(ByVal value As Integer)
            _ShopID = value
        End Set
    End Property
    Property KeyCode() As String
        Get
            Return _KeyCode
        End Get
        Set(ByVal value As String)
            _KeyCode = value
        End Set
    End Property
End Class
