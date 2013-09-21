Imports Microsoft.VisualBasic
Imports System.Data
Imports MySql.Data.MySqlClient
Imports POSMySQL.POSControl
Imports System.Collections.Generic
<Serializable()> _
Public Class ClsWebserviceInfo

    Public Shared Function Members(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal searchBy As SearchMemberBy, ByVal memberCode As String, _
                            ByVal memberMobile As String, ByRef memberData As Member_Data, ByRef dsMemberData As DataSet, ByRef strResultText As String) As Boolean
        Try
            Dim dtData As New DataTable
            Dim dtMember As New DataTable
            Dim ds As New DataSet
            memberData = New Member_Data
            ds = ClsWebServiceData.MemberData(objDB, objCnn, searchBy, memberCode, memberMobile)
            dtData = ds.Tables(0).Copy
            dtMember = ds.Tables(1).Copy
            dsMemberData.Tables.Add(dtMember)

            If dtData.Rows.Count > 0 Then
                memberData.MemberID = dtData.Rows(0)("MemberID")
                memberData.MemberGroupID = dtData.Rows(0)("MemberGroupID")
                If Not IsDBNull(dtData.Rows(0)("MemberGroupCode")) Then
                    memberData.MemberGroupCode = dtData.Rows(0)("MemberGroupCode").ToString
                Else
                    memberData.MemberGroupCode = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberGroupName")) Then
                    memberData.MemberGroupName = dtData.Rows(0)("MemberGroupName").ToString
                Else
                    memberData.MemberGroupName = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberCode")) Then
                    memberData.MemberCode = dtData.Rows(0)("MemberCode").ToString
                Else
                    memberData.MemberCode = ""
                End If
                memberData.MemberGendor = dtData.Rows(0)("MemberGender")

                If Not IsDBNull(dtData.Rows(0)("MemberFirstName")) Then
                    memberData.MemberFirstName = dtData.Rows(0)("MemberFirstName").ToString
                Else
                    memberData.MemberFirstName = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberLastName")) Then
                    memberData.MemberLastName = dtData.Rows(0)("MemberLastName").ToString
                Else
                    memberData.MemberLastName = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberAddress1")) Then
                    memberData.MemberAddress1 = dtData.Rows(0)("MemberAddress1").ToString
                Else
                    memberData.MemberAddress1 = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberAddress2")) Then
                    memberData.MemberAddress2 = dtData.Rows(0)("MemberAddress2").ToString
                Else
                    memberData.MemberAddress2 = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberTelephone")) Then
                    memberData.MemberTelephone = dtData.Rows(0)("MemberTelephone").ToString
                Else
                    memberData.MemberTelephone = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberCity")) Then
                    memberData.MemberCity = dtData.Rows(0)("MemberCity").ToString
                Else
                    memberData.MemberCity = ""
                End If
                memberData.MemberProvince = dtData.Rows(0)("MemberProvince")
                If Not IsDBNull(dtData.Rows(0)("MemberZipCode")) Then
                    memberData.MemberZipCode = dtData.Rows(0)("MemberZipCode").ToString
                Else
                    memberData.MemberZipCode = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberMobile")) Then
                    memberData.MemberMobile = dtData.Rows(0)("MemberMobile").ToString
                Else
                    memberData.MemberMobile = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberFax")) Then
                    memberData.MemberFax = dtData.Rows(0)("MemberFax").ToString
                Else
                    memberData.MemberFax = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberEmail")) Then
                    memberData.MemberEmail = dtData.Rows(0)("MemberEmail").ToString
                Else
                    memberData.MemberEmail = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberBirthDay")) Then
                    memberData.MemberBirthDay = dtData.Rows(0)("MemberBirthDay")
                Else
                    memberData.MemberBirthDay = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberExpiration")) Then
                    memberData.MemberExpiration = dtData.Rows(0)("MemberExpiration")
                Else
                    memberData.MemberExpiration = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("InputDate")) Then
                    memberData.MemberInputDate = dtData.Rows(0)("InputDate")
                Else
                    memberData.MemberInputDate = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("LastUseDate")) Then
                    memberData.MemberLastUseDate = dtData.Rows(0)("LastUseDate")
                Else
                    memberData.MemberLastUseDate = Date.MinValue
                End If
                memberData.MemberActivate = dtData.Rows(0)("Activated")

                If Not IsDBNull(dtData.Rows(0)("TotalPoint")) Then
                    memberData.MemberTotalPoint = dtData.Rows(0)("TotalPoint")
                Else
                    memberData.MemberTotalPoint = 0.0
                End If
                If Not IsDBNull(dtData.Rows(0)("Updatedate")) Then
                    memberData.MemberUpdatePoint = dtData.Rows(0)("UpdateDate")
                Else
                    memberData.MemberUpdatePoint = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberUpdatedate")) Then
                    memberData.MemberUpdaetDate = dtData.Rows(0)("MemberUpdatedate")
                Else
                    memberData.MemberUpdaetDate = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("UpDateBy")) Then
                    memberData.MemberUpdateBy = dtData.Rows(0)("UpDateBy")
                Else
                    memberData.MemberUpdateBy = 0
                End If
                If Not IsDBNull(dtData.Rows(0)("MemberAdditional")) Then
                    memberData.MemberAdditional = dtData.Rows(0)("MemberAdditional")
                Else
                    memberData.MemberAdditional = ""
                End If
                memberData.MemberGroupID = dtData.Rows(0)("MemberGroupID")
                memberData.PackageDetail = MemberPackageListData(objDB, objCnn, dtData.Rows(0)("MemberID"))
                strResultText = ""
                Return True
            Else
                strResultText = "ไม่พบข้อมูลสมาชิก"
                Return False
            End If
        Catch ex As Exception
            strResultText = ex.Message.ToString
            Return False
        End Try
    End Function
    Public Shared Function MembersList(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal searyBy As SearchMemberBy, ByVal paramSearch As String, _
                           ByRef listMemberData As List(Of Member_Data), ByRef dsMemberData As DataSet, ByRef strResultText As String) As Boolean
        Try
            Dim dtData As New DataTable
            Dim dtMember As New DataTable
            Dim ds As New DataSet
            Dim memberData As Member_Data
            listMemberData = New List(Of Member_Data)
            ds = ClsWebServiceData.MemberDataList(objDB, objCnn, searyBy, paramSearch)
            dtData = ds.Tables(0).Copy
            dtMember = ds.Tables(1).Copy
            dsMemberData.Tables.Add(dtMember)

            If dtData.Rows.Count > 0 Then
                For i As Integer = 0 To dtData.Rows.Count - 1
                    memberData = New Member_Data
                    memberData.MemberID = dtData.Rows(i)("MemberID")
                    memberData.MemberGroupID = dtData.Rows(i)("MemberGroupID")
                    If Not IsDBNull(dtData.Rows(0)("MemberGroupCode")) Then
                        memberData.MemberGroupCode = dtData.Rows(i)("MemberGroupCode").ToString
                    Else
                        memberData.MemberGroupCode = ""
                    End If
                    If Not IsDBNull(dtData.Rows(0)("MemberGroupName")) Then
                        memberData.MemberGroupName = dtData.Rows(i)("MemberGroupName").ToString
                    Else
                        memberData.MemberGroupName = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberCode")) Then
                        memberData.MemberCode = dtData.Rows(i)("MemberCode").ToString
                    Else
                        memberData.MemberCode = ""
                    End If
                    memberData.MemberGendor = dtData.Rows(i)("MemberGender")

                    If Not IsDBNull(dtData.Rows(i)("MemberFirstName")) Then
                        memberData.MemberFirstName = dtData.Rows(i)("MemberFirstName").ToString
                    Else
                        memberData.MemberFirstName = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberLastName")) Then
                        memberData.MemberLastName = dtData.Rows(i)("MemberLastName").ToString
                    Else
                        memberData.MemberLastName = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberAddress1")) Then
                        memberData.MemberAddress1 = dtData.Rows(i)("MemberAddress1").ToString
                    Else
                        memberData.MemberAddress1 = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberAddress2")) Then
                        memberData.MemberAddress2 = dtData.Rows(i)("MemberAddress2").ToString
                    Else
                        memberData.MemberAddress2 = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberTelephone")) Then
                        memberData.MemberTelephone = dtData.Rows(i)("MemberTelephone").ToString
                    Else
                        memberData.MemberTelephone = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberCity")) Then
                        memberData.MemberCity = dtData.Rows(i)("MemberCity").ToString
                    Else
                        memberData.MemberCity = ""
                    End If
                    memberData.MemberProvince = dtData.Rows(i)("MemberProvince")

                    If Not IsDBNull(dtData.Rows(i)("MemberZipCode")) Then
                        memberData.MemberZipCode = dtData.Rows(i)("MemberZipCode").ToString
                    Else
                        memberData.MemberZipCode = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberMobile")) Then
                        memberData.MemberMobile = dtData.Rows(i)("MemberMobile").ToString
                    Else
                        memberData.MemberMobile = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberFax")) Then
                        memberData.MemberFax = dtData.Rows(i)("MemberFax").ToString
                    Else
                        memberData.MemberFax = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberEmail")) Then
                        memberData.MemberEmail = dtData.Rows(i)("MemberEmail").ToString
                    Else
                        memberData.MemberEmail = ""
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberBirthDay")) Then
                        memberData.MemberBirthDay = dtData.Rows(i)("MemberBirthDay").ToString
                    Else
                        memberData.MemberBirthDay = Date.MinValue
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberExpiration")) Then
                        memberData.MemberExpiration = dtData.Rows(i)("MemberExpiration")
                    Else
                        memberData.MemberExpiration = Date.MinValue
                    End If
                    If Not IsDBNull(dtData.Rows(i)("InputDate")) Then
                        memberData.MemberInputDate = dtData.Rows(i)("InputDate")
                    Else
                        memberData.MemberInputDate = Date.MinValue
                    End If
                    If Not IsDBNull(dtData.Rows(i)("LastUseDate")) Then
                        memberData.MemberLastUseDate = dtData.Rows(i)("LastUseDate")
                    Else
                        memberData.MemberLastUseDate = Date.MinValue
                    End If
                    memberData.MemberActivate = dtData.Rows(i)("Activated")

                    If Not IsDBNull(dtData.Rows(i)("TotalPoint")) Then
                        memberData.MemberTotalPoint = dtData.Rows(i)("TotalPoint")
                    Else
                        memberData.MemberTotalPoint = 0.0
                    End If
                    If Not IsDBNull(dtData.Rows(i)("Updatedate")) Then
                        memberData.MemberUpdatePoint = dtData.Rows(i)("UpdateDate")
                    Else
                        memberData.MemberUpdatePoint = Date.MinValue
                    End If
                    If Not IsDBNull(dtData.Rows(i)("MemberUpdatedate")) Then
                        memberData.MemberUpdaetDate = dtData.Rows(i)("MemberUpdatedate")
                    Else
                        memberData.MemberUpdaetDate = Date.MinValue
                    End If
                    If Not IsDBNull(dtData.Rows(i)("UpDateBy")) Then
                        memberData.MemberUpdateBy = dtData.Rows(i)("UpDateBy")
                    Else
                        memberData.MemberUpdateBy = 0
                    End If
                    If Not IsDBNull(dtData.Rows(0)("MemberAdditional")) Then
                        memberData.MemberAdditional = dtData.Rows(0)("MemberAdditional")
                    Else
                        memberData.MemberAdditional = ""
                    End If
                    memberData.PackageDetail = MemberPackageListData(objDB, objCnn, dtData.Rows(0)("MemberID"))
                    listMemberData.Add(memberData)
                Next
            End If
            strResultText = ""
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            Return False
        End Try
    End Function
    Shared Function MemberPackageListData(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal memberID As Integer) As List(Of packagedetail)
        Dim packageData As packagedetail
        Dim packageListData As New List(Of packagedetail)
        Dim packageHistoryData As Packagehistory

        Dim ds As DataSet
        Dim dtPackageData As DataTable
        Dim dtPackageHistory As DataTable
        Dim expression As String
        Dim foundRows() As DataRow

        ds = ClsWebServiceData.PackageDetailData(objDB, objCnn, memberID)
        dtPackageData = ds.Tables("packagedetail")
        dtPackageHistory = ds.Tables("packagehistory")

        If dtPackageData.Rows.Count > 0 Then
            For i As Integer = 0 To dtPackageData.Rows.Count - 1

                packageData = New packagedetail
                packageData.PackageID = dtPackageData.Rows(i)("PackageID")
                packageData.ProductLevelID = dtPackageData.Rows(i)("ProductLevelID")
                packageData.ProductCode = dtPackageData.Rows(i)("ProductCode")
                packageData.TransactionID = dtPackageData.Rows(i)("TransactionID")
                packageData.ComputerID = dtPackageData.Rows(i)("ComputerID")
                packageData.OrderDetailID = dtPackageData.Rows(i)("OrderDetailID")
                If Not IsDBNull(dtPackageData.Rows(i)("ExpireDate")) Then
                    packageData.ExpireDate = dtPackageData.Rows(i)("ExpireDate")
                End If
                If Not IsDBNull(dtPackageData.Rows(i)("PurchaseDate")) Then
                    packageData.PurchaseDate = dtPackageData.Rows(i)("PurchaseDate")
                End If
                packageData.MemberID = dtPackageData.Rows(i)("MemberID")
                If Not IsDBNull(dtPackageData.Rows(i)("PackageFirstName")) Then
                    packageData.PackageFirstName = dtPackageData.Rows(i)("PackageFirstName")
                End If
                If Not IsDBNull(dtPackageData.Rows(i)("PackageLastName")) Then
                    packageData.PackageLastName = dtPackageData.Rows(i)("PackageLastName")
                End If
                If Not IsDBNull(dtPackageData.Rows(i)("PackageDisplayName")) Then
                    packageData.PackageDisplayName = dtPackageData.Rows(i)("PackageDisplayName")
                End If
                If Not IsDBNull(dtPackageData.Rows(i)("PackageNumber")) Then
                    packageData.PackageNumber = dtPackageData.Rows(i)("PackageNumber")
                End If
                If Not IsDBNull(dtPackageData.Rows(i)("PackageNote")) Then
                    packageData.PackageNote = dtPackageData.Rows(i)("PackageNote")
                End If
                packageData.Deleted = dtPackageData.Rows(i)("Deleted")
                If Not IsDBNull(dtPackageData.Rows(i)("UpdateDate")) Then
                    packageData.UpdateDate = dtPackageData.Rows(i)("UpdateDate")
                End If
                packageData.VATType = dtPackageData.Rows(i)("VATType")
                packageData.PackageProductType = dtPackageData.Rows(i)("PackageProductType")
                packageData.PackageOriginalAmount = dtPackageData.Rows(i)("PackageOriginalAmount")
                packageData.MaxAllowEditAmount = dtPackageData.Rows(i)("MaxAllowEditAmount")
                packageData.PackageCurrentAmount = dtPackageData.Rows(i)("PackageCurrentAmount")
                packageData.PackagePrice = dtPackageData.Rows(i)("PackagePrice")
                packageData.AllowUseInShop = dtPackageData.Rows(i)("AllowUseInShop")

                expression = "ProductlevelID=" & dtPackageData.Rows(i)("ProductlevelID") & " AND PackageID=" & dtPackageData.Rows(i)("PackageID")
                foundRows = dtPackageHistory.Select(expression)

                If foundRows.GetUpperBound(0) >= 0 Then
                    Dim packageHistoryListData As New List(Of Packagehistory)
                    For j As Integer = 0 To foundRows.Length - 1
                        packageHistoryData = New Packagehistory
                        packageHistoryData.PackageHistoryID = foundRows(j)("PackageHistoryID")
                        packageHistoryData.PackageID = foundRows(j)("PackageID")
                        packageHistoryData.ProductLevelID = foundRows(j)("ProductLevelID")
                        packageHistoryData.TransactionID = foundRows(j)("TransactionID")
                        packageHistoryData.ComputerID = foundRows(j)("ComputerID")
                        packageHistoryData.OrderDetailID = foundRows(j)("OrderDetailID")
                        packageHistoryData.ProductCode = foundRows(j)("ProductCode")
                        packageHistoryData.PackageStatus = foundRows(j)("PackageStatus")
                        packageHistoryData.CommissionPrice = foundRows(j)("CommissionPrice")
                        packageHistoryData.CommissionPriceVAT = foundRows(j)("CommissionPriceVAT")
                        packageHistoryData.AdditionalPrice = foundRows(j)("AdditionalPrice")
                        packageHistoryData.PackageSessionAmount = foundRows(j)("PackageSessionAmount")
                        If Not IsDBNull(dtPackageData.Rows(i)("UpdateDate")) Then
                            packageHistoryData.Updatedate = dtPackageData.Rows(i)("UpdateDate")
                        End If
                        packageHistoryListData.Add(packageHistoryData)
                    Next
                    packageData.PackageHistory = packageHistoryListData
                End If
                packageListData.Add(packageData)
            Next

        End If
        Return packageListData
    End Function
    Public Shared Function SummaryPoint(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal memberID As Integer, ByRef memberData As SummaryPoint_Data, ByRef strResultText As String) As Boolean
        Try
            Dim dtData As DataTable
            memberData = New SummaryPoint_Data
            dtData = ClsWebServiceData.GetPointMemberData(objDB, objCnn, memberID)
            If dtData.Rows.Count > 0 Then
                memberData.MemberID = dtData.Rows(0)("MemberID")
                If Not IsDBNull(dtData.Rows(0)("TotalPoint")) Then
                    memberData.TotalPoint = dtData.Rows(0)("TotalPoint")
                Else
                    memberData.TotalPoint = 0.0
                End If
                If Not IsDBNull(dtData.Rows(0)("Updatedate")) Then
                    memberData.UpdatePoint = dtData.Rows(0)("UpdateDate")
                Else
                    memberData.UpdatePoint = Date.MinValue
                End If
            End If
            strResultText = ""
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            Return False
        End Try
    End Function
    Public Shared Function SoftwareVersion(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal programTypeID As Integer, ByRef sotfwareData As Softwareversion_Data, ByRef strResultText As String) As Boolean
        Try
            Dim dtData As DataTable
            sotfwareData = New Softwareversion_Data
            dtData = ClsWebServiceData.GetSoftwareVersion(objDB, objCnn, programTypeID)
            If dtData.Rows.Count > 0 Then
                sotfwareData.ProgramTypeID = dtData.Rows(0)("ProgramTypeID")
                If Not IsDBNull(dtData.Rows(0)("FileVersion")) Then
                    sotfwareData.FileVersion = dtData.Rows(0)("FileVersion")
                Else
                    sotfwareData.FileVersion = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("FTPPassword")) Then
                    sotfwareData.FTPPassword = dtData.Rows(0)("FTPPassword")
                Else
                    sotfwareData.FTPPassword = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("FTPUserName")) Then
                    sotfwareData.FTPUserName = dtData.Rows(0)("FTPUserName")
                Else
                    sotfwareData.FTPUserName = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("IPHQServer")) Then
                    sotfwareData.IPHQServer = dtData.Rows(0)("IPHQServer")
                Else
                    sotfwareData.IPHQServer = ""
                End If
                If Not IsDBNull(dtData.Rows(0)("LastUpdate")) Then
                    sotfwareData.LastUpdate = dtData.Rows(0)("LastUpdate")
                Else
                    sotfwareData.LastUpdate = Date.MinValue
                End If
                If Not IsDBNull(dtData.Rows(0)("UpdateServerPath")) Then
                    sotfwareData.UpdateServerPath = dtData.Rows(0)("UpdateServerPath")
                Else
                    sotfwareData.UpdateServerPath = ""
                End If

            End If
            strResultText = ""
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            Return False
        End Try
    End Function
    Shared Function News_ListSection(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal sectionID As Integer) As List(Of News_SectionData)
        Dim dtData As DataTable
        Dim sectionData As News_SectionData
        Dim listSectionData As New List(Of News_SectionData)
        dtData = ClsWebServiceData.News_SectionData(objDB, objCnn, sectionID)
        If dtData.Rows.Count > 0 Then
            For i As Integer = 0 To dtData.Rows.Count - 1
                sectionData = New News_SectionData
                sectionData.SectionID = dtData.Rows(i)("SectionID")
                If Not IsDBNull(dtData.Rows(i)("SectionTitle")) Then
                    sectionData.SectionTitle = dtData.Rows(i)("SectionTitle")
                Else
                    sectionData.SectionTitle = ""
                End If
                If Not IsDBNull(dtData.Rows(i)("SectionDescription")) Then
                    sectionData.SectionDescription = dtData.Rows(i)("SectionDescription")
                Else
                    sectionData.SectionDescription = ""
                End If
                If Not IsDBNull(dtData.Rows(i)("Published")) Then
                    sectionData.Published = dtData.Rows(i)("Published")
                Else
                    sectionData.Published = ""
                End If
                listSectionData.Add(sectionData)
            Next
        End If
        Return listSectionData
    End Function
    Shared Function News_ListCategory(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal shopID As Integer, ByVal sectionID As Integer, ByVal limitContent As Integer) As List(Of News_CategoryData)
        Dim dtData As New DataTable
        Dim categoryData As News_CategoryData
        Dim listCategoryData As New List(Of News_CategoryData)
        dtData = ClsWebServiceData.News_CategoryData(objDB, objCnn, -1, sectionID)
        If dtData.Rows.Count > 0 Then
            For i As Integer = 0 To dtData.Rows.Count - 1
                categoryData = New News_CategoryData
                categoryData.SectionID = dtData.Rows(i)("SectionID")
                categoryData.CategoryID = dtData.Rows(i)("CategoryID")
                If Not IsDBNull(dtData.Rows(i)("CategoryTitle")) Then
                    categoryData.CategoryTitle = dtData.Rows(i)("CategoryTitle")
                Else
                    categoryData.CategoryTitle = ""
                End If
                If Not IsDBNull(dtData.Rows(i)("CategoryDescription")) Then
                    categoryData.CategoryDescription = dtData.Rows(i)("CategoryDescription")
                Else
                    categoryData.CategoryDescription = ""
                End If
                '================= ContentData =====================
                Dim dtContentData As New DataTable
                Dim listContentData As New List(Of News_ContentData)
                Dim contentData As News_ContentData
                dtContentData = ClsWebServiceData.News_ContentData(objDB, objCnn, sectionID, dtData.Rows(i)("CategoryID"), -1, 1, shopID, limitContent)
                categoryData.ContentTotal = dtContentData.Rows.Count
                If dtContentData.Rows.Count > 0 Then
                    For j As Integer = 0 To dtContentData.Rows.Count - 1
                        contentData = New News_ContentData
                        contentData.SectionID = dtContentData.Rows(j)("SectionID")
                        contentData.CategoryID = dtContentData.Rows(j)("CategoryID")
                        contentData.ContentID = dtContentData.Rows(j)("ContentID")
                        contentData.Published = dtContentData.Rows(j)("Published")
                        If Not IsDBNull(dtContentData.Rows(j)("Title")) Then
                            contentData.Title = dtContentData.Rows(j)("Title")
                        Else
                            contentData.Title = ""
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("Status")) Then
                            contentData.Status = dtContentData.Rows(j)("Status")
                        Else
                            contentData.Status = ""
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("Description")) Then
                            contentData.Description = dtContentData.Rows(j)("Description")
                        Else
                            contentData.Description = ""
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("Detail")) Then
                            contentData.Detail = dtContentData.Rows(j)("Detail")
                        Else
                            contentData.Detail = ""
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("Created")) Then
                            contentData.Created = dtContentData.Rows(j)("Created")
                        Else
                            contentData.Created = Date.MinValue
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("Modified")) Then
                            contentData.Modified = dtContentData.Rows(j)("Modified")
                        Else
                            contentData.Modified = Date.MinValue
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("StartPublished")) Then
                            contentData.StartPublished = dtContentData.Rows(j)("StartPublished")
                        Else
                            contentData.StartPublished = Date.MinValue
                        End If
                        If Not IsDBNull(dtContentData.Rows(j)("FinishPublished")) Then
                            contentData.FinishPublished = dtContentData.Rows(j)("FinishPublished")
                        Else
                            contentData.FinishPublished = Date.MinValue
                        End If
                        contentData.InsertStaffID = dtContentData.Rows(j)("InsertStaffID")
                        contentData.UpdateStaffID = dtContentData.Rows(j)("UpdateStaffID")

                        listContentData.Add(contentData)
                    Next
                End If
                categoryData.ContentDetail = listContentData
                listCategoryData.Add(categoryData)
            Next
        End If
        Return listCategoryData
    End Function

End Class
<Serializable()> _
Public Class Member_Data
    Protected _memberID As Integer = 0
    Protected _memberGroupID As Integer = 0
    Protected _memberGroupCode As String = ""
    Protected _memberGroupName As String = ""
    Protected _memberCode As String = ""
    Protected _memberGendor As Integer = 0
    Protected _memberFirstName As String = ""
    Protected _memberLastName As String = ""
    Protected _memberAddress1 As String = ""
    Protected _memberAddress2 As String = ""
    Protected _memberCity As String = ""
    Protected _memberProvince As Integer = 0
    Protected _memberZipCode As String = ""
    Protected _memberTelephone As String = ""
    Protected _memberMobile As String = ""
    Protected _memberFax As String = ""
    Protected _memberEmail As String = ""
    Protected _memberBirthDay As Date
    Protected _memberActivate As Integer = 0
    Protected _memberExpiration As Date
    Protected _memberInputDate As Date
    Protected _memberLastUseDate As Date
    Protected _memberTotalPoint As Decimal = 0
    Protected _memberUpdatePoint As DateTime
    Protected _memberUpdateDate As DateTime
    Protected _memberUpdateBy As Integer = 0
    Protected _memberAdditional As String = ""
    Protected _PackageDetail As List(Of packagedetail)

    Property MemberID() As Integer
        Get
            Return _memberID
        End Get
        Set(ByVal value As Integer)
            _memberID = value
        End Set
    End Property
    Property MemberGroupID() As Integer
        Get
            Return _memberGroupID
        End Get
        Set(ByVal value As Integer)
            _memberGroupID = value
        End Set
    End Property
    Property MemberGroupCode() As String
        Get
            Return _memberGroupCode
        End Get
        Set(ByVal value As String)
            _memberGroupCode = value
        End Set
    End Property
    Property MemberGroupName() As String
        Get
            Return _memberGroupName
        End Get
        Set(ByVal value As String)
            _memberGroupName = value
        End Set
    End Property
    Property MemberCode() As String
        Get
            Return _memberCode
        End Get
        Set(ByVal value As String)
            _memberCode = value
        End Set
    End Property
    Property MemberGendor() As Integer
        Get
            Return _memberGendor
        End Get
        Set(ByVal value As Integer)
            _memberGendor = value
        End Set
    End Property
    Property MemberFirstName() As String
        Get
            Return _memberFirstName
        End Get
        Set(ByVal value As String)
            _memberFirstName = value
        End Set
    End Property
    Property MemberLastName() As String
        Get
            Return _memberLastName
        End Get
        Set(ByVal value As String)
            _memberLastName = value
        End Set
    End Property
    Property MemberAddress1() As String
        Get
            Return _memberAddress1
        End Get
        Set(ByVal value As String)
            _memberAddress1 = value
        End Set
    End Property
    Property MemberAddress2() As String
        Get
            Return _memberAddress2
        End Get
        Set(ByVal value As String)
            _memberAddress2 = value
        End Set
    End Property
    Property MemberCity() As String
        Get
            Return _memberCity
        End Get
        Set(ByVal value As String)
            _memberCity = value
        End Set
    End Property
    Property MemberProvince() As Integer
        Get
            Return _memberProvince
        End Get
        Set(ByVal value As Integer)
            _memberProvince = value
        End Set
    End Property
    Property MemberZipCode() As String
        Get
            Return _memberZipCode
        End Get
        Set(ByVal value As String)
            _memberZipCode = value
        End Set
    End Property
    Property MemberTelephone() As String
        Get
            Return _memberTelephone
        End Get
        Set(ByVal value As String)
            _memberTelephone = value
        End Set
    End Property
    Property MemberMobile() As String
        Get
            Return _memberMobile
        End Get
        Set(ByVal value As String)
            _memberMobile = value
        End Set
    End Property
    Property MemberFax() As String
        Get
            Return _memberFax
        End Get
        Set(ByVal value As String)
            _memberFax = value
        End Set
    End Property
    Property MemberEmail() As String
        Get
            Return _memberEmail
        End Get
        Set(ByVal value As String)
            _memberEmail = value
        End Set
    End Property
    Property MemberBirthDay() As Date
        Get
            Return _memberBirthDay
        End Get
        Set(ByVal value As Date)
            _memberBirthDay = value
        End Set
    End Property
    Property MemberActivate() As Integer
        Get
            Return _memberActivate
        End Get
        Set(ByVal value As Integer)
            _memberActivate = value
        End Set
    End Property
    Property MemberExpiration() As Date
        Get
            Return _memberExpiration
        End Get
        Set(ByVal value As Date)
            _memberExpiration = value
        End Set
    End Property
    Property MemberInputDate() As Date
        Get
            Return _memberInputDate
        End Get
        Set(ByVal value As Date)
            _memberInputDate = value
        End Set
    End Property
    Property MemberLastUseDate() As Date
        Get
            Return _memberLastUseDate
        End Get
        Set(ByVal value As Date)
            _memberLastUseDate = value
        End Set
    End Property
    Property MemberUpdatePoint() As DateTime
        Get
            Return _memberUpdatePoint
        End Get
        Set(ByVal value As DateTime)
            _memberUpdatePoint = value
        End Set
    End Property
    Property MemberUpdaetDate() As DateTime
        Get
            Return _memberUpdateDate
        End Get
        Set(ByVal value As DateTime)
            _memberUpdateDate = value
        End Set
    End Property
    Property MemberUpdateBy() As Integer
        Get
            Return _memberUpdateBy
        End Get
        Set(ByVal value As Integer)
            _memberUpdateBy = value
        End Set
    End Property
    Property MemberTotalPoint() As Decimal
        Get
            Return _memberTotalPoint
        End Get
        Set(ByVal value As Decimal)
            _memberTotalPoint = value
        End Set
    End Property
    Property MemberAdditional() As String
        Get
            Return _memberAdditional
        End Get
        Set(ByVal value As String)
            _memberAdditional = value
        End Set
    End Property
    Property PackageDetail() As List(Of packagedetail)
        Get
            Return _PackageDetail
        End Get
        Set(ByVal value As List(Of packagedetail))
            _PackageDetail = value
        End Set
    End Property
End Class
<Serializable()> _
Public Class packagedetail
    Protected _PackageID As Integer
    Protected _ProductLevelID As Integer
    Protected _ProductCode As String
    Protected _TransactionID As Integer
    Protected _ComputerID As Integer
    Protected _OrderDetailID As Integer
    Protected _ExpireDate As DateTime
    Protected _PurchaseDate As DateTime
    Protected _MemberID As Integer
    Protected _PackageFirstName As String
    Protected _PackageLastName As String
    Protected _PackageDisplayName As String
    Protected _PackageNumber As String
    Protected _PackageNote As String
    Protected _Deleted As Integer
    Protected _UpdateDate As DateTime
    Protected _VATType As Integer
    Protected _PackageProductType As Integer
    Protected _PackageOriginalAmount As Decimal
    Protected _PackageCurrentAmount As Decimal
    Protected _MaxAllowEditAmount As Decimal
    Protected _PackagePrice As Integer
    Protected _AllowUseInShop As String
    Protected _PackageHistory As List(Of Packagehistory)
    Property PackageID() As Integer
        Get
            Return _PackageID
        End Get
        Set(ByVal value As Integer)
            _PackageID = value
        End Set
    End Property
    Property ProductLevelID() As Integer
        Get
            Return _ProductLevelID
        End Get
        Set(ByVal value As Integer)
            _ProductLevelID = value
        End Set
    End Property
    Property ProductCode() As String
        Get
            Return _ProductCode
        End Get
        Set(ByVal value As String)
            _ProductCode = value
        End Set
    End Property
    Property TransactionID() As Integer
        Get
            Return _TransactionID
        End Get
        Set(ByVal value As Integer)
            _TransactionID = value
        End Set
    End Property
    Property ComputerID() As Integer
        Get
            Return _ComputerID
        End Get
        Set(ByVal value As Integer)
            _ComputerID = value
        End Set
    End Property
    Property OrderDetailID() As Integer
        Get
            Return _OrderDetailID
        End Get
        Set(ByVal value As Integer)
            _OrderDetailID = value
        End Set
    End Property
    Property ExpireDate() As DateTime
        Get
            Return _ExpireDate
        End Get
        Set(ByVal value As DateTime)
            _ExpireDate = value
        End Set
    End Property
    Property PurchaseDate() As DateTime
        Get
            Return _PurchaseDate
        End Get
        Set(ByVal value As DateTime)
            _PurchaseDate = value
        End Set
    End Property
    Property MemberID() As Integer
        Get
            Return _MemberID
        End Get
        Set(ByVal value As Integer)
            _MemberID = value
        End Set
    End Property
    Property PackageFirstName() As String
        Get
            Return _PackageFirstName
        End Get
        Set(ByVal value As String)
            _PackageFirstName = value
        End Set
    End Property
    Property PackageLastName() As String
        Get
            Return _PackageLastName
        End Get
        Set(ByVal value As String)
            _PackageLastName = value
        End Set
    End Property
    Property PackageDisplayName() As String
        Get
            Return _PackageDisplayName
        End Get
        Set(ByVal value As String)
            _PackageDisplayName = value
        End Set
    End Property
    Property PackageNumber() As String
        Get
            Return _PackageNumber
        End Get
        Set(ByVal value As String)
            _PackageNumber = value
        End Set
    End Property
    Property PackageNote() As String
        Get
            Return _PackageNote
        End Get
        Set(ByVal value As String)
            _PackageNote = value
        End Set
    End Property
    Property Deleted() As Integer
        Get
            Return _Deleted
        End Get
        Set(ByVal value As Integer)
            _Deleted = value
        End Set
    End Property
    Property UpdateDate() As DateTime
        Get
            Return _UpdateDate
        End Get
        Set(ByVal value As DateTime)
            _UpdateDate = value
        End Set
    End Property
    Property VATType() As Integer
        Get
            Return _VATType
        End Get
        Set(ByVal value As Integer)
            _VATType = value
        End Set
    End Property
    Property PackageProductType() As Integer
        Get
            Return _PackageProductType
        End Get
        Set(ByVal value As Integer)
            _PackageProductType = value
        End Set
    End Property
    Property PackageOriginalAmount() As Integer
        Get
            Return _PackageOriginalAmount
        End Get
        Set(ByVal value As Integer)
            _PackageOriginalAmount = value
        End Set
    End Property
    Property PackageCurrentAmount() As Decimal
        Get
            Return _PackageCurrentAmount
        End Get
        Set(ByVal value As Decimal)
            _PackageCurrentAmount = value
        End Set
    End Property
    Property MaxAllowEditAmount() As Integer
        Get
            Return _MaxAllowEditAmount
        End Get
        Set(ByVal value As Integer)
            _MaxAllowEditAmount = value
        End Set
    End Property
    Property PackagePrice() As Decimal
        Get
            Return _PackagePrice
        End Get
        Set(ByVal value As Decimal)
            _PackagePrice = value
        End Set
    End Property
    Property AllowUseInShop() As String
        Get
            Return _AllowUseInShop
        End Get
        Set(ByVal value As String)
            _AllowUseInShop = value
        End Set
    End Property
    Property PackageHistory() As List(Of Packagehistory)
        Get
            Return _PackageHistory
        End Get
        Set(ByVal value As List(Of Packagehistory))
            _PackageHistory = value
        End Set
    End Property
End Class
Public Class Packagehistory
    Protected _PackageHistoryID As Integer
    Protected _PackageID As Integer
    Protected _ProductLevelID As Integer
    Protected _TransactionID As Integer
    Protected _ComputerID As Integer
    Protected _OrderDetailID As Integer
    Protected _ProductCode As String
    Protected _PackageStatus As Integer
    Protected _CommissionPrice As Decimal
    Protected _CommissionPriceVAT As Decimal
    Protected _AdditionalPrice As Decimal
    Protected _PackageSessionAmount As Decimal
    Protected _Updatedate As DateTime

    Property PackageHistoryID() As Integer
        Get
            Return _PackageHistoryID
        End Get
        Set(ByVal value As Integer)
            _PackageHistoryID = value
        End Set
    End Property
    Property PackageID() As Integer
        Get
            Return _PackageID
        End Get
        Set(ByVal value As Integer)
            _PackageID = value
        End Set
    End Property
    Property ProductLevelID() As Integer
        Get
            Return _ProductLevelID
        End Get
        Set(ByVal value As Integer)
            _ProductLevelID = value
        End Set
    End Property
    Property TransactionID() As Integer
        Get
            Return _TransactionID
        End Get
        Set(ByVal value As Integer)
            _TransactionID = value
        End Set
    End Property
    Property ComputerID() As Integer
        Get
            Return _ComputerID
        End Get
        Set(ByVal value As Integer)
            _ComputerID = value
        End Set
    End Property
    Property OrderDetailID() As Integer
        Get
            Return _OrderDetailID
        End Get
        Set(ByVal value As Integer)
            _OrderDetailID = value
        End Set
    End Property
    Property ProductCode() As String
        Get
            Return _ProductCode
        End Get
        Set(ByVal value As String)
            _ProductCode = value
        End Set
    End Property
    Property PackageStatus() As Integer
        Get
            Return _PackageStatus
        End Get
        Set(ByVal value As Integer)
            _PackageStatus = value
        End Set
    End Property
    Property CommissionPrice() As Decimal
        Get
            Return _CommissionPrice
        End Get
        Set(ByVal value As Decimal)
            _CommissionPrice = value
        End Set
    End Property
    Property CommissionPriceVAT() As Decimal
        Get
            Return _CommissionPriceVAT
        End Get
        Set(ByVal value As Decimal)
            _CommissionPriceVAT = value
        End Set
    End Property
    Property AdditionalPrice() As Decimal
        Get
            Return _AdditionalPrice
        End Get
        Set(ByVal value As Decimal)
            _AdditionalPrice = value
        End Set
    End Property
    Property PackageSessionAmount() As Decimal
        Get
            Return _PackageSessionAmount
        End Get
        Set(ByVal value As Decimal)
            _PackageSessionAmount = value
        End Set
    End Property
    Property Updatedate() As DateTime
        Get
            Return _Updatedate
        End Get
        Set(ByVal value As DateTime)
            _Updatedate = value
        End Set
    End Property
End Class

<Serializable()> _
Public Class SummaryPoint_Data
    Private _memberID As Integer = 0
    Private _TotalPoint As Decimal = 0.0
    Private _UpdatePoint As DateTime
    Property MemberID() As Integer
        Get
            Return _memberID
        End Get
        Set(ByVal value As Integer)
            _memberID = value
        End Set
    End Property
    Property TotalPoint() As Decimal
        Get
            Return _TotalPoint
        End Get
        Set(ByVal value As Decimal)
            _TotalPoint = value
        End Set
    End Property
    Property UpdatePoint() As DateTime
        Get
            Return _UpdatePoint
        End Get
        Set(ByVal value As DateTime)
            _UpdatePoint = value
        End Set
    End Property
End Class
<Serializable()> _
Public Class Softwareversion_Data
    Private _ProgramTypeID As Integer = 0
    Private _IPHQServer As String = ""
    Private _FTPUserName As String = ""
    Private _FTPPassword As String = ""
    Private _UpdateServerPath As String = ""
    Private _FileVersion As String = ""
    Private _LastUpdate As DateTime
    Property ProgramTypeID() As Integer
        Get
            Return _ProgramTypeID
        End Get
        Set(ByVal value As Integer)
            _ProgramTypeID = value
        End Set
    End Property
    Property IPHQServer() As String
        Get
            Return _IPHQServer
        End Get
        Set(ByVal value As String)
            _IPHQServer = value
        End Set
    End Property
    Property FTPUserName() As String
        Get
            Return _FTPUserName
        End Get
        Set(ByVal value As String)
            _FTPUserName = value
        End Set
    End Property
    Property FTPPassword() As String
        Get
            Return _FTPPassword
        End Get
        Set(ByVal value As String)
            _FTPPassword = value
        End Set
    End Property
    Property UpdateServerPath() As String
        Get
            Return _UpdateServerPath
        End Get
        Set(ByVal value As String)
            _UpdateServerPath = value
        End Set
    End Property
    Property FileVersion() As String
        Get
            Return _FileVersion
        End Get
        Set(ByVal value As String)
            _FileVersion = value
        End Set
    End Property
    Property LastUpdate() As DateTime
        Get
            Return _LastUpdate
        End Get
        Set(ByVal value As DateTime)
            _LastUpdate = value
        End Set
    End Property
End Class
<Serializable()> _
Public Class News_ContentData
    Protected _SectionID As Integer = 0
    Protected _CategoryID As Integer = 0
    Protected _ContentID As Integer = 0
    Protected _Status As String = ""
    Protected _Title As String = ""
    Protected _Description As String = ""
    Protected _Detail As String = ""
    Protected _Created As DateTime
    Protected _Modified As DateTime
    Protected _Published As Integer = 0
    Protected _StartPublished As DateTime
    Protected _FinishPublished As DateTime
    Protected _InsertStaffID As Integer = 0
    Protected _UpdateStaffID As Integer = 0
    Property SectionID() As Integer
        Get
            Return _SectionID
        End Get
        Set(ByVal value As Integer)
            _SectionID = value
        End Set
    End Property
    Property CategoryID() As Integer
        Get
            Return _CategoryID
        End Get
        Set(ByVal value As Integer)
            _CategoryID = value
        End Set
    End Property
    Property ContentID() As Integer
        Get
            Return _ContentID
        End Get
        Set(ByVal value As Integer)
            _ContentID = value
        End Set
    End Property
    Property Status() As String
        Get
            Return _Status
        End Get
        Set(ByVal value As String)
            _Status = value
        End Set
    End Property
    Property Title() As String
        Get
            Return _Title
        End Get
        Set(ByVal value As String)
            _Title = value
        End Set
    End Property
    Property Description() As String
        Get
            Return _Description
        End Get
        Set(ByVal value As String)
            _Description = value
        End Set
    End Property
    Property Detail() As String
        Get
            Return _Detail
        End Get
        Set(ByVal value As String)
            _Detail = value
        End Set
    End Property
    Property Created() As DateTime
        Get
            Return _Created
        End Get
        Set(ByVal value As DateTime)
            _Created = value
        End Set
    End Property
    Property Modified() As DateTime
        Get
            Return _Modified
        End Get
        Set(ByVal value As DateTime)
            _Modified = value
        End Set
    End Property
    Property Published() As Integer
        Get
            Return _Published
        End Get
        Set(ByVal value As Integer)
            _Published = value
        End Set
    End Property
    Property StartPublished() As DateTime
        Get
            Return _StartPublished
        End Get
        Set(ByVal value As DateTime)
            _StartPublished = value
        End Set
    End Property
    Property FinishPublished() As DateTime
        Get
            Return _FinishPublished
        End Get
        Set(ByVal value As DateTime)
            _FinishPublished = value
        End Set
    End Property
    Property InsertStaffID() As Integer
        Get
            Return _InsertStaffID
        End Get
        Set(ByVal value As Integer)
            _InsertStaffID = value
        End Set
    End Property
    Property UpdateStaffID() As Integer
        Get
            Return _UpdateStaffID
        End Get
        Set(ByVal value As Integer)
            _UpdateStaffID = value
        End Set
    End Property
End Class
<Serializable()> _
Public Class News_SectionData
    Protected _SectionID As Integer = 0
    Protected _SectionTitle As String = ""
    Protected _SectionDescription As String = ""
    Protected _Published As Integer = ""
    Property SectionID() As Integer
        Get
            Return _SectionID
        End Get
        Set(ByVal value As Integer)
            _SectionID = value
        End Set
    End Property
    Property SectionTitle() As String
        Get
            Return _SectionTitle
        End Get
        Set(ByVal value As String)
            _SectionTitle = value
        End Set
    End Property
    Property SectionDescription() As String
        Get
            Return _SectionDescription
        End Get
        Set(ByVal value As String)
            _SectionDescription = value
        End Set
    End Property
    Property Published() As Integer
        Get
            Return _Published
        End Get
        Set(ByVal value As Integer)
            _Published = value
        End Set
    End Property
End Class
<Serializable()> _
Public Class News_CategoryData
    Protected _SectionID As Integer = 0
    Protected _CategoryID As Integer = 0
    Protected _CategoryTitle As String = ""
    Protected _CategoryDescription As String = ""
    Protected _ContentDetail As List(Of News_ContentData)
    Protected _ContentTotal As Integer
    Property SectionID() As Integer
        Get
            Return _SectionID
        End Get
        Set(ByVal value As Integer)
            _SectionID = value
        End Set
    End Property
    Property CategoryID() As Integer
        Get
            Return _CategoryID
        End Get
        Set(ByVal value As Integer)
            _CategoryID = value
        End Set
    End Property
    Property CategoryTitle() As String
        Get
            Return _CategoryTitle
        End Get
        Set(ByVal value As String)
            _CategoryTitle = value
        End Set
    End Property
    Property CategoryDescription() As String
        Get
            Return _CategoryDescription
        End Get
        Set(ByVal value As String)
            _CategoryDescription = value
        End Set
    End Property
    Property ContentDetail() As List(Of News_ContentData)
        Get
            Return _ContentDetail
        End Get
        Set(ByVal value As List(Of News_ContentData))
            _ContentDetail = value
        End Set
    End Property
    Property ContentTotal() As Integer
        Get
            Return _ContentTotal
        End Get
        Set(ByVal value As Integer)
            _ContentTotal = value
        End Set
    End Property
End Class
 
Public Enum SearchMemberBy
    MemberCode = 0
    MemberMobile = 1
    MemberName = 2
End Enum



