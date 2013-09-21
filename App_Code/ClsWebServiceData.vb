Imports Microsoft.VisualBasic
Imports MySql.Data.MySqlClient
Imports POSMySQL.POSControl
Imports System.Data
Imports System.IO
Imports System.Collections.Generic
Imports System.Net

<Serializable()> _
Public Class ClsWebServiceData
    Shared Function MemberData(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal searchby As SearchMemberBy, ByVal memberCode As String, ByVal memberPhone As String) As DataSet
        Dim syn As String = ""
        Dim addional As String = ""
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim ds As New DataSet

        Select Case searchby
            Case Is = SearchMemberBy.MemberCode
                addional &= " And b.MemberCode='" & memberCode & "'"
            Case Is = SearchMemberBy.MemberMobile
                addional &= " And b.MemberMobile='" & memberPhone & "'"
            Case Else
                addional &= " And b.MemberID=0"
        End Select
        syn = "Select a.MemberGroupID,a.MemberGroupCode,a.MemberGroupName,b.MemberID,b.MemberCode,b.MemberGender,b.MemberFirstName,b.MemberLastName,"
        syn &= "b.MemberAddress1, b.MemberAddress2,b.MemberCity,b.MemberProvince,b.MemberZipCode,b.MemberFax,b.MemberEmail,b.MemberBirthDay,b.Activated,b.MemberExpiration,"
        syn &= "b.InputDate,b.LastUseDate, b.MemberTelephone, b.MemberMobile,b.UpDateDate as MemberUpdateDate,b.UpDateBy, c.TotalPoint, c.UpdateDate, b.MemberAdditional"
        syn &= " From MemberGroup a Inner Join	 Members b On a.MemberGroupID=b.MemberGroupID " & addional
        syn &= " Left Join RewardPointSummary c On b.MemberID=c.MemberID"
        syn &= " Where a.Deleted = 0 And b.Deleted=0 And b.Activated=1"
        syn &= " Limit 1"
        dt1 = objDB.List(syn, objCnn)
        dt1.TableName = "MemberData"

        syn = "Select * From Members b Where b.Deleted=0 " & addional
        syn &= " Order by b.MemberCode"
        syn &= " Limit 1"
        dt2 = objDB.List(syn, objCnn)
        dt2.TableName = "Members"

        ds.Tables.Add(dt1)
        ds.Tables.Add(dt2)

        Return ds
    End Function
    Shared Function PackageDetailData(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal memberID As Integer) As DataSet
        Dim dtPackageDetail As New DataTable
        Dim dtPackageHistory As New DataTable
        Dim dsData As New DataSet

        Dim syn As String = ""
        syn = "SELECT * FROM packagedetail WHERE MemberID=" & memberID
        dtPackageDetail = objDB.List(syn, objCnn)
        dtPackageDetail.TableName = "packagedetail"

        syn = "SELECT b.* FROM packagedetail a INNER JOIN packagehistory b ON a.PackageID=b.PackageID AND a.ProductLevelID=b.ProductLevelID " & _
              " WHERE a.MemberID=" & memberID & _
              " Order by a.PackageID,b.PackageStatus,b.ProductCode"
        dtPackageHistory = objDB.List(syn, objCnn)
        dtPackageHistory.TableName = "packagehistory"

        dsData.Tables.Add(dtPackageDetail)
        dsData.Tables.Add(dtPackageHistory)

        Return dsData
    End Function
    Shared Function Member_GenerateScriptUpdatePackage(ByVal objPackHistory As List(Of Packagehistory)) As String
        Dim strSQL As String = ""
        If objPackHistory.Count > 0 Then
            For i As Integer = 0 To objPackHistory.Count - 1
                strSQL &= "update packagehistory set PackageStatus=1,transactionid=" & objPackHistory(i).TransactionID & ",computerid=" & objPackHistory(i).ComputerID & ",orderdetailid=" & objPackHistory(i).OrderDetailID & ",commissionpricevat=" & objPackHistory(i).CommissionPriceVAT & ",additionalprice=" & objPackHistory(i).AdditionalPrice & ",packagesessionamount=" & objPackHistory(i).PackageSessionAmount & ",updatedate=" & DbUtilsModule.DbUtils.FormatDateTime(objPackHistory(i).Updatedate) & " where PackageHistoryID=" & objPackHistory(i).PackageHistoryID & " and packageid=" & objPackHistory(i).PackageID & " and productlevelid=" & objPackHistory(i).ProductLevelID & " ;"
            Next
        End If
        Return strSQL
    End Function
    Shared Function MemberDataList(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal searchBy As Integer, ByVal paramSearch As String) As DataSet
        Dim syn As String = ""
        Dim addional As String = ""
        Dim dt1 As New DataTable
        Dim dt2 As New DataTable
        Dim ds As New DataSet

        Select Case searchBy
            Case Is = SearchMemberBy.MemberCode
                If paramSearch <> "" Then
                    addional &= " And b.MemberCode like '%" & paramSearch.Trim("") & "%'"
                End If
            Case Is = SearchMemberBy.MemberMobile
                If paramSearch <> "" Then
                    addional &= " And b.MemberMobile like '%" & paramSearch.Trim("") & "%'"
                End If
            Case Is = SearchMemberBy.MemberName
                If paramSearch <> "" Then
                    addional &= " And concat(b.memberfirstname,b.memberlastname) like '%" & paramSearch.Trim("") & "%'"
                End If
        End Select

        syn = "Select a.MemberGroupID,a.MemberGroupCode,a.MemberGroupName,b.MemberID,b.MemberCode,b.MemberGender,b.MemberFirstName,b.MemberLastName,"
        syn &= "b.MemberAddress1, b.MemberAddress2,b.MemberCity,b.MemberProvince,b.MemberZipCode,b.MemberFax,b.MemberEmail,b.MemberBirthDay,b.Activated,b.MemberExpiration,"
        syn &= "b.InputDate,b.LastUseDate, b.MemberTelephone, b.MemberMobile,b.UpDateDate as MemberUpdateDate,b.UpDateBy, c.TotalPoint, c.UpdateDate, b.MemberAdditional"
        syn &= " From MemberGroup a Inner Join	 Members b On a.MemberGroupID=b.MemberGroupID " & addional
        syn &= " Left Join RewardPointSummary c On b.MemberID=c.MemberID"
        syn &= " Where a.Deleted = 0 And b.Deleted=0 And b.Activated=1"
        syn &= " Order by b.MemberCode"
        syn &= " Limit 50"
        dt1 = objDB.List(syn, objCnn)
        dt1.TableName = "MemberData"
        syn = "Select * From Members b Where b.Deleted=0 " & addional
        syn &= " Order by b.MemberCode"
        syn &= " Limit 50"
        dt2 = objDB.List(syn, objCnn)
        dt2.TableName = "Members"
        ds.Tables.Add(dt1)
        ds.Tables.Add(dt2)
        Return ds
    End Function
    Shared Function GetPointMemberData(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal memberID As Integer) As DataTable
        Dim syn As String = ""
        syn = "Select a.MemberGroupID,a.MemberGroupCode,a.MemberGroupName,b.MemberID,b.MemberCode,b.MemberGender,b.MemberFirstName,b.MemberLastName,"
        syn &= "b.MemberAddress1, b.MemberAddress2,b.MemberCity,b.MemberProvince,b.MemberZipCode,b.MemberFax,b.MemberEmail,b.MemberBirthDay,b.Activated,b.MemberExpiration,"
        syn &= "b.InputDate,b.LastUseDate, b.MemberTelephone, b.MemberMobile,b.UpDateDate as MemberUpdateDate,b.UpDateBy, c.TotalPoint, c.UpdateDate, b.MemberAdditional"
        syn &= " From MemberGroup a Inner Join	 Members b On a.MemberGroupID=b.MemberGroupID And b.MemberID=" & memberID
        syn &= " Left Join RewardPointSummary c On b.MemberID=c.MemberID"
        syn &= " Where a.Deleted = 0 And b.Deleted=0 And b.Activated=1"
        syn &= " Limit 1"
        Return objDB.List(syn, objCnn)
    End Function
    Shared Function UpdateRewardPointSummary(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal memberID As Integer, ByVal rewardpointSummary As Decimal, ByRef strResultText As String) As Boolean
        Try
            Dim syn As String = "Update RewardPointSummary Set TotalPoint=" & rewardpointSummary & " , UpdateDate=" & DbUtilsModule.DbUtils.FormatDateTime(Date.Now) & _
                                " Where MemberID=" & memberID
            objDB.sqlExecute(syn, objCnn)
            strResultText = ""
            Return True
        Catch ex As Exception
            strResultText = ex.Message.ToString()
            Return False
        End Try
    End Function
    Shared Function UpdateSoftwareVersionAtHQ(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal ComputerID As Integer, ByVal ProductLevelID As Integer, ByVal IPAddress As String, ByVal FrontVersion As String, ByVal FrontFileDate As String, ByVal FrontUpdateDate As String, ByVal backOfficeVersion As String, ByVal backOfficeFileDate As String, ByVal backOfficeUpdateDate As String, _
                                              ByVal InvVersion As String, ByVal InvFileDate As String, ByVal InvUpdateDate As String, ByRef strResultText As String) As Boolean
        Dim strSQL As String = ""
        Dim strUpdateSQL As String = ""
        Dim dt As DataTable
        strSQL = "INSERT INTO Softwareversion(ComputerID,ProductLevelID,IPAddress,FrontVersion,FrontFileDate,FrontUpdateDate,PocketVersion,PocketFileDate,PocketUpdateDate,InvVersion,InvFileDate,InvUpdateDate)"
        strSQL &= "VALUES(" & ComputerID & "," & ProductLevelID & ",'" & IPAddress & "','" & FrontVersion & "'," & FrontFileDate & "," & FrontUpdateDate
        strSQL &= ",'" & backOfficeVersion & "'," & backOfficeFileDate & "," & backOfficeUpdateDate & ",'" & InvVersion & "'," & InvFileDate & "," & InvUpdateDate & ")"

        strUpdateSQL = "UPDATE Softwareversion SET ComputerID =" & ComputerID
        strUpdateSQL &= ",ProductLevelID=" & ProductLevelID
        strUpdateSQL &= ",IPAddress='" & IPAddress & "'"
        strUpdateSQL &= ",FrontVersion='" & FrontVersion & "'"
        strUpdateSQL &= ",FrontFileDate=" & FrontFileDate
        strUpdateSQL &= ",FrontUpdateDate=" & FrontUpdateDate
        strUpdateSQL &= ",PocketVersion='" & backOfficeVersion & "'"
        strUpdateSQL &= ",PocketFileDate=" & backOfficeFileDate
        strUpdateSQL &= ",PocketUpdateDate=" & backOfficeUpdateDate
        strUpdateSQL &= ",InvVersion='" & InvVersion & "'"
        strUpdateSQL &= ",InvFileDate=" & InvFileDate
        strUpdateSQL &= ",InvUpdateDate=" & InvUpdateDate
        strUpdateSQL &= " WHERE ComputerID=" & ComputerID & " AND ProductLevelID=" & ProductLevelID
        Try
            dt = objDB.List("SELECT * FROM Softwareversion WHERE ComputerID =" & ComputerID, objCnn)
            If dt.Rows.Count > 0 Then
                objDB.sqlExecute(strUpdateSQL, objCnn)
            Else
                objDB.sqlExecute(strSQL, objCnn)
            End If
            strResultText = ""
            Return True
        Catch ex As Exception
            strResultText = ex.ToString()
            Return False
        End Try
    End Function
    Shared Function GetSoftwareVersion(ByVal objDB As CDBUtil, ByVal objCnn As MySqlConnection, ByVal programTypeID As Integer) As DataTable
        Dim strSQL As String = ""
        strSQL = "SELECT ProgramTypeID,IPHQServer,FTPUserName,FTPPassword,UpdateServerPath,FileVersion,LastUpdate FROM LiveUpdateConfig Where ProgramTypeID=" & programTypeID
        Return objDB.List(strSQL, objCnn)
    End Function
    Shared Function News_ContentData(ByVal dbutil As CDBUtil, ByVal dbcnn As MySqlConnection, _
                                  ByVal sectionID As Integer, ByVal categoryID As Integer, ByVal contentID As Integer, ByVal activate As Integer, ByVal shopID As Integer, ByVal limitContent As Integer) As DataTable
        Dim cmd As String
        Dim dt As New DataTable
        Dim dtResult As New DataTable
        Dim expression As String = ""
        Dim str As String = ""
        Dim additional As String = ""
        Dim listShop As String = ""
        Dim dtShop As DataTable
        dtShop = dbutil.List("select * from feedlink where shopid=" & shopID, dbcnn)
        If dtShop.Rows.Count > 0 Then
            For i As Integer = 0 To dtShop.Rows.Count - 1
                If listShop = "" Then
                    listShop = dtShop.Rows(i)("ContentID")
                Else
                    listShop &= "," & dtShop.Rows(i)("ContentID")
                End If
            Next
        End If
        additional = " And fc.ContentID in(" & listShop & ")"
        Select Case activate
            Case Is = -1
                dt = dbutil.List("select * from feedcontent where finishpublished <" & DbUtilsModule.DbUtils.FormatDate(Date.Now), dbcnn)
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        If str = "" Then
                            str = dt.Rows(i)("ContentID")
                        Else
                            str &= "," & dt.Rows(i)("ContentID")
                        End If
                    Next
                End If

                cmd = "Select  case fc.published when 1 then 'กดเพื่องดเผยแพร่' when 0 then 'กดเพื่อเผยแพร่' end as Status,fc.published,fc.Startpublished,fc.finishpublished," & _
                      "fc.ContentID,fc.Title,fc.Description,fc.Detail,fc.Created,fc.Modified,fc.InsertStaffID,fc.UpdateStaffID,fs.SectionID,fg.CategoryID" & _
                      " From feedcontent fc,feedsection fs, feedCategory fg" & _
                      " Where fc.SectionID=fs.SectionID" & _
                      " And fc.CategoryID=fg.CategoryID" & _
                      " And fc.Deleted=0 And fs.Deleted=0 And fg.Deleted=0" & additional
                If sectionID > -1 Then
                    cmd &= " And fs.SectionID=" & sectionID
                End If
                If categoryID > -1 Then
                    cmd &= " And fg.CategoryID=" & categoryID
                End If
                If contentID <> -1 Then
                    cmd &= " And fc.ContentID=" & contentID
                End If
                cmd &= " Order by fc.Created DESC"
                If limitContent > 0 Then
                    cmd &= " limit " & limitContent
                End If
                dtResult = dbutil.List(cmd, dbcnn)
            Case Is = 1
                dt = dbutil.List("select * from feedcontent where finishpublished <" & DbUtilsModule.DbUtils.FormatDate(Date.Now), dbcnn)
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        If str = "" Then
                            str = dt.Rows(i)("ContentID")
                        Else
                            str &= "," & dt.Rows(i)("ContentID")
                        End If
                    Next
                End If

                cmd = "Select  case fc.published when 1 then 'กดเพื่องดเผยแพร่' when 0 then 'กดเพื่อเผยแพร่' end as Status,fc.published,fc.Startpublished,fc.finishpublished," & _
                      "fc.ContentID,fc.Title,fc.Description,fc.Detail,fc.Created,fc.Modified,fc.InsertStaffID,fc.UpdateStaffID,fs.SectionID,fg.CategoryID" & _
                      " From feedcontent fc,feedsection fs, feedCategory fg" & _
                      " Where fc.SectionID=fs.SectionID" & _
                      " And fc.CategoryID=fg.CategoryID" & _
                      " And fc.Deleted=0 And fs.Deleted=0 And fg.Deleted=0" & additional
                If sectionID > -1 Then
                    cmd &= " And fs.SectionID=" & sectionID
                End If
                If categoryID > -1 Then
                    cmd &= " And fg.CategoryID=" & categoryID
                End If
                If contentID <> -1 Then
                    cmd &= " And fc.ContentID=" & contentID
                End If
                cmd &= " and fc.published=1"
                If str <> "" Then
                    cmd &= " and fc.contentid not in(" & str & ")"
                End If
                cmd &= " Order by fc.Created DESC"
                If limitContent > 0 Then
                    cmd &= " limit " & limitContent
                End If
                dtResult = dbutil.List(cmd, dbcnn)
            Case Is = 2
                dt = dbutil.List("select * from feedcontent where finishpublished <" & DbUtilsModule.DbUtils.FormatDate(Date.Now), dbcnn)
                If dt.Rows.Count > 0 Then
                    For i As Integer = 0 To dt.Rows.Count - 1
                        If str = "" Then
                            str = dt.Rows(i)("ContentID")
                        Else
                            str &= "," & dt.Rows(i)("ContentID")
                        End If
                    Next
                End If
                cmd = "Select  case fc.published when 1 then 'กดเพื่องดเผยแพร่' when 0 then 'กดเพื่อเผยแพร่' end as Status,fc.published,fc.Startpublished,fc.finishpublished," & _
                      "fc.ContentID,fc.Title,fc.Description,fc.Detail,fc.Created,fc.Modified,fc.InsertStaffID,fc.UpdateStaffID,fs.SectionID,fg.CategoryID" & _
                      " From feedcontent fc,feedsection fs, feedCategory fg" & _
                      " Where fc.SectionID=fs.SectionID" & _
                      " And fc.CategoryID=fg.CategoryID" & _
                      " And fc.Deleted=0 And fs.Deleted=0 And fg.Deleted=0" & additional
                If sectionID > -1 Then
                    cmd &= " And fs.SectionID=" & sectionID
                End If
                If categoryID > -1 Then
                    cmd &= " And fg.CategoryID=" & categoryID
                End If
                cmd &= " and fc.published=0"
                If str <> "" Then
                    cmd &= " and fc.contentid not in(" & str & ")"
                End If
                If contentID <> -1 Then
                    cmd &= " And fc.ContentID=" & contentID
                End If
                cmd &= " Order by fc.Created DESC"
                If limitContent > 0 Then
                    cmd &= " limit " & limitContent
                End If
                dtResult = dbutil.List(cmd, dbcnn)
            Case Is = 3
                cmd = "Select  case fc.published when 1 then 'กดเพื่องดเผยแพร่' when 0 then 'กดเพื่อเผยแพร่' end as Status,fc.published,fc.Startpublished,fc.finishpublished," & _
                    "fc.ContentID,fc.Title,fc.Description,fc.Detail,fc.Created,fc.Modified,fc.InsertStaffID,fc.UpdateStaffID,fs.SectionID,fg.CategoryID" & _
                    " From feedcontent fc,feedsection fs, feedCategory fg" & _
                    " Where fc.SectionID=fs.SectionID" & _
                    " And fc.CategoryID=fg.CategoryID" & _
                    " And fc.Deleted=0 And fs.Deleted=0 And fg.Deleted=0" & additional
                If sectionID > -1 Then
                    cmd &= " And fs.SectionID=" & sectionID
                End If
                If categoryID > -1 Then
                    cmd &= " And fg.CategoryID=" & categoryID
                End If
                If contentID <> -1 Then
                    cmd &= " And fc.ContentID=" & contentID
                End If
                cmd &= " and fc.finishpublished <" & DbUtilsModule.DbUtils.FormatDate(Date.Now)
                cmd &= " Order by fc.Created DESC"
                If limitContent > 0 Then
                    cmd &= " limit " & limitContent
                End If
                dtResult = dbutil.List(cmd, dbcnn)
        End Select

        Return dtResult
    End Function
    Shared Function News_SectionData(ByVal dbutil As CDBUtil, ByVal dbcnn As MySqlConnection, ByVal SectionID As Integer) As DataTable
        Dim cmd As String
        cmd = "Select SectionID,SectionTitle,SectionDescription," & _
              "Case published When 1 then 'เผยแพร่' When 0 then 'งดแพร่' End as Published" & _
              " From feedsection Where Deleted=0"
        If SectionID <> -1 Then
            cmd &= " And SectionID=" & SectionID
        End If
        Return dbutil.List(cmd, dbcnn)
    End Function
    Shared Function News_CategoryData(ByVal dbutil As CDBUtil, ByVal dbcnn As MySqlConnection, ByVal CategoryID As Integer, ByVal SectionID As Integer) As DataTable
        Dim cmd As String
        cmd = "Select fc.SectionID,fc.CategoryID,fc.CategoryTitle,fc.CategoryDescription From feedcategory fc,feedsection fs" & _
              " Where fc.SectionID=fs.SectionID And fc.Deleted=0 And fs.Deleted=0"
        If CategoryID <> -1 Then
            cmd &= " And fc.CategoryID = " & CategoryID
        End If
        If SectionID <> -1 Then
            cmd &= " And fs.SectionID=" & SectionID
        End If
        Return dbutil.List(cmd, dbcnn)
    End Function
   

End Class
