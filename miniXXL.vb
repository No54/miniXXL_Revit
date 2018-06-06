Imports System.Data.SQLite
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports Autodesk.Revit.DB
Imports Autodesk.Revit.DB.Structure
Imports Autodesk.Revit.UI
Imports Autodesk.Revit.UI.Selection

<Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)>
<Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)>
Public Class miniXXL
    Implements IExternalCommand

    Public Function Execute(commandData As ExternalCommandData, ByRef message As String, elements As ElementSet) _
        As Result Implements IExternalCommand.Execute

#Region "定义app,doc,uidoc"

        If (Directory.Exists("C:\miniXXL\") = False) Then
            Directory.CreateDirectory("C:\miniXXL\")
        End If

        Dim app As UIApplication = commandData.Application
        Dim Buidoc As UIDocument = Nothing
        Dim Bdoc As Document = Nothing
        Try
            Buidoc = app.ActiveUIDocument
            Bdoc = Buidoc.Document
        Catch ex As Exception

        End Try


        Dim doc As Document = app.Application.NewProjectDocument(UnitSystem.Metric)
        Dim saveop As SaveAsOptions = New SaveAsOptions()
        saveop.OverwriteExistingFile = True
        doc.SaveAs("C:\miniXXL\MyXXL.rvt", saveop)
        Dim uidoc As UIDocument = app.OpenAndActivateDocument(doc.PathName)
        If Bdoc IsNot Nothing Then
            If Bdoc.PathName = "C:\miniXXL\MyXXLtemp.rvt" Then
                Bdoc.Close(False)
            End If
        End If

#End Region

#Region "收集计算机名及IP"
        Dim CurHost As String = System.Environment.UserName
        Dim CurIP As String = ""

        Dim DNSname As String = Dns.GetHostName()
        Dim ipadrlist() As IPAddress = Dns.GetHostAddresses(DNSname)
        For Each ipa As IPAddress In ipadrlist
            If (ipa.AddressFamily = AddressFamily.InterNetwork) Then
                CurIP = ipa.ToString()
            End If
        Next
#End Region

#Region "加载方块族并激活，创建symbol集合"

        Dim fsRed As FamilySymbol = Nothing
        Dim fsBlue As FamilySymbol = Nothing
        Dim fsYellow As FamilySymbol = Nothing
        Dim fsGreen As FamilySymbol = Nothing

        Dim transLoad As Transaction = New Transaction(doc)
        transLoad.Start("miniXXL加载族")
        Dim loadRed As Boolean = doc.LoadFamilySymbol("\\yun\miniXXL\小红.rfa", "小红", fsRed)
        Dim loadBlue As Boolean = doc.LoadFamilySymbol("\\yun\miniXXL\小蓝.rfa", "小蓝", fsBlue)
        Dim loadYellow As Boolean = doc.LoadFamilySymbol("\\yun\miniXXL\小黄.rfa", "小黄", fsYellow)
        Dim loadGreen As Boolean = doc.LoadFamilySymbol("\\yun\miniXXL\小绿.rfa", "小绿", fsGreen)

        fsRed.Activate()
        fsBlue.Activate()
        fsYellow.Activate()
        fsGreen.Activate()

        transLoad.Commit()

#End Region

#Region "得到方块集合，并打乱"
        Dim fsSet1 As IList(Of FamilySymbol) = New List(Of FamilySymbol)
        Dim fsSet As IList(Of FamilySymbol) = New List(Of FamilySymbol)
        For i = 1 To 16
            fsSet1.Add(fsRed)
            fsSet1.Add(fsBlue)
            fsSet1.Add(fsYellow)
            fsSet1.Add(fsGreen)
        Next
        Dim myRND As New Random
        For i = 0 To 63
            Dim RNDnum As Integer = myRND.Next(0, fsSet1.Count)
            fsSet.Add(fsSet1.ElementAt(RNDnum))
            fsSet1.RemoveAt(RNDnum)
        Next
#End Region

#Region "计算所有的点xyz"
        Dim px As Double = 0
        Dim py As Double = 0
        Dim pz As Double = 0

        Dim pp As IList(Of XYZ) = New List(Of XYZ)
        For ix = 0 To 3
            For iy = 0 To 3
                For iz = 0 To 3
                    Dim newpp As XYZ = New XYZ(ix * 1000 / 304.8, iy * 1000 / 304.8, iz * 1000 / 304.8)
                    pp.Add(newpp)
                Next
            Next
        Next
#End Region

#Region "创建大方块"
        Dim InsIdSet As IList(Of ElementId) = New List(Of ElementId)


        Dim transIns As Transaction = New Transaction(doc)
        transIns.Start("miniXXL创建模型")
        For i = 0 To 63

            Dim Ins As FamilyInstance = doc.Create.NewFamilyInstance(pp.ElementAt(i), fsSet.ElementAt(i), StructuralType.NonStructural)
            InsIdSet.Add(Ins.Id)
        Next
        transIns.Commit()

#End Region

#Region "操作视图"
        Dim V3D As View3D = newV3D(uidoc)

        Dim transViewM As Transaction = New Transaction(doc)
        transViewM.Start("miniXXL创建模型")
        uidoc.ActiveView.DetailLevel = ViewDetailLevel.Fine
        uidoc.ActiveView.DisplayStyle = DisplayStyle.Realistic
        transViewM.Commit()
#End Region


        MsgBox("点击两个相同颜色的方块，将方块消除，看看你用多少时间消除全部方块。" & vbLf & "点击确定游戏开始！按Exc键退出！")

        Dim dt1 As DateTime = DateTime.Now

        '此处加入计时

        Dim CubeFilter As New ElementClassFilter(GetType(FamilyInstance))
        Dim CubeCollector As New FilteredElementCollector(doc, uidoc.ActiveView.Id)
        CubeCollector.WherePasses(CubeFilter)

        Dim query As IEnumerable(Of Element) = From element In CubeCollector
                                               Select element

       

        Do
            Try
                Dim SelA As Selection = uidoc.Selection
                Dim refA As Reference = SelA.PickObject(ObjectType.Element)
                Dim elemA As Element = Nothing
                If refA IsNot Nothing Then
                    elemA = doc.GetElement(refA)
                End If
                Dim SelB As Selection = uidoc.Selection
                Dim refB As Reference = SelB.PickObject(ObjectType.Element)
                Dim elemB As Element = Nothing
                If refB IsNot Nothing Then
                    elemB = doc.GetElement(refB)
                End If
                If elemA IsNot Nothing And
                   elemB IsNot Nothing And
                   elemA.Name.ToString = elemB.Name.ToString And
                   elemA.Id.ToString <> elemB.Id.ToString Then

                    Dim transDel As Transaction = New Transaction(doc)
                    transDel.Start("miniXXL消消乐")
                    InsIdSet.Remove(elemA.Id)
                    InsIdSet.Remove(elemB.Id)
                    uidoc.ActiveView.HideElements({elemA.Id, elemB.Id})
 
                    transDel.Commit()

                End If
            Catch ex As Exception
                If TryCast(ex, Autodesk.Revit.Exceptions.OperationCanceledException) IsNot Nothing Then
                    Exit Do

                Else
                    MsgBox(ex.ToString)
                End If
            End Try
            'MsgBox(InsIdSet.Count.ToString)
        Loop While InsIdSet.Count > 0


        If InsIdSet.Count = 0 Then
            Dim dt2 As DateTime = DateTime.Now
            Dim CurScore As String = ((dt2 - dt1).TotalSeconds).ToString()
            Dim CurDate As String = dt2.ToLongDateString().ToString()
            Dim CurTime As String = dt2.ToLongTimeString().ToString()
            Dim CurPP As String = ""                                               '用来记录排位

            Try
                Dim sourceFile As String = "\\yun\miniXXL\Score\SCORE.db"
                Dim destinationFile As String = "C:\miniXXL\SCORE.db"
                File.Copy(sourceFile, destinationFile, True)
            Catch ex As Exception

            End Try


            Dim path As String = "C:\miniXXL\SCORE.db"
            Dim SqliteConn As SQLiteConnection = New SQLiteConnection("data source=" & path)
            If (SqliteConn.State <> System.Data.ConnectionState.Open) Then
                SqliteConn.Open()
            End If
            Dim dbcmd As SQLiteCommand = New SQLiteCommand()
            dbcmd.Connection = SqliteConn

dbcmd.CommandText = "INSERT INTO Score_Total (CurHost, CurIP, CurDATE,CurTIME,CurSCORE) VALUES ('" & _
CurHost & "', '" & CurIP & "', '" & CurDate & "', '" & CurTime & "', '" & CDbl(CurScore) & "')"

            dbcmd.ExecuteNonQuery()

            dbcmd.CommandText = "SELECT CurHost, CurSCORE FROM Score_Total order by CurSCORE"

            Dim sqlSR As SQLiteDataReader = dbcmd.ExecuteReader()
            Dim dataT As DataTable = New DataTable()

            Dim nameset As IList(Of String) = New List(Of String)
            Dim scoreset As IList(Of Double) = New List(Of Double)

            If sqlSR.HasRows Then
                dataT.Load(sqlSR)
                For i = 0 To 2
                    nameset.Add(dataT.Rows(i).Item(0).ToString)
                    scoreset.Add(dataT.Rows(i).Item(1).ToString)
                Next

                Dim j As Integer = 0
                Do
                    If dataT.Rows(j).Item(1).ToString = CurScore Then
                        Exit Do
                    Else
                        j = j + 1
                    End If
                Loop While j < dataT.Rows.Count
                CurPP = j

            End If

            Dim maxLen As Integer = longlen(nameset)

            For Each name As String In nameset
                name = name + Space(maxLen - Len(name))
            Next

            sqlSR.Close()
            SqliteConn.Close()

            MsgBox("游戏完成！共用时" & CurScore & "秒！" & vbLf &
               "目前排名第" & CurPP + 1 & "位，排名前三位的是：" & vbLf &
               "1" & vbTab & nameset.ElementAt(0) & vbTab & scoreset.ElementAt(0) & "秒" & vbLf &
               "2" & vbTab & nameset.ElementAt(1) & vbTab & scoreset.ElementAt(1) & "秒" & vbLf &
               "3" & vbTab & nameset.ElementAt(2) & vbTab & scoreset.ElementAt(2) & "秒")

            Try
                Dim sourceFile As String = "C:\miniXXL\SCORE.db"
                Dim destinationFile As String = "\\yun\miniXXL\Score\SCORE.db"
                File.Copy(sourceFile, destinationFile, True)
            Catch ex As Exception

            End Try


        Else
            MsgBox("游戏中断！")
        End If


        doc.SaveAs("C:\miniXXL\MyXXLtemp.rvt", saveop)


        Return Result.Succeeded
    End Function

    Private Function newV3D(uidoc As UIDocument) As View3D

        Dim transNewV3D As Transaction = New Transaction(uidoc.Document)
        transNewV3D.Start("miniXXL创建视图")

        Dim collector1 As New FilteredElementCollector(uidoc.Document)
        collector1 = collector1.OfClass(GetType(ViewFamilyType))
        Dim viewFamilyTypes As IEnumerable(Of ViewFamilyType)

        viewFamilyTypes = From elem In collector1
                          Let vftype = TryCast(elem, ViewFamilyType)
                          Where vftype.ViewFamily = ViewFamily.ThreeDimensional
                          Select vftype
       
        newV3D = View3D.CreateIsometric(uidoc.Document, viewFamilyTypes.First().Id)

        transNewV3D.Commit()

        uidoc.ActiveView = newV3D

        Return newV3D

    End Function

    Private Function longlen(strset As IList(Of String)) As Integer
        longlen = 1

        For i = 0 To strset.Count - 2
            If Len(strset.ElementAt(i)) > Len(strset.ElementAt(i + 1)) Then
                longlen = Len(strset.ElementAt(i))
            Else
                longlen = Len(strset.ElementAt(i + 1))
            End If
        Next

        Return longlen

    End Function

End Class
