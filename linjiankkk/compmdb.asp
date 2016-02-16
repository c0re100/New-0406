<%@ LANGUAGE=VBScript codepage ="950" %>
<%if not IsArray(Session("info")) then Response.Redirect "error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
Const Jet_Conn_Partial = "Provider=Microsoft.Jet.OLEDB.4.0;Data source="
Dim strDatabase, strFolder, strFileName

strFolder = Server.MapPath("asz.asa")
strFolder = Replace(strFolder,"asz.asa","")
strDatabase = "asz.asa"

Private Sub dbCompact(strDBFileName)
Dim SourceConn
Dim DestConn
Dim oJetEngine
Dim oFSO

SourceConn = Jet_Conn_Partial & strFolder & strDatabase
DestConn = Jet_Conn_Partial & strFolder & "Temp" & strDatabase

Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
Set oJetEngine = Server.CreateObject("JRO.JetEngine")

With oFSO

If Not .FileExists(strFolder & strDatabase) Then
Response.Write ("沒有發現數據庫：" & strFolder & strDatabase)
Stop
Else
If .FileExists(strFolder & "Temp" & strDatabase) Then
Response.Write ("發生錯誤 " _
& "刪除舊的數據庫再試!")
.DeleteFile (strFolder & "Temp" & strDatabase)
End If
End If
End With

With oJetEngine
.CompactDatabase SourceConn, DestConn
End With

oFSO.DeleteFile strFolder & strDatabase
oFSO.MoveFile strFolder & "Temp" _
& strDatabase, strFolder& strDatabase

Set oFSO = Nothing
Set oJetEngine = Nothing
End Sub

Private Sub dbList()
Response.Write ("<Select Name=""DBFileName"">")
Response.Write ("<Option Value="&  Server.MapPath("taishan.asp") &" selected>用戶數據庫</Option>")
Response.Write ("<Option Value="&  Server.MapPath("../bbs/forum.mdb") &">論談數據庫</Option>")
Response.Write ("</Select>")
End Sub
%>
<%
' Compact database and tell the user the database is optimized
Select Case Request.form("cmd")
Case "壓縮"
dbCompact Request.form("DBFileName")
Response.Write ("  <p align='center'>數據庫:" & Request.form("DBFileName") & "壓縮完成！</p>")
End Select
%><title>靈劍江湖維護程序</title>
<body background="../images/8.jpg">
<p align="center"><font size="4"><b><font color="#0000FF">壓縮並修複數據庫</font></b></font></p>
<form method="POST" action="">
<p align="center">
<%dbList%>
<input type="submit" value="壓縮" name="cmd">
</p>
</form>



<p align="center"><font color="#000000">請選擇所壓縮的數據庫，按壓縮開始進行！</font><br>
此套程序需要空間有Fso支持才可以進行，進行此操作時為了避免誤操作<br>
建議備份原始數據庫文件，對於造成的後果台山明明不負任何責任！</p> 