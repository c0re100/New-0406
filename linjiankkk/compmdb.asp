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
Response.Write ("�S���o�{�ƾڮw�G" & strFolder & strDatabase)
Stop
Else
If .FileExists(strFolder & "Temp" & strDatabase) Then
Response.Write ("�o�Ϳ��~ " _
& "�R���ª��ƾڮw�A��!")
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
Response.Write ("<Option Value="&  Server.MapPath("taishan.asp") &" selected>�Τ�ƾڮw</Option>")
Response.Write ("<Option Value="&  Server.MapPath("../bbs/forum.mdb") &">�׽ͼƾڮw</Option>")
Response.Write ("</Select>")
End Sub
%>
<%
' Compact database and tell the user the database is optimized
Select Case Request.form("cmd")
Case "���Y"
dbCompact Request.form("DBFileName")
Response.Write ("  <p align='center'>�ƾڮw:" & Request.form("DBFileName") & "���Y�����I</p>")
End Select
%><title>�F�C������@�{��</title>
<body background="../images/8.jpg">
<p align="center"><font size="4"><b><font color="#0000FF">���Y�í׽Ƽƾڮw</font></b></font></p>
<form method="POST" action="">
<p align="center">
<%dbList%>
<input type="submit" value="���Y" name="cmd">
</p>
</form>



<p align="center"><font color="#000000">�п�ܩ����Y���ƾڮw�A�����Y�}�l�i��I</font><br>
���M�{�ǻݭn�Ŷ���Fso����~�i�H�i��A�i�榹�ާ@�ɬ��F�קK�~�ާ@<br>
��ĳ�ƥ���l�ƾڮw���A���y������G�x�s�������t����d���I</p> 