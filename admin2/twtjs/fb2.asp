<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

function clsRemoting()
{
this.addnl = Function('n1','n2','n3','return FeiBiao(n1,n2,n3)');
} 
var public_description = new clsRemoting();
RSDispatch(); %>
<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">
Function FeiBiao(fireds,hits,losthits)
nickname=info(0)
hits=hits * 5
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb") 
rs.open "select ����,���O,���O�[ FROM �Τ� WHERE id="&info(9),conn
sql="update �Τ� set ���O=���O+" & hits-fireds & " ,��O=��O-" & losthits & " where id="&info(9)
nlj=(rs("����")*640+2000)+rs("���O�[")
if rs("���O")<nlj then
conn.execute sql
writestr="�A�m�F�ܤ[�A�ֱo�b���A��O(�ͩR)����" & losthits & "�I,���O����" & fireds &"�I,���O�W�[" & hits & "�I�C" 
else
writestr="�ѩ�A���W��,�A�����O�{�b�L�k����,�٬O�M�A���v�Žm�Z�h�a!"
end if
rs.close
set rs=nothing 
conn.close
set conn=nothing
FeiBiao=writestr
End Function
</Script>
