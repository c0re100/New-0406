<%@ LANGUAGE=JScript%>
<%Response.charset="big5"%>
<% 
function clsRemoting()
{
this.addnl = Function('n1','n2','return FeiBiao(n1,n2)');
} 
var public_description = new clsRemoting();
RSDispatch(); %>
<!--#INCLUDE file="../_scriptlibrary/rs.asp"-->
<SCRIPT RUNAT=SERVER LANGUAGE="VBScript">
Function FeiBiao(points,missed)
nickname=Session("jiutian_name")
if missed<>1 then 
missed=10
tp2="�A�^�F�@���A"
else
tp2="�����A�A�o�ӦӦ^,"
end if
if points>100 then 
points=100
tp="�I,�o��F" & points & "��Ȥl,��ʩx���Q�V�B�F�ǰ���100��C"
elseif points<3 then
points=0
tp="�I,�o��F" & points & "��Ȥl,�ݤW�h�A�O�զ��F�C"
else
tp="�I,�o��F" & points & "��Ȥl,�i�H�n�n�𮧤@�U�F�C"
end if
connstr=Application("hg_connstr")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr
sql="update �Τ� set �Ȩ�=�Ȩ�+" & points & " ,��O=��O-" & missed & " where �m�W='" & nickname & "'"
conn.execute sql
writestr=tp2 & "��O(�ͩR)����" & missed & tp
conn.close
set conn=nothing
FeiBiao=writestr
End Function
</Script>
