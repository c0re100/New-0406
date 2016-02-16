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
tp2="你跌了一跤，"
else
tp2="哈哈，你得勝而回,"
end if
if points>100 then 
points=100
tp="點,得到了" & points & "兩銀子,驚動官府被敲詐了些隻剩100兩。"
elseif points<3 then
points=0
tp="點,得到了" & points & "兩銀子,看上去你是白忙了。"
else
tp="點,得到了" & points & "兩銀子,可以好好休息一下了。"
end if
connstr=Application("hg_connstr")
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr
sql="update 用戶 set 銀兩=銀兩+" & points & " ,體力=體力-" & missed & " where 姓名='" & nickname & "'"
conn.execute sql
writestr=tp2 & "體力(生命)消耗" & missed & tp
conn.close
set conn=nothing
FeiBiao=writestr
End Function
</Script>
