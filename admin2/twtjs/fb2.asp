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
rs.open "select 等級,內力,內力加 FROM 用戶 WHERE id="&info(9),conn
sql="update 用戶 set 內力=內力+" & hits-fireds & " ,體力=體力-" & losthits & " where id="&info(9)
nlj=(rs("等級")*640+2000)+rs("內力加")
if rs("內力")<nlj then
conn.execute sql
writestr="你練了很久，累得半死，體力(生命)消耗" & losthits & "點,內力消耗" & fireds &"點,內力增加" & hits & "點。" 
else
writestr="由於你的上限,你的內力現在無法提升,還是和你的師傅練武去吧!"
end if
rs.close
set rs=nothing 
conn.close
set conn=nothing
FeiBiao=writestr
End Function
</Script>
