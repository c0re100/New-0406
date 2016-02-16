<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
my=info(0)
%><!--#include file="dadata.asp"-->
<%
sql="select 性別 from 用戶 where id="&info(9)
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "你不是江湖中人，不能訂購酒宴"
response.end
else
id=request("id")
sex=rs("性別")
if sex="男" then
sex="女"
else
sex="男"	
end if	
%>
<HTML>
<HEAD>
<TITLE>請輸入你要宴請的<%=sex%>朋友的名字</TITLE>
<script language=JavaScript>
document.onmousedown=click;
function click(){if(event.button==2){if(confirm("是否關閉窗口？","江湖提示")){window.close();}}}
</script>
<META http-equiv=Content-Type content="text/html; charset=big5">

<STYLE>TD {
FONT-SIZE: 12px; LINE-HEIGHT: 16px
}
input {border: 1px solid; font-size: 9pt; font-family: "新細明體"; border-color:
#000000 solid

}
</STYLE>
<META content="Microsoft FrontPage 4.0" name=GENERATOR>
</HEAD>
<BODY text=#cccccc vLink=#ffff00 aLink=#ffff00 link=#ffff00 bgColor=#3a4b91
leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE cellSpacing=0 cellPadding=0 width=400 align=left border=0>
<TBODY>
<TR>
<TD vAlign=top height=365>
<TABLE cellSpacing=0 cellPadding=0 width=400 border=0>
<TBODY>
<TR>
<TD vAlign=top height=105>
<P>&nbsp;</P>
<P>     說明：請輸入你要宴請的<%=sex%>朋友的名字。 </P></TD></TR></TBODY></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=360 align=center border=0>
<FORM action="jiudian33.asp" method=post>

<TR>
<TD class=p9>
<div align="center">姓名：
<INPUT size=10 name=name maxlength="10">
<INPUT size=2 readonly type="hidden" name=id maxlength="2" value=<%=id%>>
</div>
</TD></TR>
<TR>
<TD class=p11 height=1></TD></TR>
<TR>
<TD
class=p9>
</TD></TR>
<TR>
<TD class=p11 height=1></TD></TR>
<TR>
<TD class=p9 align=middle>
<div align="center">
<INPUT type=submit value=確認>
<INPUT onclick=window.close() type=button value=關閉>
</div>
</TD></TR></FORM>
</TABLE>
<P></P>

</TD>
</TR></TBODY></TABLE></BODY></HTML>
<%
rs.close
conn.close
set rs=nothing
set conn=nothing
end if%>
