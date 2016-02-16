<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%><html><head><title>圖片</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+"[IMG]"+list+"[/IMG]";parent.f2.document.af.sytemp.focus();}</script>
<style TYPE="text/css">
td{line-height:110%;font-size:20pt;}
.p9{line-height:150%;font-size:9pt;}
A{color:FFFFFF;text-decoration:none;}
A:Visited{color:FFFFFF;}
A:Active{color:FFFF00;}
A:Hover{color:Black;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" text="#FFFFFF">
<a name="1"></a> 
<table border="1" align="center" cellpadding="3" bordercolor="#CCCCCC" width="110">
<tr><td><font color="#FFFF00"><div align=center class=p9>貼圖</div></font></td>
</tr>
<tr><td colspan="3"><div align=center class=p9><a href=javascript:history.go(0)>刷新</a></div></td></tr>
</table><%For t=1 to 246%>
<a href=javascript:s("<%=t%>.gif")><img border=0 src="pic/<%=t%>.gif"</a>
<%next%><table border="1" align="center" cellpadding="3" bordercolor="#CCCCCC" width="110"><tr><td><div align=center class=p9><a href=#1>返回頁首</a></div></td></tr></table>
</body></html> 