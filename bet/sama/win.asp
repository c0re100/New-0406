<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
horsecalc=Request.QueryString("calc")
horsewin=Request.QueryString("win")
bet=Request.QueryString("bet")
msg="<head><link rel=stylesheet href='css.css'></head><body bgcolor='"&bgcolor&"' background='"&bgimage&"'><table width=50% align=center><tr><td>�@�U�`�G"&horsecalc&"</td><td>��ӡG"&horsewin&"����</td><td>�U�`�G"&bet&"</td><td>"
win=horsecalc-bet*10
if win>0 then
msg=msg&"��G"&win
else
msg=msg&"Ĺ�G"&abs(win)
end if	
msg=msg&"</td></tr><tr><td colspan=6 align=center><input type=button onclick=javascript:location.replace('chipin.asp'); value=' �� �^ '></td></tr></table></body>"
Response.Write msg
%>
