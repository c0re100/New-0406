<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<script>
if(window.name!="ljjhhorse"){ var i=1;  while (i<=50)  {    window.alert("�A�Q�@����r�A�§ڡH�o�̬O���檺�A�h�O�B���h�a�I���I�C�C�I50���I�I");    i=i+1;  }top.location.href="../../exit.asp"} </script>
<head><title>�F�C�����ɰ��{�ǡI</title><link rel=stylesheet href='css.css'></head>
<frameset rows="*,50">
<frame name=compfrm src="compete.asp" noresize  scrolling=no >
<frame name=betfrm src="chipin.asp"  scrolling=no marginheight=0 framespacing=0 marginwidth=0>
</frameset>