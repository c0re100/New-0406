<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Select Case request.querystring("fun")
Case "�ڪ���O"
jumpage="diary/disp_diary.asp?name="&request.querystring("hostname")
case "�ڪ���ñ"
jumpage="bookmark/lookusermark.asp?user="&request.querystring("hostname")
ENd Select
Response.Redirect( jumpage)
%>