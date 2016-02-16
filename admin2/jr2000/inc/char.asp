<%

'髒字過濾
Function CkBadWords(fString)
    IF not(isnull(BadWords) or isnull(fString)) then
    BadWords = split(BadWords, "|")
	dim i
    For i = 0 to ubound(BadWords)
        fString = Replace(fString, BadWords(i), string(len(BadWords(i)),"*")) 
    Next
    CkBadWords = fString
    End IF
End Function
'檢查年齡
Function CkAge(LngAge)
IF LngAge="" Then Error("您忘了填寫年齡")
Dim RE
Set RE = New RegExp
RE.Pattern = "[0-9]{"&Len(LngAge)&"}"
RE.IgnoreCase = False
RE.Global = False
CkAge = RE.Test(LngAge)
Set RE = Nothing
IF CkAge = True Then
IF LngAge < 2 Then Error("您的年齡太小了<br>PS：孩子 您這麼小就能上網啊 真厲害! 將來肯定比我強! 來哥哥(SLIGHTBOY)賞你個棒棒糖 ^*^")
IF LngAge > 100 Then Error("您的年齡太大了<br>PS：不會吧？ 老爺爺 精力充沛啊 還上網啊 活這麼久能告訴我方法嗎？ 我幫你宣傳宣傳 ^*^")
Else
Error("年齡欄隻能輸入數字")
End IF
End Function
'郵件檢查
Function IsValidEmail(email)

dim names, name, i, c

'Check for valid syntax in an email address.

IsValidEmail = true
names = Split(email, "@")
if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
end if
for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
next
if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
end if
i = Len(names(1)) - InStrRev(names(1), ".")
if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
end if
if InStr(email, "..") > 0 then
   IsValidEmail = false
end if
End Function

'HTML 編碼 (不支持)
Function HTMLEncode(fString)
IF not isnull(fString) then
    fString = Replace(fString, ">", "&gt;")
    fString = Replace(fString, "<", "&lt;")
    fString = Replace(fString, CHR(32), "&nbsp;")
    fString = Replace(fString, CHR(34), "&quot;")
    fString = Replace(fString, CHR(39), "&#39;")
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P> ")
    fString = Replace(fString, CHR(10), "<BR> ")
    HTMLEncode = fString
End IF
End Function
'HTML 支持
Function HTMLcode(fString)
IF not isnull(fString) then
    fString = Replace(fString, CHR(13), "")
    fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
    fString = Replace(fString, CHR(10), "<BR>")
    HTMLcode = fString
End IF
End Function
%>