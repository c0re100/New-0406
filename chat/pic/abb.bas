 OPEN "dir.txt" FOR INPUT AS #1
 abc = 0
 FOR t = 1 TO 261
 LINE INPUT #1, a$
 a$ = RTRIM$(LTRIM$(a$))
 b$ = RTRIM$(LTRIM$(STR$(t)))
 aa$ = "ren " + a$ + " " + b$ + ".gif"
 SHELL aa$
 NEXT t
 CLOSE

