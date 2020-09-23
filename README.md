<div align="center">

## International Date  Issues


</div>

### Description



If you wander away from VB's built-in date handling (and even if you stick fairly close to it) eventually you will have trouble if users can input a date. While MS and the WWW have made most computer literate people used to the the USA format M/D/Y there are many others. MS recognises 3 MDY, DMY and YMD. There are also lots of date dividers '/-\,.'. If you don't provide a calander or listbox system to control date input then you have to program for end-users using their local format. The following article is a at least part of a quick and dirty answer to the problem. (article is also in zip for easier downloading in English and Indonesian)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2003-08-19 01:00:04
**By**             |[Roger Gilchrist](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roger-gilchrist.md)
**Level**          |Intermediate
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[Internatio1631978192003\.zip](https://github.com/Planet-Source-Code/roger-gilchrist-international-date-issues__1-47822/archive/master.zip)





### Source Code

```

DATE INTERNATIONALI(S|Z)ATION
<P>
This was inspired by a recent upload which ran into a problem with the way VB struggles to work out if dates are legal.
The upload was from Turkey and looking into the code Turkish uses a '.' separator for dates. VB is fairly tolerant of many date dividers '/\-' but '.' confuses it.
As a result the upload mentioned above fails on most systems. But the following routines should get you through.
Remember like most internationalisation (or internationalization if you're American) problems it is a pain to change your system just to test it so you have to take some of it on trust.
The following is based on code in Michael S. Kaplan's 'Internationalization with Visual Basic' (c)2000 Sams Publishing.
I have simplified it a bit, see the book if your really interested, it is very detailed and very good.
<pre>
Private Const LOCALE_SDATE As Long = &H1D
Private Const LOCALE_ILDATE As Long = &H22
'You can find many others
'in VB help under 'Locale Information ' No values but lots of explanations
'or in API viewer search for 'LOCAL_' ' No explanations but has values
Private Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" ( _
 ByVal Locale As Long, _
 ByVal LCType As Long, _
 ByVal lpLCData As String, _
 ByVal cchData As Long) As Long
Public Function LocalizationData(ByVal LData As Long) As String
'This is a general routine to read whatever bit of data
'you want based on the constants fed to it as LData
 Dim stBuff As String * 255
 Dim Ret  As Long
 Ret = GetLocaleInfo(1024, LData, ByVal stBuff, Len(stBuff))
 If Ret Then
 'for systems using UniCode (Win2K+)
 LocalizationData = Left$(stBuff, Ret - 1)
 'For Ascii systems (Pre Win2K)
 'LocalizationData = Left$(stBuff, Ret)
 'If you are not sure set a watch point and check whether
 'there is a Null character on end of return or not.
 'You want the return without the null
 'You could also use a Function which strips nulls
 'LocalizationData = StripNulls(Left$(stBuff, Ret))
 End If
End Function
Public Function LocalDateDiv() As String
' gets the date divisor
 LocalDateDiv = LocalizationData(LOCALE_SDATE)
End Function
Public Function LocalDMY() As Integer
'gets the D M Y order
'Returns 0,1, or 2
'0 Month -Day - Year
'1 Day -Month - Year
'2 Year -Month - Day
LocalDMY = LocalizationData(LOCALE_ILDATE)
End Function
Function StripNulls(strTest as string) as string
StripNulls = Replace(strTest, vbNullString, "")
End Function
</pre>
and use like this
<pre>
Public Function RealDate(ByVal D As Integer, _
          ByVal M As Integer, _
          ByVal Y As Long) As Boolean
 Select Case LocalDMY
 Case 0
 RealDate=IsDate(Format$(M, "00") & LocalDateDiv & Format$(D, "00") & LocalDateDiv & Y)
 Case 1
 RealDate=IsDate(Format$(D, "00") & LocalDateDiv & Format$(M, "00") & LocalDateDiv & Y)
 Case 2
 RealDate=IsDate(Y & LocalDateDiv & Format$(M, "00") & LocalDateDiv & Format$(D, "00")
 End Select
End Function
</pre>
<P>
(c) 2003 Roger Gilchrist
<P>
rojagilkrist@hotmail.com
```

