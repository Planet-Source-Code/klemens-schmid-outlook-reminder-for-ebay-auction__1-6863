VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDateFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Outlook productivity tools from Klemens Schmid (klemens.schmid@gmx.de)
'For more visit www.schmidks.de


'This class provides a conversion from an arbitrary date format to VB type Date.
'It can be used e.g. for conversion of dates found in Web pages. It can handle
'numeric months as well as month names.

'Sample for using this class:
'Dim oDF as New clsDateFormat
'oDF.MonthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
'MsgBox oDF.ConvertDate("2000-Sep-01", "y-m-d")

'29-Aug-00: Adding function ToLocalTime converting a time value to your local time zone.

Option Explicit

Const Err_UnexpectedChar = 1 + vbObjectError + 512
Const Err_UnexpectedEnd = 2 + vbObjectError + 512

Private Declare Function GetTimeZoneInformation _
   Lib "kernel32" (lpTimeZoneInformation As _
   TIME_ZONE_INFORMATION) As Long

Private Type SYSTEMTIME
   wYear As Integer
   wMonth As Integer
   wDayOfWeek As Integer
   wDay As Integer
   wHour As Integer
   wMinute As Integer
   wSecond As Integer
   wMilliseconds As Integer
End Type

Private Type TIME_ZONE_INFORMATION
   Bias As Long
   StandardName(0 To 63) As Byte
   StandardDate As SYSTEMTIME
   StandardBias As Long
   DaylightName(0 To 63) As Byte
   DaylightDate As SYSTEMTIME
   DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2

Dim mintPosDay  As Integer
Dim mintPosMonth  As Integer
Dim mintPosYear  As Integer
Dim mstrDelim As String
Dim maMonthNames As Variant
Dim mintDay As Integer
Dim mintMonth As Integer
Dim mintYear As Integer
Dim mcntDMY_value As Integer
Dim mcntDMY_format As Integer

Private Property Let DMY_format(s As String)
'called by the parser to assign the position of day, month and year in the format string
Select Case s
   Case "d":   mintPosDay = mcntDMY_format
   Case "m":   mintPosMonth = mcntDMY_format
   Case "y":   mintPosYear = mcntDMY_format
End Select
mcntDMY_format = mcntDMY_format + 1
End Property

Private Property Let Delim(s As String)
'delimiter between day, month and year
If mstrDelim <> "" Then
   If mstrDelim <> s Then
      Err.Raise Err_UnexpectedChar
   End If
Else
   mstrDelim = s
End If
End Property

Private Property Get Delim() As String
Delim = mstrDelim
End Property

Private Sub CheckDelim(ch As String)
'check whether delimiter matches the format
If ch <> mstrDelim Then
   Err.Raise Err_UnexpectedChar
End If
End Sub

Private Property Let DMY_value(s As String)
'check the passed string and assign it to the appropriate member

Select Case mcntDMY_value
   Case mintPosYear
      mintYear = s
   Case mintPosDay
      mintDay = s
   Case mintPosMonth
      If Val(s) = 0 Then
         'seems to be a month name
         'get the month index
         If IsEmpty(maMonthNames) Then
            'need the month names
            Err.Raise 13 'ok, 'Type mismatch' doesn't fit perfectly
         End If
         Dim i%
         For i = 0 To 11
            If LCase(maMonthNames(i)) = LCase(s) Then
               mintMonth = i + 1
               Exit For
            End If
         Next
      Else
         'assign month index
         mintMonth = Val(s)
      End If
End Select
'proceed
mcntDMY_value = mcntDMY_value + 1
End Property

Public Property Let DateFormat(ByVal s As String)
mcntDMY_format = 0 'init

'This code was generated by Klemens' Lex4VB
'Get it from http://www.schmidks.de
Dim Token$
Dim State%, OldState%
Dim Cnt%
Dim ch$
Dim p%

On Error GoTo Trap
p = 1: State = 0: OldState = -1
s = s & Chr(0)
Do While p <= Len(s)
If State = OldState Then Cnt = Cnt + 1 Else Cnt = 0
OldState = State
ch = Mid$(s, p, 1)
Select Case State
Case 0:
If ch Like "[dmy]" Then
DMY_format = ch
State = 1
Else: Err.Raise Err_UnexpectedChar
End If
Case 1:
If ch Like "[-/.]" Then
Delim = ch
State = 2
Else: Err.Raise Err_UnexpectedChar
End If
Case 2:
If ch Like "[dmy]" Then
DMY_format = ch
State = 3
Else: Err.Raise Err_UnexpectedChar
End If
Case 3:
If ch Like "[-/.]" Then
Delim = ch
State = 4
Else: Err.Raise Err_UnexpectedChar
End If
Case 4:
If ch Like "[dmy]" Then
DMY_format = ch
State = 5
Else: Err.Raise Err_UnexpectedChar
End If
Case 5:
If Asc(ch) = 0 Then
State = 5
Else: Err.Raise Err_UnexpectedChar
End If
End Select
p = p + 1
Loop
If State <> 5 Then Err.Raise Err_UnexpectedEnd
Exit Property

Trap:
If Err.Number = Err_UnexpectedEnd Or ch = vbNullChar Then
Err.Description = "Unexpected end of string"
Else
Err.Description = "Unexpected character " & ch & " at position " & p
End If
Err.Raise Err.Number
End Property

Public Property Let MonthNames(a As Variant)
If UBound(a) <> 11 Then
   Err.Raise 9 'subscript out of range
End If
maMonthNames = a
End Property

Public Function ConvertDate(DateString As String, Optional DateFormat As String)
'convert the free-format date/time to VB date
'optionally the time zone is adapted
mcntDMY_value = 0
If Len(DateFormat) > 0 Then
   Me.DateFormat = DateFormat
End If
ParseDate DateString
ConvertDate = DateSerial(mintYear, mintMonth, mintDay)
End Function

Sub ParseDate(ByVal s As String)
'This code was generated by Klemens' Lex4VB
'Get it from http://www.schmidks.de
Dim Token$
Dim State%, OldState%
Dim Cnt%
Dim ch$
Dim p%

On Error GoTo Trap
p = 1: State = 0: OldState = -1
s = s & Chr(0)
Do While p <= Len(s)
If State = OldState Then Cnt = Cnt + 1 Else Cnt = 0
OldState = State
ch = Mid$(s, p, 1)
Select Case State
Case 0:
If ch Like "[0-9a-zA-Z]" Then
If Cnt > 3 Then Err.Raise Err_UnexpectedChar
Token = Token & ch
State = 0
ElseIf ch Like "[-/.]" Then
DMY_value = Token: CheckDelim ch
Token = ""
State = 1
Else: Err.Raise Err_UnexpectedChar
End If
Case 1:
If ch Like "[0-9a-zA-Z]" Then
If Cnt > 3 Then Err.Raise Err_UnexpectedChar
Token = Token & ch
State = 1
ElseIf ch Like "[-/.]" Then
DMY_value = Token: CheckDelim ch
Token = ""
State = 2
Else: Err.Raise Err_UnexpectedChar
End If
Case 2:
If ch Like "[0-9a-zA-Z]" Then
If Cnt > 3 Then Err.Raise Err_UnexpectedChar
Token = Token & ch
State = 2
ElseIf Asc(ch) = 0 Then
DMY_value = Token
State = 2
Else: Err.Raise Err_UnexpectedChar
End If
End Select
p = p + 1
Loop
If State <> 2 Then Err.Raise Err_UnexpectedEnd
Exit Sub

Trap:
If Err.Number = Err_UnexpectedEnd Or ch = vbNullChar Then
Err.Description = "Unexpected end of string"
Else
Err.Description = "Unexpected character " & ch & " at position " & p
End If
Err.Raise Err.Number
End Sub

Public Function ToLocalTime(d As Date, TimeZone As String) As Date
'translate date/time from passed time zone to local time zone
Dim nRet As Long
Dim tz As TIME_ZONE_INFORMATION
Dim BiasLocal As Long
Dim BiasOther As Long

'get the bias of the local time zone
nRet = GetTimeZoneInformation(tz)
If nRet = TIME_ZONE_ID_INVALID Then
   'leave time unchanged
   Debug.Print "Invalid time zone"
   ToLocalTime = d
End If
BiasLocal = (tz.Bias + IIf(nRet = TIME_ZONE_ID_DAYLIGHT, tz.DaylightBias, tz.StandardBias)) / 60
'get the bias of the other time zone
BiasOther = GetTimeZoneBias(TimeZone)
'transform time
ToLocalTime = d - BiasLocal / 24 - BiasOther / 24
End Function

Private Function GetTimeZoneBias(TimeZone As String) As Long
'get the bias in hours for the passed time zone
'source for this data is http://www.iol.ie/~taeger/tables/tab7.htm
Select Case TimeZone
   Case "NZDT", "NZT/S", "NZWT":
      GetTimeZoneBias = 13
   Case "OZDT", "OZT/S", "OZWT":
      GetTimeZoneBias = 12.5
   Case "IDLE", "NZT", "R11D", "R11W", "Z11/S":
      GetTimeZoneBias = 12
   Case "OZT":
      GetTimeZoneBias = 11.5
   Case "AEDT", "AEST/S", "AEWT", "GDT", "GST/S", "GWT", "R11T", "Z11":
      GetTimeZoneBias = -11
   Case "ACDT", "ACST/S", "SADT", "SAT/S", "SAWT":
      GetTimeZoneBias = 10.5
   Case "AEST", "GST", "JDT", "JST/S", "JWT":
      GetTimeZoneBias = 10
   Case "ACST", "SAT":
      GetTimeZoneBias = 9.5
   Case "AWDT", "AWST/S", "CCDT", "CCT/S", "CCWT", "JST":
      GetTimeZoneBias = 9
   Case "JVDT", "JVT/S", "JVWT":
      GetTimeZoneBias = 8.5
   Case "AWST", "CCT", "SSDT", "SST/S", "SSWT":
      GetTimeZoneBias = 8
   Case "JVT", "NSDT", "NST/S", "NSWT":
      GetTimeZoneBias = 7.5
   Case "R5DT", "R5T/S", "R5WT", "SST":
      GetTimeZoneBias = 7
   Case "IDT", "IST/S", "IWT", "NST":
      GetTimeZoneBias = 6.5
   Case "R4DT", "R4T/S", "R4WT", "R5T":
      GetTimeZoneBias = 6
   Case "AGDT", "AGT/S", "IST":
      GetTimeZoneBias = 5.5
   Case "R3DT ", "R3T/S", "R3WT ", "R4T":
      GetTimeZoneBias = 5
   Case "AGT", "IRDT", "IRT/S", "IRWT":
      GetTimeZoneBias = 4.5
   Case "BGDT", "BGT/S", "BGWT", "R3T":
      GetTimeZoneBias = 4
   Case "KEDT", "KET/S", "KEWT":
      GetTimeZoneBias = 3.5
   Case "BGT", "EEDT", "EET/S", "EEWT", "MEDST", "MET/DS", "OESZ", "R2T":
      GetTimeZoneBias = 3
   Case "KET":
      GetTimeZoneBias = 2.5
   Case "EET", "MEDT", "MESZ", "MET/S", "MEWT", "OEZ", "R1T", "WDST", "WEDSZ", "WET/DS":
      GetTimeZoneBias = 2
   Case "CET", "MET", "MEZ", "WEDT", "WESZ", "WET/S", "WEWT":
      GetTimeZoneBias = 1
   Case "AZDT", "AZT/S", "AZWT", "WAT":
      GetTimeZoneBias = -1
   Case "NFDD", "NFT/DS":
      GetTimeZoneBias = -1.5
   Case "AZT", "BZDT", "BZT/S", "BZWT":
      GetTimeZoneBias = -2
   Case "NFDT", "NFT/S", "NFWT":
      GetTimeZoneBias = -2.5
   Case "ADT", "AST/S", "AWT", "BZT":
      GetTimeZoneBias = -3
   Case "NFT":
      GetTimeZoneBias = -3.5
   Case "AST", "EDT", "EST/S", "EWT":
      GetTimeZoneBias = -4
   Case "CDT", "CST/S", "CWT", "EST":
      GetTimeZoneBias = -5
   Case "CST", "MDT", "MST/S", "MWT":
      GetTimeZoneBias = -6
   Case "MST", "PDT", "PST/S", "PWT":
      GetTimeZoneBias = -7
   Case "PST", "YDT", "YST/S", "YWT":
      GetTimeZoneBias = -8
   Case "YST":
      GetTimeZoneBias = -9
   Case "HDT", "HST/S", "HWT":
      GetTimeZoneBias = -9.5
   Case "AHST", "AHT", "NADT", "NAT/S", "NAWT":
      GetTimeZoneBias = -10
   Case "HST":
      GetTimeZoneBias = -10.5
   Case "NAT":
      GetTimeZoneBias = -11
   Case "IDLW":
      GetTimeZoneBias = -12
   Case "WET", "WEZ", "WADT", "WAT/S", "WAWT":
      GetTimeZoneBias = 0
   Case Else
      GetTimeZoneBias = 0
End Select
   
End Function
