Attribute VB_Name = "modReminder"
'Outlook productivity tools from Klemens Schmid (klemens.schmid@gmx.de)
'For more visit www.schmidks.de

'29-Aug-00: Adapted to Ebay's new page layout
'           Process the time zone found on the page

Option Explicit

Sub RememberWebPage()
'create an appointment for a Web page with a link to the page
'preferred logic for German and US eBay auctions: extract date from page
'New: Adapted to new page l

Dim oSWs As ShellWindows
Dim oIE As SHDocVw.InternetExplorer
Dim oDoc As Object
Dim oDocHTML As HTMLDocument
Dim oAppt As AppointmentItem
Dim oRefs As IHTMLElementCollection
Dim oTDs As IHTMLElementCollection
Dim oTD As IHTMLElement
Dim oDF As New clsDateFormat
Dim strEndsDate As String
Dim strEndsTime As String
Dim strEnds As String
Dim strTimeZone As String
Dim p%, p2%

Set oSWs = New ShellWindows
'loop thru the open browser windows and take the first
For Each oIE In oSWs
   Set oDoc = oIE.document
   If TypeOf oDoc Is HTMLDocument Then
      Set oDocHTML = oDoc
      'create a new appoinment
      Set oAppt = Application.CreateItem(olAppointmentItem)
      oAppt.Subject = oDocHTML.Title
      oAppt.Body = oDocHTML.URL
      'is it an ebay auction?
      Set oRefs = oDocHTML.getElementsByTagName("a")
      If oRefs(1) Like "*ebay.de/*" Then
         'it's a German Ebay auction
         oAppt.Subject = "eBay: " & oDocHTML.Title
         'now extract the date
         'it's the <td> tag following the <td..>Auktionsende</td>
         Set oTDs = oDocHTML.getElementsByTagName("td")
         For Each oTD In oTDs
            If LCase(oTD.innerText) Like "*ende" Or LCase(oTD.innerText) Like "ends*" Then
               strEnds = oDocHTML.all(oTD.sourceIndex + 2).innerText
               'Date maybe in different format than our machine's default. Need convert it
               p = InStr(strEnds, " ")
               strEndsDate = Left$(strEnds, p - 1)
               p2 = InStr(p + 1, strEnds, " ")
               strEndsTime = Mid$(strEnds, p, p2 - p)
               strTimeZone = Mid$(strEnds, p2 + 1)
               'set the appts start date and reminder
               oAppt.Start = oDF.ToLocalTime(oDF.ConvertDate(strEndsDate, "d.m.y") & strEndsTime, strTimeZone)
               oAppt.Duration = 0
               oAppt.ReminderMinutesBeforeStart = 600
               oAppt.Sensitivity = olPrivate
               oAppt.BusyStatus = olFree
               'set a category (sample)
               oAppt.Categories = "Auctions"
               'display the appointment for review
               oAppt.Display True
               'take only the first
               GoTo TheEnd
            End If
         Next
      ElseIf oRefs(1) Like "*ebay.com*" Then
         'it's an US Ebay auction
         oAppt.Subject = "eBay: " & oDocHTML.Title
         'now extract the date
         'it's the <td> tag following the <td..>Auktionsende</td>
         Set oTDs = oDocHTML.getElementsByTagName("td")
         For Each oTD In oTDs
            If oTD.innerText Like "Ends*" Then
               strEnds = oDocHTML.all(oTD.sourceIndex + 2).innerText
               'Date maybe in different format than our machine's default. Need convert it
               oDF.MonthNames = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
               p = InStr(strEnds, " ")
               strEndsDate = Left$(strEnds, p - 1)
               p2 = InStr(p + 1, strEnds, " ")
               strEndsTime = Mid$(strEnds, p, p2 - p)
               strTimeZone = Mid$(strEnds, p2 + 1)
               'set the appts start date and reminder
               oAppt.Start = oDF.ToLocalTime(oDF.ConvertDate(strEndsDate, "m-d-y") & strEndsTime, strTimeZone)
               oAppt.Duration = 0
               oAppt.ReminderMinutesBeforeStart = 600
               oAppt.Sensitivity = olPrivate
               oAppt.BusyStatus = olFree
               'set a category (sample)
               oAppt.Categories = "Auctions"
               'display the appointment for review
               oAppt.Display True
               'take only the first
               GoTo TheEnd
            End If
         Next
            
      End If
   End If
Next

TheEnd:
End Sub
