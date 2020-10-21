Attribute VB_Name = "Complex"
Sub DownloadAddressBook()
'This is an Excel Macro
Dim i As Long, j As Long, lastRow As Long
'Set up Outlook
Dim olApp As Outlook.Application
Dim olNS As Outlook.Namespace
Dim olGAL As Outlook.AddressList
Dim olEntry As Outlook.AddressEntries
Dim olMember As Outlook.AddressEntry

Set olApp = Outlook.Application
Set olNS = olApp.GetNamespace("MAPI")
Set olGAL = olNS.GetGlobalAddressList()

'Set Up Excel
Dim wb As Workbook, ws As Worksheet

'set the workbook:
Set wb = ThisWorkbook
'set the worksheet where you want to post Outlook data:
Set ws = wb.Sheets("Sheet1")

'clear all current entries
Cells.Select
Selection.ClearContents

'set and format headings in the worksheet:
ws.Cells(1, 1).Value = "First Name"
ws.Cells(1, 2).Value = "Last Name"
ws.Cells(1, 3).Value = "Email"
ws.Cells(1, 4).Value = "Department"
ws.Cells(1, 5).Value = "Country"
Application.ScreenUpdating = False
With ws.Range("A1:E1")

.Font.Bold = True
.HorizontalAlignment = xlCenter

End With

Set olEntry = olGAL.AddressEntries
On Error Resume Next
'first row of entries
j = 2

' loop through dist list and extract members
For i = 1 To olEntry.Count

Set olMember = olEntry.Item(i)

    If olMember.AddressEntryUserType = olExchangeUserAddressEntry Then
'add to worksheet
    ws.Cells(j, 1).Value = olMember.GetExchangeUser.FirstName
    ws.Cells(j, 2).Value = olMember.GetExchangeUser.LastName
    ws.Cells(j, 3).Value = olMember.GetExchangeUser.PrimarySmtpAddress
    ws.Cells(j, 4).Value = olMember.GetExchangeUser.Department
    ws.Cells(j, 5).Value = olMember.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3A26001E")

    j = j + 1
    End If
Next i
Application.ScreenUpdating = True
'determine last data row, basis column B (contains Last Name):
lastRow = ws.Cells(Rows.Count, "B").End(xlUp).Row

'format worksheet data area:
ws.Range("A2:E" & lastRow).Sort Key1:=ws.Range("B2"), Order1:=xlAscending
ws.Range("A2:E" & lastRow).HorizontalAlignment = xlLeft
ws.Columns("A:E").EntireColumn.AutoFit

wb.Save

'quit the Outlook application:
applOutlook.Quit

'clear the variables:
Set olApp = Nothing
Set olNS = Nothing
Set olGAL = Nothing

End Sub
