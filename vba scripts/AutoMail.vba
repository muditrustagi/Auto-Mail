'-----FORMAT--STRING--TEMPLATE------------'
'This function formats the string, removing placeholders and inserting values'
Function Format(str As String) As String

'col: the column to begin the placehlder thing from
'rol: the row where placeholder thing are present
Col = 4
rol = 1


Do While True
'traverses the row, until it finds a empty cell
  If IsEmpty(Cells(rol, Col)) Then
    Exit Do
  Else
    If IsEmpty(Cells(ActiveCell.Row, Col)) Then
        'calls the first row, when the current cell is empty
      str = Replace(str, Cells(rol, Col).Text, Cells(rol + 1, Col).Text)

    Else
        'calls the current cell, since it is not empty
      str = Replace(str, Cells(rol, Col).Text, Cells(ActiveCell.Row, Col).Text)

    End If
  End If
  Col = Col + 1
Loop

'returns the new fomatted string
Format = str
End Function



'-----CHECK--PLACE--HOLDER------------'
'This function checks if the string ahs been devoid of placeholders completely and thus returns true or flase based on it'
Function checkPlaceHolder(str As String) As Boolean

'col: the column to begin the placehlder thing from
'rol: the row where placeholder thing are present
Col = 4
rol = 1


Do While True
'traverses the row, until it finds a empty cell
  If IsEmpty(Cells(rol, Col)) Then
    Exit Do
  Else
    'returns false if the string is not valid in any form
      If InStr(1, str, Cells(rol, Col).Text, vbTextCompare) Then
          checkPlaceHolder = False
          Exit Function
      Else
          'do nothing'
      End If

  End If
  Col = Col + 1
Loop

'returns true if the string is valid now
checkPlaceHolder = True
End Function




'-----BUILD--THE--TABLE------------'
Function Files(storeCode As String, templateNumber As Long, ByRef flag As Integer, Col As String) As String
    Dim src As Workbook
    Dim f As String
    Dim leng As Long
    Dim I As Integer
    Dim s As String



    s = Sheets("Templates").Cells(Sheets("Templates").Range("file").Row, templateNumber)
    toSum = Sheets("Templates").Cells(Sheets("Templates").Range("toSum").Row, templateNumber)
    Set src = Workbooks.Open(s, True, True)

    src.Worksheets("Sheet1").Activate
    Range("A1").Select


    Row = 1

    leng = Cells(1, Columns.Count).End(xlToLeft).Column

    Dim arr(1 To 100) As Long
    flag = -1
    Do While True

        Range(Col & Row).Select
        s = ActiveCell.Text

        If IsEmpty(ActiveCell) Then
            Exit Do
        Else
            Range("A" & Row).Select

            If Row = 1 Then

                For I = 1 To leng
                    f = f & "<th>" & ActiveCell.Text & "</th>"
                    ActiveCell.Offset(0, 1).Select

                Next
                f = f & "</tr>"
            Else
            End If
            If s = storeCode Then
                flag = 0

                f = f & "<tr>"
                For I = 1 To leng



                    If InStr(1, toSum, Chr(I + 64), vbTextCompare) Then

                    f = f & "<td class=" & Chr(34) & "righty" & Chr(34) & ">" & ActiveCell.Text & "</td>"
                        arr(I) = arr(I) + CDbl(ActiveCell.Value)
                    Else

                        f = f & "<td>" & ActiveCell.Text & "</td>"

                        arr(I) = -1
                    End If

                    ActiveCell.Offset(0, 1).Select

                Next
                f = f & "</tr>"
            Else
            End If
        End If
        Row = Row + 1
    Loop

    f = f & "<tr style=" & Chr(34) & "background-color:#dddddd;" & Chr(34) & " >"

    For I = 1 To leng
        f = f & "<th>"
        If arr(I) = -1 Then
            If I = 1 Then
                f = f & "TOTAL"
            Else
                'do nothing
            End If
        Else
            Dim str As String
            str = FormatNumber(arr(I), 2, , , vbTrue)
            str = Left(str, Len(str) - 3)
            f = f & str
        End If

        f = f & "</th>"
    Next

    f = f & "</tr>"
    ActiveWorkbook.Close

    Files = f

End Function


Sub AUTOMAIL()
Application.Wait (Now + TimeValue("0:00:01"))
If MsgBox("Are you sure you want to AutoMail, all recipients?" & vbNewLine & "Press Yes to Confirm" & vbNewLine & "Press No to Quit", vbYesNo, "AutoMail") = vbNo Then
    MsgBox ("No Mails Sent")
    Exit Sub
Else
  'do nothing'
End If

tempory = MsgBox("REMINDER:" & vbNewLine & "YOUR OUTLOOK SHOULD BE OPEN" & vbNewLine & " IF NOT, OPEN IT NOW", vbCritical)


Dim outlookApp As Outlook.Application
Dim myMail As Outlook.MailItem
Set outlookApp = New Outlook.Application
Set myMail = outlookApp.CreateItem(olMailItem)

Dim s As String
Dim name As String
Dim temporary As String
Dim Row As Long
Dim templateOption As String
Dim templateNumber As Long
Dim storeCode As String
Dim objFSO As FileSystemObject
Dim objFolder As Folder
Dim objFile As File
Dim strPath As String
Dim strFile As String
Dim NextRow As Long
Dim m As String
Dim flag As Integer
Dim Col As String

Application.ScreenUpdating = False


'Row defines the row number to start from in AutoMail'
Row = 2


Application.Wait (Now + TimeValue("0:00:01"))
'TemplateNumber defines the template to be used'
templateOption = InputBox("Enter Template Column ( 1 - 100 )" & vbNewLine & "Default Template : 01", "AutoMail Template", 1)


If StrPtr(templateOption) = 0 Then
    MsgBox ("No Mails Sent")
    Exit Sub
Else
    templateNumber = CInt(templateOption) + 1
End If

Application.Wait (Now + TimeValue("0:00:01"))


Col = Sheets("Templates").Cells(Sheets("Templates").Range("key").Row, templateNumber)


Do While True

    Set outlookApp = New Outlook.Application
    Set myMail = outlookApp.CreateItem(olMailItem)

    flag = 0
    Range("B" & Row).Activate
    If IsEmpty(ActiveCell) Then
        Exit Do
    Else

    storeCode = Cells(Row, Range("storeCode").Column)

    myMail.To = Cells(Row, Range("to").Column)

    myMail.CC = Format(Sheets("Templates").Cells(Sheets("Templates").Range("cc").Row, templateNumber))

    myMail.BCC = Sheets("Templates").Cells(Sheets("Templates").Range("bcc").Row, templateNumber)

    m = Format(Sheets("Templates").Cells(Sheets("Templates").Range("subject").Row, templateNumber))


    If checkPlaceHolder(m) Then
        myMail.Subject = m
    Else
        flag = -1
    End If



    m = Format(Sheets("Templates").Cells(Sheets("Templates").Range("style").Row, templateNumber))
    m = m & Format(Sheets("Templates").Cells(Sheets("Templates").Range("head").Row, templateNumber))

    If (IsEmpty(Sheets("Templates").Cells(Sheets("Templates").Range("table").Row, templateNumber))) Then
        'do nothing'
    Else
        m = m & Format(Sheets("Templates").Cells(Sheets("Templates").Range("table").Row, templateNumber))

        m = m & Files(storeCode, templateNumber, flag, Col)
        m = m & Format(Sheets("Templates").Cells(Sheets("Templates").Range("tableEnd").Row, templateNumber))
    End If



    m = m & Format(Sheets("Templates").Cells(Sheets("Templates").Range("foot").Row, templateNumber))



    If (checkPlaceHolder(m)) Then
        myMail.HTMLBody = m
    Else
        flag = -1
    End If




    'Dim obj As New DataObject
    'obj.SetText m
    'obj.PutInClipboard
    'MsgBox ("ClipBoard Updated")



    Dim strFileName As String
    Dim strFolder As String
    Dim strFileSpec As String

    strFolder = Sheets("Templates").Cells(Sheets("Templates").Range("folder").Row, templateNumber)
    strFileSpec = strFolder & "*.*"
    strFilePath = strFolder & strFileName
    strFileName = Dir(strFileSpec)


    Do While Len(strFileName) > 0
        strFilePath = strFolder & strFileName
        If InStr(1, strFileName, storeCode, vbTextCompare) > 0 Then
            myMail.Attachments.Add strFilePath
        Else:
            'do nothing'
        End If
        strFileName = Dir
    Loop

Application.Wait (Now + TimeValue("0:00:01"))



    If (flag = 0) Then
    

        Application.Wait (Now + TimeValue("0:00:02"))

        myMail.SendUsingAccount = outlookApp.Session.Accounts.Item(Sheets("Templates").Cells(Sheets("Templates").Range("from").Row, templateNumber))
        myMail.Send
        'myMail.Display


        Application.Wait (Now + TimeValue("0:00:02"))

        Range("A" & (Row)).Value = "SENT"
    Else
        Range("A" & (Row)).Value = "NOT SENT"
    End If

    Row = Row + 1
  End If

Loop

Application.ScreenUpdating = True
MsgBox ("AutoMail Complete")
End Sub
