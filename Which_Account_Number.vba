Sub Which_Account_Number()
    Dim OutApp As Outlook.Application
    Dim I As Long

    Set OutApp = CreateObject("Outlook.Application")

    For I = 1 To OutApp.Session.Accounts.Count
        MsgBox OutApp.Session.Accounts.Item(I) & " : This is account number ---> " & I
    Next I

    MsgBox ("That is all")

End Sub
