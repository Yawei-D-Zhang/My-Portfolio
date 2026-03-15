'send quote to client via outlook
Private Sub BtnEmail_Click()

    Dim O As Outlook.Application
    Dim M As Outlook.MailItem
    
    Set O = New Outlook.Application
    Set M = O.CreateItem(olMailItem)
    
    With M
        .BodyFormat = olFormatPlain
        .Body = "Hi, " & _
        "please review the pricing for " & Client & ". Quote No. is  " & QuotationNumber & "."
        .To = Email
        .Subject = "BMS Project Works Pricing - Action required"
        .Display
    End With
    
    Set M = Nothing
    Set O = Nothing


End Sub

'Refresh button in Project Info
Private Sub BtnRefresh_Click()

Me.Totalpoints = Me.Admintotalpoints
Me.Totaladmincost = Me.AdmincostCal
Me.ProjectCost = Me.CalCost
If IsNull(NewCalSell) Then
Me.ProjectSelling = Me.CalSelling
Else
Me.ProjectSelling = Me.NewCalSell
End If
Me.Refresh
End Sub

'Generate quote
Private Sub GetQuote_Click()
DoCmd.OpenForm FormName:="FRMQuotation"
Forms!FRMQuotation!ProjectID = Me.ProjeciID
Forms!FRMQuotation!ProjectName = Me.ProjectName
Forms!FRMQuotation!ProjectType = Me.ProjectType
Forms!FRMQuotation!Client = Me.Client
Forms!FRMQuotation!PreparedBy = Me.PreparedBy
Forms!FRMQuotation!QuotationNumber = Me.QuotationNumber
Forms!FRMQuotation!JobDescription = Me.JobDescription
Forms!FRMQuotation!SystemType = Me.SystemType
Forms!FRMQuotation!ProjectSelling = Me.ProjectSelling

End Sub

