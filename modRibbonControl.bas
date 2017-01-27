Attribute VB_Name = "modRibbonControl"
' Handle calling the invoice import routine with the appropriate
' buisness service line parameter.
' The service line will be passed off by the UI control tag attribute.
'
' target expects service line param: BusinessLine
' exmaple: service1, service2, service3, service4
Sub ButtonOnAction(control As IRibbonControl)

    Select Case control.Tag
        Case "service1"
            refresh_report.GetInvoices control.Tag
        Case "service2"
            refresh_report.GetInvoices control.Tag
        Case "service3"
            refresh_report.GetInvoices control.Tag
        Case Else
            refresh_report.GetInvoices control.Tag
    End Select

lbl_exit:
    Exit Sub
End Sub
