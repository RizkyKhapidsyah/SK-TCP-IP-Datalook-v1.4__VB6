Attribute VB_Name = "modStart"
Public fMainForm As frmMain
Public ipsMain As New ipStats
Public iphDNS As New IPHostResolver

Sub Main()
    frmSplash.Show
    frmSplash.Refresh
    'Delay time of 1 second..
    oldTime = Timer
    While (oldTime + 5 > Timer)
    DoEvents
    Wend
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash

    'frmDocument.Show
    fMainForm.Show
End Sub

