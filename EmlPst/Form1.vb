Imports Redemption.rdoSaveAsType
Public Class Form1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Sub ImportEML(ByVal EmlPath As String)
        Try
            Dim objPost As Microsoft.Office.Interop.Outlook.PostItem
            Dim objSafePost As Global.Redemption.SafePostItem
            Dim objNS As Microsoft.Office.Interop.Outlook.NameSpace
            Dim objInbox As Microsoft.Office.Interop.Outlook.MAPIFolder
            Const PR_ICON_INDEX = &H10800003
            '' On Error Resume Next

            Dim objApp As New Microsoft.Office.Interop.Outlook.Application

            objNS = objApp.GetNamespace("MAPI")
            objInbox = objNS.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderInbox)
            objPost = objInbox.Items.Add(Microsoft.Office.Interop.Outlook.OlItemType.olPostItem)
            ''''objPost = objInbox.Items.Add(olPostItem)

            objSafePost = CreateObject("Redemption.SafePostItem")
            objPost.Save()
            objSafePost.Item = objPost
            objSafePost.Import(EmlPath, olRFC822) ''''''objSafePost.Import("c:\emltest.eml", olRFC822)
            objSafePost.MessageClass = "IPM.Note"
            ' remove IPM.Post icon
            objSafePost.Fields(PR_ICON_INDEX) = Nothing
            objSafePost.Save()

            objSafePost = Nothing
            objPost = Nothing
            objInbox = Nothing
            objNS = Nothing

            If chkRename.Checked Then Rename(EmlPath, EmlPath & ".imported")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        Try
            FB.ShowDialog()
            lblFolder.Text = FB.SelectedPath
            Dim D As System.IO.DirectoryInfo = New System.IO.DirectoryInfo(FB.SelectedPath)
            Dim F() As System.IO.FileInfo = D.GetFiles("*.eml", IO.SearchOption.AllDirectories)

            If UBound(F) > -1 Then
                For i As Integer = 0 To UBound(F)
                    lstFiles.Items.Add(F(i).FullName)
                    Application.DoEvents()
                Next
            End If
            lblProcess.Text = UBound(F).ToString & " Files Found"
            MessageBox.Show("All files loaded.")
        Catch ex As Exception

        End Try
    End Sub

    Private Sub cmdStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdStart.Click
        For i As Integer = 0 To lstFiles.Items.Count - 1
            ImportEML(lstFiles.Items(i).ToString)
            lblProcess.Text = i + 1 & " / " & lstFiles.Items.Count
            PB1.Value = (i * 100) / lstFiles.Items.Count
            lblRec.Text = lstFiles.Items(i).ToString
            Application.DoEvents()
        Next
    End Sub
End Class
