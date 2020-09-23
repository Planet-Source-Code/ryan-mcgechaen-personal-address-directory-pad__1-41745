Attribute VB_Name = "modpic"
'Used to display pictures in the frmAdd and frmEdit forms
'I'm 100% sure there's an easier way, but If history has taught anyone anything, it's that Ryan McGechaen is lazy


Public Sub picupdateADD()

Dim sExtension As String
 sExtension = UCase(Right$(frmAdd.txtPhoto, 3))
 
 If Dir$(frmAdd.txtPhoto) = "" Then
    MsgBox "Invaild Path", vbExclamation, "Picture Link Field "
    Exit Sub
 End If
 
 Select Case sExtension
    Case "JPG", "GIF", "BMP"
        frmAdd.picPhoto = LoadPicture(frmAdd.txtPhoto.Text)
    Case Else
        MsgBox "Invaild Path", vbExclamation, "Link Field"
End Select

End Sub



Public Sub picupdateEDIT()

Dim sExtension As String
 sExtension = UCase(Right$(frmEdit.txtPhoto, 3))
 
 If Dir$(frmEdit.txtPhoto) = "" Then
    MsgBox "Invaild Path", vbExclamation, "Picture Link Field "
    Exit Sub
 End If
 
 Select Case sExtension
    Case "JPG", "GIF", "BMP"
        frmEdit.picPhoto = LoadPicture(frmEdit.txtPhoto.Text)
    Case Else
        MsgBox "Invaild Path", vbExclamation, "Link Field"
End Select


End Sub

