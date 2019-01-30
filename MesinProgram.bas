Attribute VB_Name = "MesinProgram"
Option Explicit

Public CN As New ADODB.Connection
Public RS As New ADODB.Recordset
Public Objek As Control
Public Pesan As Integer
Public X As Integer
Public LI As ListItem

Public Sub NyambungUtama()
If CN.State = adStateOpen Then CN.Close
    CN.CursorLocation = adUseClient
    CN.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\ksdb.mdb;Persist Security Info=False"
End Sub


'FUNGSI DEFAULT SIMPAN
Public Sub DefaultFormat()
    Select Case FormPengaturan.cmbDefaultSimpan.ListIndex
        Case Is = 0
            FormUtama.CommonDialog1.FilterIndex = 1
        Case Is = 1
            FormUtama.CommonDialog1.FilterIndex = 2
        Case Is = 2
            FormUtama.CommonDialog1.FilterIndex = 3
        Case Is = 3
            FormUtama.CommonDialog1.FilterIndex = 4
        Case Is = 4
            FormUtama.CommonDialog1.FilterIndex = 5
    End Select
End Sub

'BAGIAN UNTUK MEMANTAU ERROR PADA PROGRAM
Public Sub PusatError()
    If Err.Number = 53 Then
        MsgBox "Maaf, gambar source tidak ditemukan!" & vbCrLf & _
                "Silahkan re-install program ini.", vbCritical + vbOKOnly, "Main System - Error"
    ElseIf Err.Number = 380 Then
       On Error Resume Next
    Else
        MsgBox "Maaf, ada kesalahan program, silahkan re-install kembali aplikasi ini!" & vbCrLf & _
                "Error Kode : " & Err.Number & vbCrLf & _
                "Deskripsi : " & Err.Description, vbCritical + vbOKOnly, "Main System - Error"
    End If
End Sub

'BAGIAN UNTUK MEMBULATKAN PECAHAN
Public Function Round(nValue As Double, nDigits As Integer) As Double
    Round = Int(nValue * (10 ^ nDigits) + 0.5) / (10 ^ nDigits)
End Function

