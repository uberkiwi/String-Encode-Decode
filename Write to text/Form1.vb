Imports System
Imports System.IO
Imports System.Text

Public Class Form1
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles Entry.TextChanged

    End Sub

    Private Sub Encode_Click(sender As Object, e As EventArgs) Handles Encode.Click
        Dim sText As String
        Dim i As Integer
        Dim lenText As Integer
        Dim Msg As Integer
        Dim lenKey As Integer
        Dim KeyAsc As String
        Dim KeyCode As String
        Dim MultList As New List(Of Integer)
        Dim numKey As Integer
        Dim Coded As String

        'Check if message and key box is empty
        If (Entry.Text = "" Or Entry.Text = "Please enter something." Or Entry.Text = "Enter the message...") Then
            Entry.Text = "Please enter something."
            Return
        End If
        If (Key.Text = "" Or Key.Text = "Please enter something." Or Key.Text = "Key") Then
            Entry.Text = "Please enter something."
            Return
        End If

        'Gets message and key, and their lengths
        sText = Entry.Text
        lenText = Len(sText)
        KeyCode = Key.Text
        lenKey = Len(KeyCode)

        'Converts the characters to ASCII and adds each ASCII value to MultList
        For i = 1 To lenKey
            KeyAsc = KeyAsc & CStr(Asc(Mid$(KeyCode, i, 1)))
            MultList.Add(KeyAsc)
            KeyAsc = ""
        Next i

        'Length of MultList
        Dim LenMult As Integer
        For Each int As Integer In MultList
            LenMult = LenMult + 1
        Next

        'Converts Each letter in the message to ASCII and encodes it.'
        For i = 1 To lenText
            Msg = Asc(Mid$(sText, i, 1))
            Dim Index As Integer = ((i - 1) Mod LenMult)
            Dim Multiplier As Integer = MultList(Index)
            Msg = Msg * (i + 1) * Multiplier
            numKey = Math.Floor(Msg / 222)
            Coded = Coded & numKey & Chr((Msg Mod 222) + 33) & " "
        Next i

        'Writes encoded message to text file on desktop
        Dim fw As StreamWriter
        Dim ReadString As String
        Try
            Dim path As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
            fw = New StreamWriter(path & "\secretMessage.txt", True)
            ReadString = Coded
            fw.WriteLine(ReadString)
        Finally
            fw.Close()
        End Try
    End Sub

    Private Sub Decode_Click(sender As Object, e As EventArgs) Handles Decode.Click
        Dim sText As String
        Dim lenText As Integer
        Dim Msg As Integer
        Dim lenKey As Integer
        Dim KeyAsc As String
        Dim KeyCode As String
        Dim Mult As String
        Dim MultList As New List(Of Integer)
        Dim Asci As Integer
        Dim Coded As String
        Dim lenCode As Integer
        Dim AscList As New List(Of Integer)
        Dim LenAsc As Integer

        'Check if message and key box is empty
        If (Entry.Text = "" Or Entry.Text = "Please enter something." Or Entry.Text = "Enter the message...") Then
            Entry.Text = "Please enter something."
            Return
        End If
        If (Key.Text = "" Or Key.Text = "Please enter something." Or Key.Text = "Key") Then
            Entry.Text = "Please enter something."
            Return
        End If

        'Gets Message and Key and their lengths
        sText = Entry.Text
        lenText = Len(sText)
        KeyCode = Key.Text
        lenKey = Len(KeyCode)

        'Converts the characters to ASCII and adds each ASCII value to MultList
        For i = 1 To lenKey
            KeyAsc = KeyAsc & CStr(Asc(Mid$(KeyCode, i, 1)))
            MultList.Add(KeyAsc)
            KeyAsc = ""
        Next i

        'Length of MultList
        Dim LenMult As Integer
        For Each int As Integer In MultList
            LenMult = LenMult + 1
        Next

        'ENCODES TEXT
        'Checks for space in text
        lenCode = Len(sText)
        If GetChar(sText, lenCode) <> " " Then
            sText = sText + " "
            lenCode = Len(sText)
        End If

        Dim Space3 As Integer
        Dim Space4 As Integer
        Space3 = InStr(1, sText, " ")
        Space4 = 1
        lenCode = Len(sText)
        Mult = ""

        While lenCode <> Space3
            For i = Space4 To (Space3 - 2)
                Mult = Mult + GetChar(sText, i)
            Next i
            Asci = Convert.ToInt16(Mult)
            Asci = Asci * 222
            Mult = GetChar(sText, Space3 - 1)
            Asci = Asci + Asc(Mult) - 33
            AscList.Add(Asci)
            Mult = ""
            Space4 = Space3 + 1
            Space3 = InStr(Space4, sText, " ")
        End While
        Space3 = lenCode
        For i = Space4 To (Space3 - 2)
            Mult = Mult + GetChar(sText, i)
        Next i
        Asci = Convert.ToInt16(Mult)
        Asci = Asci * 222
        Mult = GetChar(sText, Space3 - 1)
        Asci = Asci + Asc(Mult) - 33
        AscList.Add(Asci)
        LenMult = 0
        LenAsc = 0
        For Each int As Integer In MultList
            LenMult = LenMult + 1
        Next
        For Each int As Integer In AscList
            LenAsc = LenAsc + 1
        Next


        For i = 1 To LenAsc
            Dim Index As Integer = ((i - 1) Mod LenMult)
            Dim Multiplier As Integer = MultList(Index)

            Msg = ((AscList(i - 1)) / Multiplier) / (i + 1)
            Coded = Coded & Chr(Msg)
        Next i

        MsgBox("Secret Code is: " & Coded)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class

