Set objFso = CreateObject("Scripting.FileSystemObject")

Set Folder = objFSO.GetFolder("C:\Users\ron\Desktop\��վ\����ģ��\img\sets\82series")

i = 1

For Each File In Folder.Files


    sNewFile = File.Name

    sNewFile = Replace(sNewFile, sNewFile, Cstr(i)&".jpg")

    if (sNewFile<>File.Name) then

        File.Move(File.ParentFolder+"\"+sNewFile)

    end if

    i = i + 1

Next