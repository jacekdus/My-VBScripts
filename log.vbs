Option Explicit
Dim user,onet,ofs,DriveCollection,MsgString,i
MsgString = ""
set onet = createobject("WScript.Network")
set ofs = createobject("scripting.filesystemobject")
user = onet.UserName
If ofs.DriveExists("P:") = True Then
   onet.RemoveNetworkDrive "P:", True, True
End If
On Error Resume Next 
onet.mapnetworkdrive "P:", "\\192.168.1.12\Projekty"
If ofs.DriveExists("Q:") = True Then
   onet.RemoveNetworkDrive "Q:", True, True
End If
onet.mapnetworkdrive "Q:", "\\192.168.1.12\Projekty Zamkniete"
If ofs.DriveExists("H:") = True Then
   onet.RemoveNetworkDrive "H:", True, True
End If
onet.mapnetworkdrive "H:", "\\192.168.1.12\" & user
If ofs.DriveExists("R:") = True Then
   onet.RemoveNetworkDrive "R:", True, True
End If
onet.mapnetworkdrive "R:", "\\192.168.1.14\Dz"
If ofs.DriveExists("O:") = True Then
   onet.RemoveNetworkDrive "O:", True, True
End If
onet.mapnetworkdrive "O:", "\\192.168.1.14\ORGANIZACJA"
If ofs.DriveExists("S:") = True Then
   onet.RemoveNetworkDrive "S:", True, True
End If
onet.mapnetworkdrive "S:", "\\192.168.1.14\Rozliczenie Godzin"
IF user = "rybak_e" or user = "biuro_k"  or user = "zamors_d" or user = "milkow_z" or user = "pajonk_p" then
    If ofs.DriveExists("J:") = True Then
       onet.RemoveNetworkDrive "J:", True, True
    End If
    onet.mapnetworkdrive "J:", "\\192.168.1.14\BIURO"
    If ofs.DriveExists("K:") = True Then
       onet.RemoveNetworkDrive "K:", True, True
    End If
    onet.mapnetworkdrive "K:", "\\192.168.1.14\korespap"
END IF
IF user = "zak_j"  or user = "olszow_m" then
    If ofs.DriveExists("J:") = True Then
       onet.RemoveNetworkDrive "J:", True, True
    End If
    onet.mapnetworkdrive "J:", "\\192.168.1.14\BIURO"
    If ofs.DriveExists("K:") = True Then
       onet.RemoveNetworkDrive "K:", True, True
    End If
    onet.mapnetworkdrive "K:", "\\192.168.1.14\korespap"
END IF
IF user = "spadzi_j" then
    If ofs.DriveExists("K:") = True Then
       onet.RemoveNetworkDrive "K:", True, True
    End If
    onet.mapnetworkdrive "K:", "\\192.168.1.14\korespap"
END IF
IF user = "chowan_g" or user = "milkow_a" or user = "milkow_s" or user = "sojka_r" or user = "wardzi_d" then
    If ofs.DriveExists("M:") = True Then
       onet.RemoveNetworkDrive "M:", True, True
    End If
    onet.mapnetworkdrive "M:", "\\192.168.1.14\Dyrektorzy"
    If ofs.DriveExists("J:") = True Then
       onet.RemoveNetworkDrive "J:", True, True
    End If
    onet.mapnetworkdrive "J:", "\\192.168.1.14\BIURO"
    If ofs.DriveExists("K:") = True Then
       onet.RemoveNetworkDrive "K:", True, True
    End If
    onet.mapnetworkdrive "K:", "\\192.168.1.14\korespap"
END IF
IF user = "wardzi_d" or user = "milkow_s" or user = "milkow_z" or user = "biuro_k" then
    If ofs.DriveExists("I:") = True Then
       onet.RemoveNetworkDrive "I:", True, True
    End If
    onet.mapnetworkdrive "I:", "\\192.168.1.14\CDN"
END IF
IF user = "Admin" then
    If ofs.DriveExists("M:") = True Then
       onet.RemoveNetworkDrive "M:", True, True
    End If
    onet.mapnetworkdrive "M:", "\\192.168.1.14\Dyrektorzy"
    If ofs.DriveExists("J:") = True Then
       onet.RemoveNetworkDrive "J:", True, True
    End If
    onet.mapnetworkdrive "J:", "\\192.168.1.14\BIURO"
    If ofs.DriveExists("K:") = True Then
       onet.RemoveNetworkDrive "K:", True, True
    End If
    onet.mapnetworkdrive "K:", "\\192.168.1.14\korespap"
END IF
Set DriveCollection = onet.EnumNetworkDrives
MsgString = user & " masz podłączone takie dyski sieciowe : " & vbCRLF
For i = 0 To DriveCollection.Count - 1 Step 2
    MsgString = MsgString & vbCRLF & DriveCollection(i) & Chr(9) & DriveCollection(i + 1)
Next    
MsgBox MsgString,vbInformation,"Logowanie"
Set onet = Nothing
Set ofs = Nothing
Set DriveCollection = Nothing
WScript.Quit
