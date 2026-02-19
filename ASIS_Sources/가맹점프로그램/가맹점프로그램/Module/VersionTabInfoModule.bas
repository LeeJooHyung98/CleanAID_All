Attribute VB_Name = "VersionTabInfoModule"
Option Explicit

Public Type VS_FIXEDFILEINFO
            dwSignature            As Long
            dwStrucVersion         As Long
            dwFileVersionMS        As Long
            dwFileVersionLS        As Long
            dwProductVersionMS     As Long
            dwProductVersionLS     As Long
            dwFileFlagsMask        As Long
            dwFileFlags            As Long
            dwFileOS               As Long
            dwFileType             As Long
            dwFileSubtype          As Long
            dwFileDateMS           As Long
            dwFileDateLS           As Long
End Type

Public Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias _
                       "GetFileVersionInfoSizeA" _
                       (ByVal lptstrFilename As String, _
                        lpdwHandle As Long) As Long
Public Declare Function GetFileVersionInfo Lib "Version.dll" Alias _
                       "GetFileVersionInfoA" _
                       (ByVal lptstrFilename As String, _
                        ByVal dwHandle As Long, _
                        ByVal dwLen As Long, _
                        lpData As Any) As Long
Public Declare Function VerQueryValue Lib "Version.dll" Alias _
                       "VerQueryValueA" _
                       (pBlock As Any, _
                        ByVal lpSubBlock As String, _
                        lplpBuffer As Any, _
                        nVerSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias _
                       "RtlMoveMemory" _
                       (Destination As Any, _
                        Source As Any, _
                        ByVal Length As Long)

