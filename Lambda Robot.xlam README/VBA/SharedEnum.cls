VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SharedEnum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Lambda.Editor.Utility"
Option Explicit

Public Enum VarNamingStyle
    PASCAL_CASE = 1
    CAMEL_CASE = 2
    SNAKE_CASE = 3
End Enum

Public Enum Group
    LAMBDA_PARAMETERS = 0
    LAMBDA_Dependencies = 1
    CUSTOM_PROPERTIES = 2
End Enum

Public Enum CloseBy
    User = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum
