Attribute VB_Name = "CustomError"
'namespace=vba-files/competition/util
Option Explicit
Option Private Module

Public Enum CustomErrorCodeEnum
    UndefinedError = vbObjectError + 256
    DuplicatedRound
    NotApplicated
    TemplateFileNotFound
    EventFileNotFound
End Enum
