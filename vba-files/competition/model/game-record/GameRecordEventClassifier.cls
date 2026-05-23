VERSION 1.0 CLASS
BEGIN
MultiUse = -1  'True
END
Attribute VB_Name = "GameRecordEventClassifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'namespace=vba-files/competition/model/game-record
Option Explicit
Implements IObjectClassifier

Private Function IObjectClassifier_Classify(Element As Object) As Variant
    Dim pElement As GameRecordModel: Set pElement = Element
    IObjectClassifier_Classify = pElement.EventId
End Function
