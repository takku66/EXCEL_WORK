VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MF_SHAPE_DATA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim mf_sheet As Worksheet
Dim mf_shape As Shape

Dim mf_name As String
Dim mf_content As String
Dim mf_hyperlink As String
Dim mf_subhyperlink As String


Public Sub SET_SHAPE(shp As Shape)
    Set mf_shape = shp
End Sub

Public Sub SET_SHEET(sht As Worksheet)
    Set mf_sheet = sht
End Sub

Public Sub SET_NAME(name As String)
    mf_name = name
End Sub

Public Sub SET_CONTENT(content As String)
    mf_content = content
End Sub

Public Sub SET_HYPERLINK(hyperlink As String)
    mf_hyperlink = hyperlink
End Sub

Public Sub SET_SUBHYPERLINK(subhyperlink As String)
    mf_subhyperlink = subhyperlink
End Sub


Public Function GET_SHAPE()
    Set GET_SHAPE = mf_shape
End Function

Public Function GET_SHEET()
    Set GET_SHEET = mf_sheet
End Function

Public Function GET_NAME()
    GET_NAME = mf_name
End Function

Public Function GET_CONTENT()
    GET_CONTENT = mf_content
End Function

Public Function GET_HYPERLINK()
    GET_HYPERLINK = mf_hyperlink
End Function

Public Function GET_SUBHYPERLINK()
    GET_SUBHYPERLINK = mf_subhyperlink
End Function


