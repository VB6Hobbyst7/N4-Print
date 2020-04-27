Imports System.Collections.Generic

Public Class N4JsonDocument
    Public Property company As String
    Public Property company_code As String
    Public Property document As String
    Public Property license As String
    Public Property printer As String
    Public Property start As String
    Public Property terminal As String
    Public Property ttl As Integer
    Public Property containers As New List(Of Container)

    Structure Container
        Public Property category As String
        Public Property changed As String
        Public Property changer As String
        Public Property created As String
        Public Property creator As String
        Public Property damage As String
        Public Property freightkind As String
        Public Property gross_weight As String
        Public Property iso_code As String
        Public Property iso_text As String
        Public Property line As String
        Public Property number As String
        Public Property pod As String
        Public Property position As String
        Public Property seal1 As String
        Public Property seal2 As String
        Public Property trans_type As String
        Public Property vessel_code As String
        Public Property vessel_name As String
        Public Property voy_in As String
        Public Property voy_out As String
        Public Property temperature As String
        Public Property terminal As String
        Public Property dg As String
    End Structure
End Class


