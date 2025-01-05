Imports Microsoft.VisualBasic

Public Class JSON_GGL_ResponseGeocoding
    Public results() As JSON_GGL_ResponseGeocoding_Result
    Public status As String = ""
End Class

Public Class JSON_GGL_ResponseGeocoding_Result
    Public address_components() As JSON_GGL_ResponseGeocoding_Result_AddressComponent
    Public formatted_address As String = ""
    Public geometry As New JSON_GGL_ResponseGeocoding_Result_Geometry
    Public place_id As String = ""
    Public plus_code As New JSON_GGL_ResponseGeocoding_Result_PlusCode
    Public types As String()
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_AddressComponent
    Public long_name As String = ""
    Public short_name As String = ""
    Public types As String()
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_Geometry
    Public location As New JSON_GGL_ResponseGeocoding_Result_Geometry_Location
    Public location_type As String
    Public viewport As New JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_Geometry_Location
    Public lat As Double = 0
    Public lng As Double = 0
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport
    Public northeast As New JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport_Northeast
    Public southwest As New JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport_Southwest
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport_Northeast
    Public lat As Double = 0
    Public lng As Double = 0
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_Geometry_Viewport_Southwest
    Public lat As Double = 0
    Public lng As Double = 0
End Class

Public Class JSON_GGL_ResponseGeocoding_Result_PlusCode
    Public compound_code As String = ""
    Public global_code As String = ""
End Class