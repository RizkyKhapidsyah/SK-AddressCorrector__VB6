Attribute VB_Name = "modMain"
' Requires Microsoft XML SDK 3.0 available at msdn.microsoft.com.

Public Function AddrCorrect(ByRef address As String, ByRef city As String, ByRef state As String, ByRef zip As String, Optional LicenseKey As String) As String
  Dim oXMLHTTP As Object

  
  ' Call the web service to get an XML document
  Set oXMLHTTP = CreateObject("Msxml2.ServerXMLHTTP")

  oXMLHTTP.Open "POST", _
                "http://ws.cdyne.com/psaddress/addresslookup.asmx/CheckAddress", _
                False
  oXMLHTTP.setRequestHeader "Content-Type", _
                            "application/x-www-form-urlencoded"
  oXMLHTTP.send "AddressLine=" & URLEncode(address) & "&ZipCode=" & URLEncode(zip) & "&City=" & URLEncode(city) & "&StateAbbrev=" & URLEncode(state) & "&LicenseKey=" & URLEncode(LicenseKey)
  
  If oXMLHTTP.Status <> 200 Then
    MsgBox "Service Unavailable. Try again later"
    Set oXMLHTTP = Nothing

    Exit Function
  End If
  Dim oDOM As Object
  
  Set oDOM = oXMLHTTP.responseXML
  Dim oNL As Object
  Dim oCN As Object
  Dim oCC As Object
  Set oNL = oDOM.getElementsByTagName("Address")
  For Each oCN In oNL
    For Each oCC In oCN.childNodes
        Select Case LCase(oCC.nodeName)
            Case "serviceerror"
                If CBool(oCC.Text) = True Then
                    AddrCorrect = "Service Error.  Try again Later"
                    GoTo leaveit
                End If
            Case "addresserror"
                If CBool(oCC.Text) = True Then
                    AddrCorrect = "Address uncorrectable."
                    GoTo leaveit
                End If
            Case "servicecurrentlyunavailable"
                If CBool(oCC.Text) = True Then
                    AddrCorrect = "Service Unavailable.  Try again Later"
                    GoTo leaveit
                End If
            Case "addressfoundbemorespecific"
                If CBool(oCC.Text) = True Then
                    AddrCorrect = "Address Found.  Be more Specific."
                    GoTo leaveit
                End If
            Case "deliveryaddress"
                address = oCC.Text
            Case "city"
                city = oCC.Text
            Case "stateabbrev"
                state = oCC.Text
            Case "zipcode"
                zip = oCC.Text
        End Select
    Next
  Next
AddrCorrect = "OK" ' Address corrected
   
 
leaveit:
  Set oCC = Nothing
  Set oCN = Nothing
  Set oNL = Nothing
  Set oDOM = Nothing
  Set oXMLHTTP = Nothing

End Function

Public Function URLEncode(inS As String) As String
  Dim i As Long
  Dim inC, outC As String
  For i = 1 To Len(inS)
    inC = Mid(inS, i, 1)
    Select Case inC
      Case " "
       outC = "+"
      Case "&"
       outC = "%38"
      Case "!" To "~"
       outC = inC
      Case Else
       outC = "%" + Right("00" + Hex(Asc(inC)), 2)
    End Select
    URLEncode = URLEncode + outC
  Next i
End Function
