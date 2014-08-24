Option Explicit
Dim myOlApp As Outlook.Application
Public WithEvents myOlItems As Outlook.Items
Private oSoapClient As SoapClient30
Private Const ListView As String = "{9C7F2AC0-05F2-4506-B5C0-1DFF907F3902}"
Private Const ListID As String = "{785F9A6A-2C91-4264-99EB-26ED6379515C}"



Public Sub AddToSharePoint(ByVal Subject As String, ByVal Location As String, ByVal MeetingDate As String, ByVal Details As String)
    Dim XmlDoc As New DOMDocument
    Dim batch As IXMLDOMElement
    Dim BatchXML As String
    
    Set batch = XmlDoc.createElement("Batch")
    
    BatchXML = "<Batch OnError='continue' ListVersion='1' ViewName='" & ListView & "'>"
    BatchXML = BatchXML & "<Method ID='1' Cmd='New'>"
    
    BatchXML = BatchXML & "<Field Name='Title'>" & Subject & "</Field>"
    BatchXML = BatchXML & "<Field Name='Location'>" & Location & "</Field>"
    BatchXML = BatchXML & "<Field Name='MeetingDate'>" & MeetingDate & "</Field>"
    BatchXML = BatchXML & "<Field Name='Details'>" & Details & "</Field>"
    
    BatchXML = BatchXML & "</Method></Batch>"
        
    Set oSoapClient = New SoapClient30
    Call oSoapClient.MSSoapInit(par_WSDLFile:="http://intranet.nnfcc.local/_vti_bin/Lists.asmx?WSDL")
        
    Call oSoapClient.UpdateListItems(ListID, BatchXML)
    
    Set oSoapClient = Nothing
    
End Sub

Private Sub myOlItems_ItemAdd(ByVal Item As Object)
    Dim ThisItem As AppointmentItem
    Set ThisItem = Item
    Dim DateString As String
    DateString = Format(ThisItem.Start, "yyyy-MM-ddTHH:mm:ssZ")
    
    If ThisItem.Categories = "NNFCC Meeting" Then
        Call AddToSharePoint(ThisItem.Subject, ThisItem.Location, DateString, ThisItem.Body)
    Else
    End If
    
End Sub

Public Sub Application_Startup()
   Set myOlItems = Outlook.GetNamespace("MAPI").GetDefaultFolder(olFolderCalendar).Items
End Sub
