
Sub MakeStore()
'
' MakeStore Macro
'

    'Reeferncing Aprobados Boss
    
    Dim aprobadosBossWs As Worksheet
    Set aprobadosBossWs = ThisWorkbook.Sheets("Aprobados Boss")
    Dim aprobadosBossTable As ListObject
    Set aprobadosBossTable = aprobadosBossWs.ListObjects("Aprobados_Boss")
    
    'Getting all the indexes of column headers in AprobadosBoss
    Dim streetIdx As Integer
    Dim colonyIdx As Integer
    Dim cityIdx As Integer
    Dim stateIdx As Integer
    Dim cpIdx As Integer
    Dim contactIdx As Integer
    Dim mailIdx As Integer
    Dim phoneIdx As Integer
    
    streetIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.Street").Index
    colonyIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.Colony").Index
    cityIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.City").Index
    stateIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.State").Index
    cpIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.CP").Index
    contactIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.Contact").Index
    mailIdx = aprobadosBossTable.ListColumns("CORREO ELECTRONICO").Index
    phoneIdx = aprobadosBossTable.ListColumns("SpecialOpsThisWeek.Phone_").Index
    
    'Selecting environment
    
    Dim environmentChoice As Variant
    
    environmentChoice = Application.InputBox("Que ambiente desea utilizar? " & Chr(13) & _
    "1-FDMX (prod)" & Chr(13) & _
    "2-SCOTIAPOS (prod)" & Chr(13) & _
    "3-TEST (prod)" & Chr(13) _
    )
    
    Dim baseURL As String
    Dim apiPassword As String
    Dim apiUser As String
    Dim certificateName As String
    
    Select Case environmentChoice
        Case "1"
            MsgBox "FDMX SELECTED"
            baseURL = "https://www2.ipg-online.com/mcsWebService"
            apiPassword = "z>GiE69~sh"
            apiUser = "WST315869._.1"
            certificateName = "WST315869._.1"
        
            
        Case "2"
            MsgBox "SCOTIAPOS SELECTED"
            baseURL = "https://www2.ipg-online.com/mcsWebService"
            apiPassword = "u4P~YL2(ky"
            apiUser = "ST302603._.1"
            certificateName = "ST302603._.1"
            
        Case "3"
            MsgBox "TEST SELECTED"
            baseURL = "https://test.ipg-online.com/mcsWebService"
            apiPassword = "tester02"
            apiUser = "WSIPG"
            certificateName = "WSIPG"
            
        Case Else
            MsgBox "Please chose a valid option"
            Exit Sub
    End Select
    
    
    
    'Getting the list of selected rows
    
    Dim selectedRange As String
    Dim numOfSelectedRows As Integer
    Dim firstRowIdx As Integer
    Dim lastRowIdx
    selectedRange = Selection.Address
    numOfSelectedRows = Selection.Rows.Count
    firstRowIdx = Selection(1).Row
    lastRowIdx = firstRowIdx + numOfSelectedRows - 1
    
    For I = firstRowIdx To lastRowIdx
    
        ' We obtain the info from the row of the active cell
        Dim activeRow As Integer
        activeRow = I
        activeRow = activeRow - 2
        
        ' Getting info from CardNotPresentTable
        
        Dim merchantDBAName As String
        Dim midPrd As String
        Dim sidPrd As String
        Dim tidPrd As String
        Dim prosa As String
        Dim mccPrd As String
        Dim visaMid As String
        
        merchantDBAName = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[MerchantDBAName]].Column).Value
        midPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[MIDprd]].Column).Value
        sidPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[SIDprd]].Column).Value
        tidPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[TIDprd]].Column).Value
        prosa = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[Prosa]].Column).Value
        mccPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[SIC/MCC]].Column).Value
        visaMid = prosa
        
        If environmentChoice = "3" Then
            MsgBox "Using Test data for request.."
            midPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[MIDtest]].Column).Value
            sidPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[SIDtest]].Column).Value
            tidPrd = [NoPresentCardData].Cells(activeRow, [NoPresentCardData[TIDtest]].Column).Value
            prosa = "9900004"
            mccPrd = "5533"
            visaMid = "9900003"
            
        Else
        
            MsgBox ("Trying to find the MID of: " & merchantDBAName & " in Aprobados Boss")
            Dim matchResult As Variant
            Dim aprobadoBossMIDList As Range
            Dim lookUpMID As Double
            
            If midPrd = "" Then
             MsgBox "MID vacío"
             Exit Sub
            End If
            
            lookUpMID = CDbl(midPrd)
            
            Set rng = aprobadosBossTable.ListColumns("MID").DataBodyRange
            
            matchResult = Application.Match(lookUpMID, rng, 0)
            
            If IsError(mathcResult) Then
                ' handle error
                MsgBox "No se encontro el MID:" & lookUpMID & "en la tabla Aprobados_Boss"
                Exit Sub
                
            End If
            
            MsgBox ("Se encontro una coincidencia con el registro " & matchResult)
            Dim streetResult As String
            Dim colonyResult As String
            Dim cityResult As String
            Dim stateResult As String
            Dim cpResult As String
            Dim contactResult As String
            Dim mailResult As String
            Dim phoneResult As String
            
            streetResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, streetIdx).Value
            colonyResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, colonyIdx).Value
            cityResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, cityIdx).Value
            stateResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, stateIdx).Value
            cpResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, cpIdx).Value
            contactResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, contactIdx).Value
            mailResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, mailIdx).Value
            phoneResult = aprobadosBossTable.DataBodyRange.Cells(matchResult, phoneIdx).Value
        End If
        
        Dim dtToday As String
        dtToday = Format(Date, "yyyy-mm-dd")
        
        'Selcting Store Type

        Dim storeTypeChoice As Variant
        Dim isvChoise As Variant
        
        storeTypeChoice = Application.InputBox("Que ambiente desea utilizar? " & Chr(13) & _
        "1-Payment URL" & Chr(13) & _
        "2-Link de Pago" & Chr(13) & _
        "3-VT/MOTO" & Chr(13) & _
        "4-ISV" & Chr(13) _
        )
        
        'Creating the XML Body
        Dim soapBody As String
        Select Case storeTypeChoice
            Case "1"
                MsgBox "PaymentURL SELECTED"
                
                    'Petición PaymentURL
                    soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                    soapBody = soapBody & "   <SOAP-ENV:Header />"
                    soapBody = soapBody & "   <SOAP-ENV:Body>"
                    soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                    soapBody = soapBody & "         <ns2:createStore>"
                    soapBody = soapBody & "            <ns2:store>"
                    soapBody = soapBody & "                    <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                    soapBody = soapBody & "                    <ns2:storeAdmin>"
                    soapBody = soapBody & "                        <ns2:id>" & sidPrd & "</ns2:id>"
                    soapBody = soapBody & "                    </ns2:storeAdmin>"
                    soapBody = soapBody & "                    <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                    soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                    soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                    soapBody = soapBody & "                    <ns2:reseller>FDMX</ns2:reseller>"
                    soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                    soapBody = soapBody & "                    <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                    soapBody = soapBody & "                    <ns2:timezone>America/Mexico_City</ns2:timezone>"
                    soapBody = soapBody & "                    <ns2:status>OPEN</ns2:status>"
                    soapBody = soapBody & "                    <ns2:openDate>" & dtToday & "</ns2:openDate>"
                    soapBody = soapBody & "                    <ns2:acquirer>FDMX-MXN</ns2:acquirer>"
                    soapBody = soapBody & "                    <ns2:address>"
                    soapBody = soapBody & "                        <ns2:address1>" & streetResult & "</ns2:address1>"
                    soapBody = soapBody & "                        <ns2:zip>" & cpResult & "</ns2:zip>"
                    soapBody = soapBody & "                        <ns2:city>" & cityResult & "</ns2:city>"
                    soapBody = soapBody & "                        <ns2:state>" & stateResult & "</ns2:state>"
                    soapBody = soapBody & "                        <ns2:country>MEX</ns2:country>"
                    soapBody = soapBody & "                    </ns2:address>"
                    soapBody = soapBody & "                    <ns2:contact>"
                    soapBody = soapBody & "                        <ns2:name>" & contactResult & "</ns2:name>"
                    soapBody = soapBody & "                        <ns2:email>" & mailResult & "</ns2:email>"
                    soapBody = soapBody & "                        <ns2:phone>" & phoneResult & "</ns2:phone>"
                    soapBody = soapBody & "                    </ns2:contact>"
                    soapBody = soapBody & "                    <ns2:service>"
                    soapBody = soapBody & "                        <ns2:type>connect</ns2:type>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>overwriteURLsAllowed</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>reviewURL</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>transactionNotificationURL</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowVoid</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>responseSuccessURL</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>responseFailureURL</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>checkoutOptionCombinedPage</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>skipResultPageForFailure</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowMaestroWithout3DS</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowMOTO</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowExtendedHashCalculationWithoutCardDetails</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>cardCodeBehaviour</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>mandatory</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>hideCardBrandLogoInCombinedPage</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>skipResultPageForSuccess</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>sharedSecret</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>sharedSecret</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                    </ns2:service>"
                    soapBody = soapBody & "                    <ns2:service>"
                    soapBody = soapBody & "                        <ns2:type>creditCard</ns2:type>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>Mastercard</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>JCB</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>cardCodeMandatory</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI7_EPAS</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI7_API</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>IPaddressInAuthRequest</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>Diners</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>accountUpdaterVisa</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>Amex</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>AVS</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>0</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI7_Connect</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>Visa</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>CUP</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>mexicoLocal</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI7_VT</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowPurchaseCards</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowCredits</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                    </ns2:service>"
                    soapBody = soapBody & "                    <ns2:service>"
                    soapBody = soapBody & "                        <ns2:type>installment</ns2:type>"
                    soapBody = soapBody & "                    </ns2:service>"
                    soapBody = soapBody & "                    <ns2:service>"
                    soapBody = soapBody & "                        <ns2:type>paymentUrl</ns2:type>"
                    soapBody = soapBody & "                    </ns2:service>"
                    soapBody = soapBody & "                    <ns2:service>"
                    soapBody = soapBody & "                        <ns2:type>3dSecure</ns2:type>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>amex3DSRequestorId</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>acquirerName</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowed3dsServerVersionMax</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI7onIVRPAResUifInternational</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowSplitAuthentication</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>TRAGhostCall</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowECI1andECI6</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>amexMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>iciciOnUsVereq</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>dinersMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>mc3DS2DataOnly</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>integration</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>Modirum3DSServer</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>maestroMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowMPIViaAPI</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>vmid</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>declineSameXIDandECI</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>transactionRiskAnalysis</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>jcbMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>declineSameAAV</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>mastercardMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>" & prosa & "</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>visaMID</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>" & visaMid & "</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowAuthenticationRetry</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowMPIViaRESTAPI</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>allowSoftDeclineRetry_Connect</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>visaPassword</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>excludedCardCountries</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                        <ns2:config>"
                    soapBody = soapBody & "                            <ns2:item>jcbPassword</ns2:item>"
                    soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                    soapBody = soapBody & "                        </ns2:config>"
                    soapBody = soapBody & "                    </ns2:service>"
                    soapBody = soapBody & "                    <ns2:terminal>"
                    soapBody = soapBody & "                        <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                    soapBody = soapBody & "                        <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                    soapBody = soapBody & "                        <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                    soapBody = soapBody & "                        <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                    soapBody = soapBody & "                        <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                    soapBody = soapBody & "                        <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                    soapBody = soapBody & "                        <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                    soapBody = soapBody & "                        <ns2:submissionComponent>CONNECT</ns2:submissionComponent>"
                    soapBody = soapBody & "                        <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                    soapBody = soapBody & "                        <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                    soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                    soapBody = soapBody & "                        <ns2:active>true</ns2:active>"
                    soapBody = soapBody & "                    </ns2:terminal>"
                    soapBody = soapBody & "                    <ns2:purchaseLimit>"
                    soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                    soapBody = soapBody & "                        <ns2:limit>250000</ns2:limit>"
                    soapBody = soapBody & "                    </ns2:purchaseLimit>"
                    soapBody = soapBody & "                    <ns2:fraudSettings>"
                    soapBody = soapBody & "                        <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                    soapBody = soapBody & "                        <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                    soapBody = soapBody & "                        <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                    soapBody = soapBody & "                        <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                    soapBody = soapBody & "                        <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                    soapBody = soapBody & "                        <ns2:duplicateLockoutTimeSeconds>60</ns2:duplicateLockoutTimeSeconds>"
                    soapBody = soapBody & "                        <ns2:autoLockoutTimeSeconds>30</ns2:autoLockoutTimeSeconds>"
                    soapBody = soapBody & "                    </ns2:fraudSettings>"
                    soapBody = soapBody & "                </ns2:store>"
                    soapBody = soapBody & "         </ns2:createStore>"
                    soapBody = soapBody & "      </ns2:mcsRequest>"
                    soapBody = soapBody & "   </SOAP-ENV:Body>"
                    soapBody = soapBody & "</SOAP-ENV:Envelope>"
                    
                Case "2"
                    MsgBox "Link de Pago SELECTED"
                    
                    'Peticion Link de Pago
                    soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                    soapBody = soapBody & "   <SOAP-ENV:Header />"
                    soapBody = soapBody & "   <SOAP-ENV:Body>"
                    soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                    soapBody = soapBody & "         <ns2:createStore>"
                    soapBody = soapBody & "             <ns2:store>"
                    soapBody = soapBody & "                 <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                    soapBody = soapBody & "                 <ns2:storeAdmin>"
                    soapBody = soapBody & "                     <ns2:id>" & sidPrd & "</ns2:id>"
                    soapBody = soapBody & "                 </ns2:storeAdmin>"
                    soapBody = soapBody & "                 <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                    soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                    soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                    soapBody = soapBody & "                    <ns2:reseller>FDMX</ns2:reseller>"
                    soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                    soapBody = soapBody & "                 <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                    soapBody = soapBody & "                 <ns2:timezone>America/Mexico_City</ns2:timezone>"
                    soapBody = soapBody & "                 <ns2:status>OPEN</ns2:status>"
                    soapBody = soapBody & "                 <ns2:openDate>" & dtToday & "</ns2:openDate>"
                    soapBody = soapBody & "                 <ns2:acquirer>FDMX-MXN</ns2:acquirer>"
                    soapBody = soapBody & "                 <ns2:address>"
                    soapBody = soapBody & "                     <ns2:address1>" & streetResult & "</ns2:address1>"
                    soapBody = soapBody & "                     <ns2:address2>" & colonyResult & "</ns2:address2>"
                    soapBody = soapBody & "                     <ns2:zip>" & cpResult & "</ns2:zip>"
                    soapBody = soapBody & "                     <ns2:city>" & cityResult & "</ns2:city>"
                    soapBody = soapBody & "                     <ns2:state>" & stateResult & "</ns2:state>"
                    soapBody = soapBody & "                     <ns2:country>MEX</ns2:country>"
                    soapBody = soapBody & "                 </ns2:address>"
                    soapBody = soapBody & "                 <ns2:contact>"
                    soapBody = soapBody & "                     <ns2:name>" & contactResult & "</ns2:name>"
                    soapBody = soapBody & "                     <ns2:email>" & mailResult & "</ns2:email>"
                    soapBody = soapBody & "                     <ns2:phone>" & phoneResult & "</ns2:phone>"
                    soapBody = soapBody & "                 </ns2:contact>"
                    soapBody = soapBody & "                 <ns2:service>"
                    soapBody = soapBody & "                     <ns2:type>api</ns2:type>"
                    soapBody = soapBody & "                 </ns2:service>"
                    soapBody = soapBody & "                 <ns2:service>"
                    soapBody = soapBody & "                     <ns2:type>creditCard</ns2:type>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI7_EPAS</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI7_API</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>accountUpdaterVisa</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>AVS</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>IPaddressInAuthRequest</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>0</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>Visa</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI7_VT</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>Amex</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowCredits</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>cardCodeMandatory</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>CUP</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>Diners</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowPurchaseCards</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>JCB</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>Mastercard</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>mexicoLocal</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI7_Connect</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                 </ns2:service>"
                    soapBody = soapBody & "                 <ns2:service>"
                    soapBody = soapBody & "                     <ns2:type>hostedData</ns2:type>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>flagRecurring</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                 </ns2:service>"
                    soapBody = soapBody & "                 <ns2:service>"
                    soapBody = soapBody & "                     <ns2:type>3dSecure</ns2:type>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>declineSameXIDandECI</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>dinersMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>jcbMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>maestroMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowSplitAuthentication</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>integration</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>Modirum3DSServer</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>excludedCardCountries</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>transactionRiskAnalysis</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI1andECI6</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>visaPassword</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowECI7onIVRPAResUifInternational</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowSoftDeclineRetry_Connect</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>visaMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>" & visaMid & "</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>amexMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>mastercardMID</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>" & prosa & "</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>vmid</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowed3dsServerVersionMax</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>iciciOnUsVereq</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>mc3DS2DataOnly</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>TRAGhostCall</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowAuthenticationRetry</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>acquirerName</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowMPIViaAPI</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>allowMPIViaRESTAPI</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>true</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>amex3DSRequestorId</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>declineSameAAV</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value>false</ns2:value>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                     <ns2:config>"
                    soapBody = soapBody & "                         <ns2:item>jcbPassword</ns2:item>"
                    soapBody = soapBody & "                         <ns2:value xsi:nil=""true"" xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""/>"
                    soapBody = soapBody & "                     </ns2:config>"
                    soapBody = soapBody & "                 </ns2:service>"
                    soapBody = soapBody & "                 <ns2:terminal>"
                    soapBody = soapBody & "                     <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                    soapBody = soapBody & "                     <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                    soapBody = soapBody & "                     <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                    soapBody = soapBody & "                     <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                    soapBody = soapBody & "                     <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                    soapBody = soapBody & "                     <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                    soapBody = soapBody & "                     <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                    soapBody = soapBody & "                     <ns2:submissionComponent>API</ns2:submissionComponent>"
                    soapBody = soapBody & "                     <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                    soapBody = soapBody & "                     <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                    soapBody = soapBody & "                     <ns2:currency>MXN</ns2:currency>"
                    soapBody = soapBody & "                     <ns2:active>true</ns2:active>"
                    soapBody = soapBody & "                 </ns2:terminal>"
                    soapBody = soapBody & "                 <ns2:purchaseLimit>"
                    soapBody = soapBody & "                     <ns2:currency>MXN</ns2:currency>"
                    soapBody = soapBody & "                     <ns2:limit>250000</ns2:limit>"
                    soapBody = soapBody & "                 </ns2:purchaseLimit>"
                    soapBody = soapBody & "                 <ns2:fraudSettings>"
                    soapBody = soapBody & "                     <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                    soapBody = soapBody & "                     <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                    soapBody = soapBody & "                     <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                    soapBody = soapBody & "                     <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                    soapBody = soapBody & "                     <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                    soapBody = soapBody & "                     <ns2:duplicateLockoutTimeSeconds>0</ns2:duplicateLockoutTimeSeconds>"
                    soapBody = soapBody & "                     <ns2:autoLockoutTimeSeconds>0</ns2:autoLockoutTimeSeconds>"
                    soapBody = soapBody & "                 </ns2:fraudSettings>"
                    soapBody = soapBody & "             </ns2:store>"
                    soapBody = soapBody & "         </ns2:createStore>"
                    soapBody = soapBody & "      </ns2:mcsRequest>"
                    soapBody = soapBody & "   </SOAP-ENV:Body>"
                    soapBody = soapBody & "</SOAP-ENV:Envelope>"
                    
                Case "4"
                MsgBox "ISV SELECTED"
                    isvChoice = Application.InputBox("Que ambiente desea utilizar? " & Chr(13) & _
                    "1-INGENICO FDMX" & Chr(13) & _
                    "2-INGENICO SCOTIAPOS" & Chr(13) & _
                    "3-KARLOPAY" & Chr(13) & _
                    "4-PROFITROOM" & Chr(13) _
                    )

                    Select Case isvChoice
                        Case "1"
                
                        'Petición INGENICO FDMX
                        soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                        soapBody = soapBody & "   <SOAP-ENV:Header />"
                        soapBody = soapBody & "   <SOAP-ENV:Body>"
                        soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                        soapBody = soapBody & "         <ns2:createStore>"
                        soapBody = soapBody & "            <ns2:store>"
                        soapBody = soapBody & "                    <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                        soapBody = soapBody & "                    <ns2:storeAdmin>"
                        soapBody = soapBody & "                        <ns2:id>" & sidPrd & "</ns2:id>"
                        soapBody = soapBody & "                    </ns2:storeAdmin>"
                        soapBody = soapBody & "                    <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                        soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                        soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                        soapBody = soapBody & "                    <ns2:reseller>FDMX</ns2:reseller>"
                        soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                        soapBody = soapBody & "                    <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                        soapBody = soapBody & "                    <ns2:timezone>America/Mexico_City</ns2:timezone>"
                        soapBody = soapBody & "                    <ns2:status>OPEN</ns2:status>"
                        soapBody = soapBody & "                    <ns2:openDate>" & dtToday & "</ns2:openDate>"
                        soapBody = soapBody & "                    <ns2:acquirer>FDMX-MXN</ns2:acquirer>"
                        soapBody = soapBody & "                    <ns2:address>"
                        soapBody = soapBody & "                        <ns2:address1>" & streetResult & "</ns2:address1>"
                        soapBody = soapBody & "                        <ns2:zip>" & cpResult & "</ns2:zip>"
                        soapBody = soapBody & "                        <ns2:city>" & cityResult & "</ns2:city>"
                        soapBody = soapBody & "                        <ns2:state>" & stateResult & "</ns2:state>"
                        soapBody = soapBody & "                        <ns2:country>MEX</ns2:country>"
                        soapBody = soapBody & "                    </ns2:address>"
                        soapBody = soapBody & "                    <ns2:contact>"
                        soapBody = soapBody & "                        <ns2:name>" & contactResult & "</ns2:name>"
                        soapBody = soapBody & "                        <ns2:email>" & mailResult & "</ns2:email>"
                        soapBody = soapBody & "                        <ns2:phone>" & phoneResult & "</ns2:phone>"
                        soapBody = soapBody & "                    </ns2:contact>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>api</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>creditCard</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Mastercard</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>JCB</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>cardCodeMandatory</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_EPAS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_API</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>IPaddressInAuthRequest</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Diners</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>accountUpdaterVisa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Amex</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>AVS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>0</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_Connect</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Visa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>CUP</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>mexicoLocal</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_VT</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowPurchaseCards</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowCredits</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>installment</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>hostedData</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>flagRecurring</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>recurringPayment</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowRecurringType</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:terminal>"
                        soapBody = soapBody & "                        <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                        soapBody = soapBody & "                        <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                        soapBody = soapBody & "                        <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                        soapBody = soapBody & "                        <ns2:submissionComponent>API</ns2:submissionComponent>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:active>true</ns2:active>"
                        soapBody = soapBody & "                    </ns2:terminal>"
                        soapBody = soapBody & "                    <ns2:purchaseLimit>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:limit>250000</ns2:limit>"
                        soapBody = soapBody & "                    </ns2:purchaseLimit>"
                        soapBody = soapBody & "                    <ns2:fraudSettings>"
                        soapBody = soapBody & "                        <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                        soapBody = soapBody & "                        <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                        soapBody = soapBody & "                        <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                        soapBody = soapBody & "                        <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                        soapBody = soapBody & "                        <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                        soapBody = soapBody & "                        <ns2:duplicateLockoutTimeSeconds>60</ns2:duplicateLockoutTimeSeconds>"
                        soapBody = soapBody & "                        <ns2:autoLockoutTimeSeconds>30</ns2:autoLockoutTimeSeconds>"
                        soapBody = soapBody & "                    </ns2:fraudSettings>"
                        soapBody = soapBody & "                </ns2:store>"
                        soapBody = soapBody & "         </ns2:createStore>"
                        soapBody = soapBody & "      </ns2:mcsRequest>"
                        soapBody = soapBody & "   </SOAP-ENV:Body>"
                        soapBody = soapBody & "</SOAP-ENV:Envelope>"

                        Case "2"
                        'Petición INGENICO SCOTIAPOS
                        soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                        soapBody = soapBody & "   <SOAP-ENV:Header />"
                        soapBody = soapBody & "   <SOAP-ENV:Body>"
                        soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                        soapBody = soapBody & "         <ns2:createStore>"
                        soapBody = soapBody & "            <ns2:store>"
                        soapBody = soapBody & "                    <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                        soapBody = soapBody & "                    <ns2:storeAdmin>"
                        soapBody = soapBody & "                        <ns2:id>" & sidPrd & "</ns2:id>"
                        soapBody = soapBody & "                    </ns2:storeAdmin>"
                        soapBody = soapBody & "                    <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                        soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                        soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                        soapBody = soapBody & "                    <ns2:reseller>SCOTIAPOS</ns2:reseller>"
                        soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                        soapBody = soapBody & "                    <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                        soapBody = soapBody & "                    <ns2:timezone>America/Mexico_City</ns2:timezone>"
                        soapBody = soapBody & "                    <ns2:status>OPEN</ns2:status>"
                        soapBody = soapBody & "                    <ns2:openDate>" & dtToday & "</ns2:openDate>"
                        soapBody = soapBody & "                    <ns2:acquirer>ScotiaPOS</ns2:acquirer>"
                        soapBody = soapBody & "                    <ns2:address>"
                        soapBody = soapBody & "                        <ns2:address1>" & streetResult & "</ns2:address1>"
                        soapBody = soapBody & "                        <ns2:zip>" & cpResult & "</ns2:zip>"
                        soapBody = soapBody & "                        <ns2:city>" & cityResult & "</ns2:city>"
                        soapBody = soapBody & "                        <ns2:state>" & stateResult & "</ns2:state>"
                        soapBody = soapBody & "                        <ns2:country>MEX</ns2:country>"
                        soapBody = soapBody & "                    </ns2:address>"
                        soapBody = soapBody & "                    <ns2:contact>"
                        soapBody = soapBody & "                        <ns2:name>" & contactResult & "</ns2:name>"
                        soapBody = soapBody & "                        <ns2:email>" & mailResult & "</ns2:email>"
                        soapBody = soapBody & "                        <ns2:phone>" & phoneResult & "</ns2:phone>"
                        soapBody = soapBody & "                    </ns2:contact>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>api</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>creditCard</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Mastercard</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>JCB</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>cardCodeMandatory</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_EPAS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_API</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>IPaddressInAuthRequest</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Diners</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>accountUpdaterVisa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Amex</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>AVS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>0</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_Connect</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Visa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>CUP</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>mexicoLocal</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_VT</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowPurchaseCards</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowCredits</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>installment</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>hostedData</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>flagRecurring</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>recurringPayment</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowRecurringType</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>connect</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>overwriteURLsAllowed</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>reviewURL</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>transactionNotificationURL</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowVoid</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>responseSuccessURL</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>responseFailureURL</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xsi:nil=""true""/>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>checkoutOptionCombinedPage</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>skipResultPageForFailure</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowMaestroWithout3DS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowMOTO</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowExtendedHashCalculationWithoutCardDetails</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>cardCodeBehaviour</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>mandatory</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>hideCardBrandLogoInCombinedPage</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>skipResultPageForSuccess</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>sharedSecret</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>sharedSecret</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:terminal>"
                        soapBody = soapBody & "                        <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                        soapBody = soapBody & "                        <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                        soapBody = soapBody & "                        <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                        soapBody = soapBody & "                        <ns2:submissionComponent>CONNECT</ns2:transactionOrigin>"
                        soapBody = soapBody & "                        <ns2:submissionComponent>API</ns2:submissionComponent>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:active>true</ns2:active>"
                        soapBody = soapBody & "                    </ns2:terminal>"
                        soapBody = soapBody & "                    <ns2:purchaseLimit>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:limit>250000</ns2:limit>"
                        soapBody = soapBody & "                    </ns2:purchaseLimit>"
                        soapBody = soapBody & "                    <ns2:fraudSettings>"
                        soapBody = soapBody & "                        <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                        soapBody = soapBody & "                        <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                        soapBody = soapBody & "                        <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                        soapBody = soapBody & "                        <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                        soapBody = soapBody & "                        <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                        soapBody = soapBody & "                        <ns2:duplicateLockoutTimeSeconds>60</ns2:duplicateLockoutTimeSeconds>"
                        soapBody = soapBody & "                        <ns2:autoLockoutTimeSeconds>30</ns2:autoLockoutTimeSeconds>"
                        soapBody = soapBody & "                    </ns2:fraudSettings>"
                        soapBody = soapBody & "                </ns2:store>"
                        soapBody = soapBody & "         </ns2:createStore>"
                        soapBody = soapBody & "      </ns2:mcsRequest>"
                        soapBody = soapBody & "   </SOAP-ENV:Body>"
                        soapBody = soapBody & "</SOAP-ENV:Envelope>"

                        Case "3"
                        'Petición KARLOPAY
                        soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                        soapBody = soapBody & "   <SOAP-ENV:Header />"
                        soapBody = soapBody & "   <SOAP-ENV:Body>"
                        soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                        soapBody = soapBody & "         <ns2:createStore>"
                        soapBody = soapBody & "            <ns2:store>"
                        soapBody = soapBody & "                    <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                        soapBody = soapBody & "                    <ns2:storeAdmin>"
                        soapBody = soapBody & "                        <ns2:id>" & sidPrd & "</ns2:id>"
                        soapBody = soapBody & "                    </ns2:storeAdmin>"
                        soapBody = soapBody & "                    <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                        soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                        soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                        soapBody = soapBody & "                    <ns2:reseller>FDMX</ns2:reseller>"
                        soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                        soapBody = soapBody & "                    <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                        soapBody = soapBody & "                    <ns2:timezone>America/Mexico_City</ns2:timezone>"
                        soapBody = soapBody & "                    <ns2:status>OPEN</ns2:status>"
                        soapBody = soapBody & "                    <ns2:openDate>" & dtToday & "</ns2:openDate>"
                        soapBody = soapBody & "                    <ns2:acquirer>FDMX-MXN</ns2:acquirer>"
                        soapBody = soapBody & "                    <ns2:address>"
                        soapBody = soapBody & "                        <ns2:address1>" & streetResult & "</ns2:address1>"
                        soapBody = soapBody & "                        <ns2:zip>" & cpResult & "</ns2:zip>"
                        soapBody = soapBody & "                        <ns2:city>" & cityResult & "</ns2:city>"
                        soapBody = soapBody & "                        <ns2:state>" & stateResult & "</ns2:state>"
                        soapBody = soapBody & "                        <ns2:country>MEX</ns2:country>"
                        soapBody = soapBody & "                    </ns2:address>"
                        soapBody = soapBody & "                    <ns2:contact>"
                        soapBody = soapBody & "                        <ns2:name>" & contactResult & "</ns2:name>"
                        soapBody = soapBody & "                        <ns2:email>" & mailResult & "</ns2:email>"
                        soapBody = soapBody & "                        <ns2:phone>" & phoneResult & "</ns2:phone>"
                        soapBody = soapBody & "                    </ns2:contact>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>api</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>creditCard</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Mastercard</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>JCB</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>cardCodeMandatory</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_EPAS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_API</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>IPaddressInAuthRequest</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Diners</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>accountUpdaterVisa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Amex</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>AVS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>0</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_Connect</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Visa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>CUP</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>mexicoLocal</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_VT</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowPurchaseCards</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowCredits</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>installment</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>hostedData</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>flagRecurring</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>paymentUrl</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowRecurringType</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:terminal>"
                        soapBody = soapBody & "                        <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                        soapBody = soapBody & "                        <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                        soapBody = soapBody & "                        <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                        soapBody = soapBody & "                        <ns2:submissionComponent>API</ns2:submissionComponent>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:active>true</ns2:active>"
                        soapBody = soapBody & "                    </ns2:terminal>"
                        soapBody = soapBody & "                    <ns2:purchaseLimit>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:limit>250000</ns2:limit>"
                        soapBody = soapBody & "                    </ns2:purchaseLimit>"
                        soapBody = soapBody & "                    <ns2:fraudSettings>"
                        soapBody = soapBody & "                        <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                        soapBody = soapBody & "                        <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                        soapBody = soapBody & "                        <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                        soapBody = soapBody & "                        <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                        soapBody = soapBody & "                        <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                        soapBody = soapBody & "                        <ns2:duplicateLockoutTimeSeconds>60</ns2:duplicateLockoutTimeSeconds>"
                        soapBody = soapBody & "                        <ns2:autoLockoutTimeSeconds>30</ns2:autoLockoutTimeSeconds>"
                        soapBody = soapBody & "                    </ns2:fraudSettings>"
                        soapBody = soapBody & "                </ns2:store>"
                        soapBody = soapBody & "         </ns2:createStore>"
                        soapBody = soapBody & "      </ns2:mcsRequest>"
                        soapBody = soapBody & "   </SOAP-ENV:Body>"
                        soapBody = soapBody & "</SOAP-ENV:Envelope>"

                        Case "4"
                        'Petición PROFITROOM
                        soapBody = "<SOAP-ENV:Envelope xmlns:SOAP-ENV=""http://schemas.xmlsoap.org/soap/envelope/"">"
                        soapBody = soapBody & "   <SOAP-ENV:Header />"
                        soapBody = soapBody & "   <SOAP-ENV:Body>"
                        soapBody = soapBody & "      <ns2:mcsRequest xmlns:ns2=""http://www.ipg-online.com/mcsWebService"">"
                        soapBody = soapBody & "         <ns2:createStore>"
                        soapBody = soapBody & "            <ns2:store>"
                        soapBody = soapBody & "                    <ns2:storeID>" & sidPrd & "</ns2:storeID>"
                        soapBody = soapBody & "                    <ns2:storeAdmin>"
                        soapBody = soapBody & "                        <ns2:id>" & sidPrd & "</ns2:id>"
                        soapBody = soapBody & "                    </ns2:storeAdmin>"
                        soapBody = soapBody & "                    <ns2:mcc>" & mccPrd & "</ns2:mcc>"
                        soapBody = soapBody & "                    <ns2:legalName>" & merchantDBAName & "</ns2:legalName>"
                        soapBody = soapBody & "                    <ns2:dba>" & merchantDBAName & "</ns2:dba>"
                        soapBody = soapBody & "                    <ns2:reseller>FDMX</ns2:reseller>"
                        soapBody = soapBody & "                    <ns2:url>https://www.fiserv.com/es-mx.html</ns2:url>"
                        soapBody = soapBody & "                    <ns2:defaultCurrency>MXN</ns2:defaultCurrency>"
                        soapBody = soapBody & "                    <ns2:timezone>America/Mexico_City</ns2:timezone>"
                        soapBody = soapBody & "                    <ns2:status>OPEN</ns2:status>"
                        soapBody = soapBody & "                    <ns2:openDate>" & dtToday & "</ns2:openDate>"
                        soapBody = soapBody & "                    <ns2:acquirer>FDMX-MXN</ns2:acquirer>"
                        soapBody = soapBody & "                    <ns2:address>"
                        soapBody = soapBody & "                        <ns2:address1>" & streetResult & "</ns2:address1>"
                        soapBody = soapBody & "                        <ns2:zip>" & cpResult & "</ns2:zip>"
                        soapBody = soapBody & "                        <ns2:city>" & cityResult & "</ns2:city>"
                        soapBody = soapBody & "                        <ns2:state>" & stateResult & "</ns2:state>"
                        soapBody = soapBody & "                        <ns2:country>MEX</ns2:country>"
                        soapBody = soapBody & "                    </ns2:address>"
                        soapBody = soapBody & "                    <ns2:contact>"
                        soapBody = soapBody & "                        <ns2:name>" & contactResult & "</ns2:name>"
                        soapBody = soapBody & "                        <ns2:email>" & mailResult & "</ns2:email>"
                        soapBody = soapBody & "                        <ns2:phone>" & phoneResult & "</ns2:phone>"
                        soapBody = soapBody & "                    </ns2:contact>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>creditCard</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Mastercard</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>JCB</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>cardCodeMandatory</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_EPAS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_API</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>IPaddressInAuthRequest</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Diners</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>accountUpdaterVisa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Amex</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>AVS</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>exceedPreauthAllowedInPercent</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>0</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_Connect</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>Visa</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>CUP</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>mexicoLocal</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowOnlineAuthForRefund</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowECI7_VT</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowPurchaseCards</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowCredits</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>installment</ns2:type>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>hostedData</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>flagRecurring</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>false</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:service>"
                        soapBody = soapBody & "                        <ns2:type>recurringPayment</ns2:type>"
                        soapBody = soapBody & "                        <ns2:config>"
                        soapBody = soapBody & "                            <ns2:item>allowRecurringType</ns2:item>"
                        soapBody = soapBody & "                            <ns2:value>true</ns2:value>"
                        soapBody = soapBody & "                        </ns2:config>"
                        soapBody = soapBody & "                    </ns2:service>"
                        soapBody = soapBody & "                    <ns2:terminal>"
                        soapBody = soapBody & "                        <ns2:terminalID>" & tidPrd & "</ns2:terminalID>"
                        soapBody = soapBody & "                        <ns2:externalMerchantID>" & midPrd & "</ns2:externalMerchantID>"
                        soapBody = soapBody & "                        <ns2:endpointID>NASHVILLE MEXICO</ns2:endpointID>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MASTERCARD</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>VISA</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:paymentMethod>MEXICOLOCAL</ns2:paymentMethod>"
                        soapBody = soapBody & "                        <ns2:transactionOrigin>ECI</ns2:transactionOrigin>"
                        soapBody = soapBody & "                        <ns2:submissionComponent>API</ns2:submissionComponent>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:payerSecurityLevel>NOT EMPTY</ns2:payerSecurityLevel>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:active>true</ns2:active>"
                        soapBody = soapBody & "                    </ns2:terminal>"
                        soapBody = soapBody & "                    <ns2:purchaseLimit>"
                        soapBody = soapBody & "                        <ns2:currency>MXN</ns2:currency>"
                        soapBody = soapBody & "                        <ns2:limit>250000</ns2:limit>"
                        soapBody = soapBody & "                    </ns2:purchaseLimit>"
                        soapBody = soapBody & "                    <ns2:fraudSettings>"
                        soapBody = soapBody & "                        <ns2:binBlockProfile>99</ns2:binBlockProfile>"
                        soapBody = soapBody & "                        <ns2:checkBlockedIP>true</ns2:checkBlockedIP>"
                        soapBody = soapBody & "                        <ns2:checkBlockedClass-C>true</ns2:checkBlockedClass-C>"
                        soapBody = soapBody & "                        <ns2:checkBlockedName>true</ns2:checkBlockedName>"
                        soapBody = soapBody & "                        <ns2:checkBlockedCard>true</ns2:checkBlockedCard>"
                        soapBody = soapBody & "                        <ns2:duplicateLockoutTimeSeconds>60</ns2:duplicateLockoutTimeSeconds>"
                        soapBody = soapBody & "                        <ns2:autoLockoutTimeSeconds>30</ns2:autoLockoutTimeSeconds>"
                        soapBody = soapBody & "                    </ns2:fraudSettings>"
                        soapBody = soapBody & "                </ns2:store>"
                        soapBody = soapBody & "         </ns2:createStore>"
                        soapBody = soapBody & "      </ns2:mcsRequest>"
                        soapBody = soapBody & "   </SOAP-ENV:Body>"
                        soapBody = soapBody & "</SOAP-ENV:Envelope>"

                        Case Else
                        MsgBox "Please chose a valid option"
                        Exit Sub
                        End Select

                Case Else
                    MsgBox "Please chose a valid option"
                    Exit Sub
            End Select
            
            Dim continueConfirm
            continueConfirm = MsgBox("¿Seguro de continuar?", vbOKCancel)
            If continueConfirm = vbCancel Then
                Exit Sub
            End If
            
            MsgBox "Sending request..."
            Dim http As WinHttpRequest
            Set http = New WinHttpRequest
            
            'Los nombres de los certificados instalados pueden verse en PS
            'Get-ChildItem Cert:\CurrentUser\My | ft
            'El nombre del certificado esta en CN=
            'Tambien se pued visualizar con la harrmienta Manage user certificates
            
            With http
                .Open "POST", baseURL, False
                .SetRequestHeader "Content-type", "text/xml"
                .SetRequestHeader "Authorization", "Basic " + Base64Encode(apiUser + ":" + apiPassword)
                .SetClientCertificate certificateName
                .Send soapBody
                Debug.Print .ResponseText
                Debug.Print .Status
                Dim statusRes As String
                statusRes = "HTTP STATUS:" & .Status & Chr(13)
                RequestResults.resultsTextBox.MultiLine = True
                
                On Error GoTo ErrHandler
                    RequestResults.resultsTextBox.Text = statusRes + PrettyPrintXML(.ResponseText)
                    GoTo NextCode
ErrHandler:
                    RequestResults.resultsTextBox.Text = statusRes + .ResponseText
                
NextCode:
                
            End With
            Set xmlReq = Nothing
            Set xmlResponse = Nothing
                
        RequestResults.sidTextBox.Text = sidPrd
        RequestResults.setNodeMenu
        RequestResults.Show
        
    Next I
    
End Sub
Public Function SendRquest(baseURL As String, certName As String, user As String, pwd As String, xmlBody As String)

    With http
        .Open "POST", myURL, False
        .SetRequestHeader "Content-type", "text/xml"
        .SetRequestHeader "Authorization", "Basic " + Base64Encode(apiUser + ":" + apiPassword)
        .SetClientCertificate "WSIPG"
        .Send soapBody
        Debug.Print .ResponseText
        Debug.Print .Status
        Dim statusRes As String

End Function
Public Function PrettyPrintXML(XML As String) As String

  Dim Reader As New SAXXMLReader60
  Dim Writer As New MXXMLWriter60

  Writer.indent = True
  Writer.standalone = False
  Writer.omitXMLDeclaration = False
  Writer.Encoding = "utf-8"

  Set Reader.contentHandler = Writer
  Set Reader.dtdHandler = Writer
  Set Reader.errorHandler = Writer

  Call Reader.putProperty("http://xml.org/sax/properties/declaration-handler", _
          Writer)
  Call Reader.putProperty("http://xml.org/sax/properties/lexical-handler", _
          Writer)

  Call Reader.Parse(XML)

  PrettyPrintXML = Writer.output

End Function
Public Function Base64Encode(sText)
Dim oXML, oNode
Set oXML = CreateObject("Msxml2.DOMDocument.3.0")
Set oNode = oXML.createElement("base64")
oNode.DataType = "bin.base64"
oNode.nodeTypedValue = StringToBinary(sText)


Base64Encode = oNode.Text
Set oNode = Nothing
Set oXML = Nothing
End Function


Public Function StringToBinary(Text)
Const adTypeText = 2
Const adTypeBinary = 1

Dim BinaryStream
Set BinaryStream = CreateObject("ADODB.Stream")

BinaryStream.Type = adTypeText
BinaryStream.Charset = "us-ascii"
BinaryStream.Open
BinaryStream.WriteText Text

'Change stream type To binary
BinaryStream.Position = 0
BinaryStream.Type = adTypeBinary

'Ignore first two bytes - sign of
BinaryStream.Position = 0

StringToBinary = BinaryStream.Read

Set BinaryStream = Nothing
End Function

Function RandomString(Length As Integer) As String
'PURPOSE: Create a Randomized String of Characters
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'Passwords must contain at least one alphabetical letter, one digit, one special character, a maximum of two repeating characters and a length between 8 and 32 characters
Dim CharacterBank As Variant
Dim DigitsBank As Variant
Dim SpecialCharBank As Variant
Dim AllChars As Variant

Dim x As Long
Dim str As String

'Test Length Input
  If Length < 1 Then
    MsgBox "Length variable must be greater than 8"
    Exit Function
  End If

CharacterBank = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z")
  
DigitsBank = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9")

SpecialCharBank = Array("!", "@", "#", "$", "%", "^", "&", "*")

AllChars = Array("a", "b", "c", "d", "e", "f", "g", "h", "i", "j", _
  "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", _
  "y", "z", "A", "B", "C", "D", "E", "F", "G", "H", _
  "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", _
  "W", "X", "Y", "Z", _
  "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", _
  "!", "@", "#", "$", "%", "^", "&", "*")
  

str = str & CharacterBank(Int((UBound(CharacterBank) - LBound(CharacterBank) + 1) * Rnd + LBound(CharacterBank)))
str = str & DigitsBank(Int((UBound(DigitsBank) - LBound(DigitsBank) + 1) * Rnd + LBound(DigitsBank)))
str = str & SpecialCharBank(Int((UBound(SpecialCharBank) - LBound(SpecialCharBank) + 1) * Rnd + LBound(SpecialCharBank)))

'Randomly Select Characters One-by-One
Dim randomChar As String
  For x = 1 To Length - 3
    Randomize
    
    Dim charIsValid As Boolean
    charIsValid = False
    
    While Not charIsValid
        randomChar = AllChars(Int((UBound(AllChars) - LBound(AllChars) + 1) * Rnd + LBound(AllChars)))
        If InStr(str, randomChar) = 0 Then
            charIsValid = True
        End If
    Wend
    
    str = str & randomChar
    
  Next x

'Output Randomly Generated String
  RandomString = str

End Function