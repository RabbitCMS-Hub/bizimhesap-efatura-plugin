<%
'**********************************************
'**********************************************
'               _ _                 
'      /\      | (_)                
'     /  \   __| |_  __ _ _ __  ___ 
'    / /\ \ / _` | |/ _` | '_ \/ __|
'   / ____ \ (_| | | (_| | | | \__ \
'  /_/    \_\__,_| |\__,_|_| |_|___/
'               _/ | Digital Agency
'              |__/ 
' 
'* Project  : RabbitCMS
'* Developer: <Anthony Burak DURSUN>
'* E-Mail   : badursun@adjans.com.tr
'* Corp     : https://adjans.com.tr
'**********************************************
' LAST UPDATE: 28.10.2022 15:33 @badursun
'**********************************************

Class bizimhesap_efatura_plugin
	Private PLUGIN_CODE, PLUGIN_DB_NAME, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_CREDITS, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT

	Private BIZIMHESAP_ACTIVE, BIZIM_API_KEY, BIZIM_FATURA_TIPI, BIZIM_FATURA_DATAPACK, BIZIM_API_BASE
	Private PARAM_FATURA_ACIKLAMASI, SIPARIS_NO, SIPARIS_ID
	Private BIZIM_USERID, BIZIM_UNVAN, BIZIM_VERGI_DAIRESI, BIZIM_VERGI_NO, BIZIM_MUSTERI_EPOSTA
	Private BIZIM_MUSTERI_TELEFON, BIZIM_MUSTERI_ADRES
	Private BIZIM_HESAP_PARABIRIMI, BIZIM_HESAP_SEPET_TOPLAM, BIZIM_HESAP_INDIRIM_TUTARI, BIZIM_HESAP_ARATOPLAM, BIZIM_HESAP_VERGITUTARI, BIZIM_HESAP_GENELTOPLAM
    Private SEPET_COUNT
    Dim SEPET()


	'---------------------------------------------------------------
	' Register Class
	'---------------------------------------------------------------
	Public Property Get class_register()
		DebugTimer ""& PLUGIN_CODE &" class_register() Start"
		
		' Check Register
		'------------------------------
		If CheckSettings("PLUGIN:"& PLUGIN_CODE &"") = True Then 
			DebugTimer ""& PLUGIN_CODE &" class_registered"
			Exit Property
		End If

		' Check And Create Table
		'------------------------------
		Dim PluginTableName
			PluginTableName = "tbl_plugin_" & PLUGIN_DB_NAME

    	If TableExist(PluginTableName) = False Then
    		Conn.Execute("SET NAMES utf8mb4;") 
    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 0;") 
    		
    		Conn.Execute("DROP TABLE IF EXISTS `"& PluginTableName &"`")

    		q=""
    		q=q+"CREATE TABLE `"& PluginTableName &"` ( "
    		q=q+"  `ID` int(11) NOT NULL AUTO_INCREMENT, "
    		q=q+"  `ORDER_ID` int(11) DEFAULT 0, "
    		q=q+"  `ORDER_NO` bigint(20) DEFAULT 0, "
    		q=q+"  `GUID` varchar(100) DEFAULT NULL, "
    		q=q+"  `MSG` varchar(255) DEFAULT NULL, "
    		q=q+"  `FATURA_TARIHI` datetime DEFAULT current_timestamp(), "
    		q=q+"  `FATURA_URL` varchar(255) DEFAULT NULL, "
    		q=q+"  PRIMARY KEY (`ID`), "
    		q=q+"  KEY `IND1` (`ORDER_ID`) "
    		q=q+") ENGINE=MyISAM DEFAULT CHARSET=utf8; "
			Conn.Execute(q)

    		Conn.Execute("SET FOREIGN_KEY_CHECKS = 1;") 

			' Create Log
			'------------------------------
    		Call PanelLog(""& PLUGIN_CODE &" için database tablosu oluşturuldu", 0, ""& PLUGIN_CODE &"", 0)

			' Register Settings
			'------------------------------
			DebugTimer ""& PLUGIN_CODE &" class_register() End"
    	End If

		' Register Settings
		'------------------------------
		a=GetSettings("PLUGIN:"& PLUGIN_CODE &"", PLUGIN_CODE)
		a=GetSettings(""&PLUGIN_CODE&"_PLUGIN_NAME", PLUGIN_NAME)
		a=GetSettings(""&PLUGIN_CODE&"_CLASS", "bizimhesap_efatura_plugin")
		a=GetSettings(""&PLUGIN_CODE&"_REGISTERED", ""& Now() &"")
		a=GetSettings(""&PLUGIN_CODE&"_CODENO", "830")
		a=GetSettings(""&PLUGIN_CODE&"_ACTIVE", "1")
		a=GetSettings(""&PLUGIN_CODE&"_API_KEY", "")
		a=GetSettings(""&PLUGIN_CODE&"_FOLDER", "bizimhesap-efatura-plugin")

		' Register Settings
		'------------------------------
		DebugTimer ""& PLUGIN_CODE &" class_register() End"
	End Property
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Public sub LoadPanel()
		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "AJAX:RemoveLog" Then
    		Query.PageContentType = "json"
    		
    		Call AdminSessionChecker()

    		REC_ID = Query.Data("RecID")

		    Conn.Execute("DELETE FROM tbl_plugin_bizimhesap WHERE ID="& REC_ID &"")
		    
		    Query.jsonResponse 200, "Güncellendi"
    		
    		Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Sub Page 
		'--------------------------------------------------------
		If Query.Data("Page") = "BizimHesapLog" Then
			Call PluginPage("Header")

			With Response 
				.Write "<div class=""table-responsive"">"
				.Write "	<table class=""table table-striped table-bordered"">"
				.Write "		<thead>"
				.Write "			<tr>"
				.Write "				<th>Sipariş ID</th>"
				.Write "				<th>Sipariş No</th>"
				.Write "				<th>GUID / Fatura URL</th>"
				.Write "				<th>Servis Cevabı</th>"
				.Write "				<th>İşlem Tarihi</th>"
				.Write "				<th></th>"
				.Write "			</tr>"
				.Write "		</thead>"
				.Write "		<tbody>"
				Set Siteler = Conn.Execute("SELECT * FROM tbl_plugin_bizimhesap ORDER BY ID DESC")
				If Siteler.Eof Then 
				    Response.Write "<tr>"
				        Response.Write "<td colspan=""5"" align=""center"">İşlem Geçmişi Bulunamadı</td>"
				    Response.Write "</tr>"
				End If
				Do While Not Siteler.Eof
				.Write "			<tr>"
				.Write "				<td>"& Siteler("ORDER_ID") &"</td>"
				.Write "				<td>"& Siteler("ORDER_NO") &"</td>"
				.Write "				<td>"
				.Write "					<div>"& Siteler("GUID") &"</div>"
				.Write "					<div><small><a href="""& Siteler("FATURA_URL") &""" target=""_blank"">"& Siteler("FATURA_URL") &"</a></small>"
				.Write "				</td>"
				.Write "				<td>"& Siteler("MSG") &"</td>"
				.Write "				<td>"& Siteler("FATURA_TARIHI") &"</td>"
				.Write "				<td align=""right"">"
				.Write "					<a data-ajax=""true"" data-remove=""tr"" href=""/panel/ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=RemoveLog&RecID="& Siteler("ID") &""" class=""btn btn-sm btn-danger"">"
				.Write "						Sil"
				.Write "					</a>"
				.Write "				</td>"
				.Write "			</tr>"
				Siteler.MoveNext : Loop
				Siteler.Close : Set Siteler = Nothing
				.Write "		</tbody>"
				.Write "	</table>"
				.Write "</div>"
				.Write "<script defer src=""/panel/vendors/bower_components/remarkable-bootstrap-notify/bootstrap-notify.min.js""></script>"
				.Write "<script src=""/panel/vendors/bower_components/jquery/dist/jquery.min.js""></script>"
				.Write "<script defer src=""/panel/js/RabbitJSVendor.js""></script>"
				.Write "<script defer src=""/panel/js/rabbit-alert.js""></script>"
				.Write "<script defer src=""/panel/js/rabbit-app.js""></script>"
				.Write "<script defer src=""/panel/js/custom.js""></script>"
			End With

			Call PluginPage("Footer")
			Call SystemTeardown("destroy")
		End If

		'--------------------------------------------------------
		' Main Page
		'--------------------------------------------------------
		With Response
			'------------------------------------------------------------------------------------------
				PLUGIN_PANEL_MASTER_HEADER This()
			'------------------------------------------------------------------------------------------
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write  		QuickSettings("select", ""&PLUGIN_CODE&"_FATURA_TIPI", "Fatura Tipi", "3#Satış Faturası Kes|5#Alış Faturası Kes", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""&PLUGIN_CODE&"_NIHAI_VERGINO", "Nihai Tüketiciye Kesilecek TC No", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""&PLUGIN_CODE&"_API_KEY", "API Anahtarı", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-6 col-sm-12"">"
			.Write 			QuickSettings("input", ""&PLUGIN_CODE&"_API_BASE", "API Base URL", "", TO_DB)
			.Write "    </div>"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write 			QuickSettings("textarea", ""&PLUGIN_CODE&"_FATURA_ACIKLAMASI", "Fatura Açıklamasına Ek Metin", "", TO_DB)
			.Write "    </div>"
			.Write "</div>"
			.Write "<div class=""row"">"
			.Write "    <div class=""col-lg-12 col-sm-12"">"
			.Write "        <a open-iframe href=""ajax.asp?Cmd=PluginSettings&PluginName="& PLUGIN_CODE &"&Page=BizimHesapLog"" class=""btn btn-sm btn-primary"">"
			.Write "        	Önbelleklenmiş Dosyaları Göster"
			.Write "        </a>"
			.Write "    </div>"
			.Write "</div>"
		End With
	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	private sub class_initialize()
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------
    	PLUGIN_NAME 			= "Bizim Hesap - E-Fatura Plugin"
    	PLUGIN_CODE 			= "BIZIM_HESAP_EFATURA"
    	PLUGIN_DB_NAME 			= "bizimhesap"
    	PLUGIN_VERSION 			= "1.0.0"
    	PLUGIN_CREDITS 			= "@badursun Anthony Burak DURSUN"
    	PLUGIN_GIT 				= "https://github.com/RabbitCMS-Hub/bizimhesap-efatura-plugin"
    	PLUGIN_DEV_URL 			= "https://adjans.com.tr"
    	PLUGIN_FILES_ROOT 		= PLUGIN_VIRTUAL_FOLDER(This)
    	'-------------------------------------------------------------------------------------
    	' PluginTemplate Main Variables
    	'-------------------------------------------------------------------------------------

    	BIZIMHESAP_ACTIVE 			= Cint( GetSettings(""&PLUGIN_CODE&"_ACTIVE", 1) )
    	BIZIM_API_KEY 				= GetSettings(""&PLUGIN_CODE&"_API_KEY", "")
    	BIZIM_FATURA_TIPI 			= GetSettings(""&PLUGIN_CODE&"_FATURA_TIPI", "3")
    	BIZIM_API_BASE 				= GetSettings(""&PLUGIN_CODE&"_API_BASE", "https://bizimhesap.com/api/")
    	PARAM_FATURA_ACIKLAMASI 	= GetSettings(""&PLUGIN_CODE&"_FATURA_ACIKLAMASI", "")
    	BIZIM_FATURA_DATAPACK 		= ""
    	SIPARIS_NO 					= 0
    	SIPARIS_ID 					= 0
    	BIZIM_USERID 				= 0
    	BIZIM_UNVAN 				= ""
    	BIZIM_VERGI_DAIRESI 		= ""
    	BIZIM_VERGI_NO 				= GetSettings(""&PLUGIN_CODE&"_NIHAI_VERGINO", "11111111111")
    	BIZIM_MUSTERI_EPOSTA 		= ""
    	BIZIM_MUSTERI_TELEFON 		= ""
    	BIZIM_MUSTERI_ADRES 		= ""
    	BIZIM_HESAP_PARABIRIMI 		= "0.00"
    	BIZIM_HESAP_SEPET_TOPLAM 	= "0.00"
    	BIZIM_HESAP_INDIRIM_TUTARI 	= "0.00"
    	BIZIM_HESAP_ARATOPLAM 		= "0.00"
    	BIZIM_HESAP_VERGITUTARI 	= "0.00"
    	BIZIM_HESAP_GENELTOPLAM 	= "0.00"
    	
    	class_register()
	end sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------
	Private sub class_terminate()

	End Sub
	'---------------------------------------------------------------
	'
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------
	Public Property Get PluginCredits()
		PluginCredits = PLUGIN_CREDITS
	End Property

	Public Property Get PluginCode()
		PluginCode = PLUGIN_CODE
	End Property

	Public Property Get PluginName()
		PluginName = PLUGIN_NAME
	End Property

	Public Property Get PluginVersion()
		PluginVersion = PLUGIN_VERSION
	End Property
	Public Property Get PluginGit()
		PluginGit = PLUGIN_GIT
	End Property
	Public Property Get PluginDevURL()
		PluginDevURL = PLUGIN_DEV_URL
	End Property
	Private Property Get This()
		This = Array(PLUGIN_CODE, PLUGIN_NAME, PLUGIN_VERSION, PLUGIN_GIT, PLUGIN_DEV_URL, PLUGIN_FILES_ROOT)
	End Property
	Public Property Get PluginFolder()
		PluginFolder = PLUGIN_FILES_ROOT
	End Property
	'---------------------------------------------------------------
	' Plugin Defines
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Property Get XMLHttp(Uri, xType, Data, AuthType)
		If BIZIMHESAP_ACTIVE = 0 Then Exit Property
		If Not Len(BIZIM_API_KEY) = 32 Then 
				CreateLog "BizimHesapEFatura.XMLHttp", "API Key Required", "API Key Required", 403, "GET"
			Exit Property
		End If

		' ByPass Error
		'------------------------------------------------
		On Error Resume Next

		' Send Data
		'------------------------------------------------
	    Dim objXMLhttp : Set objXMLhttp = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
			objXMLhttp.open xType, Uri, false
            objXMLhttp.setOption(2) = SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS
            objXMLhttp.setTimeouts 5000, 5000, 10000, 10000 'ms
			' objXMLhttp.setRequestHeader "X-IBM-Client-Id" 		, MNGKARGO_CLIENTID
			objXMLhttp.setRequestHeader "Accept" 				, "application/json"
			objXMLhttp.setRequestHeader "Content-type" 			, "application/json"
			objXMLhttp.setRequestHeader "CharSet" 				, "UTF-8"
			objXMLhttp.setRequestHeader "Content-Length" 		, Len(Data)
			objXMLhttp.send TurkceKarakter2HTML( Data )
			' objXMLhttp.send Data
			
			' CreateLog "BizimHesapEFatura.XMLHttp", jSONReplace(TurkceKarakter2HTML(Data)), jSONReplace(TurkceKarakter2HTML(objXMLhttp.responseText)), objXMLhttp.Status, xType
	    
			XMLHttp = Array(objXMLhttp.Status, objXMLhttp.responseText)
	    Set objXMLhttp = Nothing
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Let FaturaAciklamasi(Val) 		: PARAM_FATURA_ACIKLAMASI = Val 		: End Property
	Public Property Let MusteriID(Val) 				: BIZIM_USERID = Val 					: End Property
	Public Property Let MusteriUnvan(Val) 			: BIZIM_UNVAN = Val 					: End Property
	Public Property Let MusteriVergiDairesi(Val)	: BIZIM_VERGI_DAIRESI = Val 			: End Property
	Public Property Let MusteriVergiNo(Val)
		If Len(Val) = 0 Then 
			Exit Property
		End If
		BIZIM_VERGI_NO = Val
	End Property
	Public Property Let MusteriEPosta(Val) 			: BIZIM_MUSTERI_EPOSTA = Val 			: End Property
	Public Property Let MusteriTelefon(Val) 		: BIZIM_MUSTERI_TELEFON = Val 			: End Property
	Public Property Let MusteriAdres(Val) 			: BIZIM_MUSTERI_ADRES = Val 			: End Property
	Public Property Let FaturaParaBirimi(Val) 		: BIZIM_HESAP_PARABIRIMI = Val 			: End Property
	Public Property Let FaturaSepetToplami(Val) 	: BIZIM_HESAP_SEPET_TOPLAM = Val 		: End Property
	Public Property Let FaturaIndirimTutari(Val) 	: BIZIM_HESAP_INDIRIM_TUTARI = Val 		: End Property
	Public Property Let FaturaAraToplam(Val) 		: BIZIM_HESAP_ARATOPLAM = Val 			: End Property
	Public Property Let FaturaVergiToplami(Val) 	: BIZIM_HESAP_VERGITUTARI = Val 		: End Property
	Public Property Let FaturaGenelToplam(Val) 		: BIZIM_HESAP_GENELTOPLAM = Val 		: End Property
	Public Property Let SiparisID(Val) 				: SIPARIS_ID = Val 						: End Property
	Public Property Let SiparisNo(Val) 				: SIPARIS_NO = Val 						: End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Property Get CreateInvoice()
	    i=0
	    Set oJSON = New aspJSON
	        With oJSON.data
	        	.Add "firmId" 					, BIZIM_API_KEY
	        	.Add "invoiceNo" 				, ""
	        	.Add "invoiceType" 				, BIZIM_FATURA_TIPI
	        	.Add "note" 					, ""& SIPARIS_NO &" Numaralı Sipariş. " & PARAM_FATURA_ACIKLAMASI 
	        	.Add "dates" 					, oJSON.Collection()
	        	With oJSON.data("dates")
	        		.Add "invoiceDate" 			, InvoiceDate(Now()) ' Fatura tarihi
	        		.Add "deliveryDate" 		, "" ' Ödeme vadesi
	        		.Add "dueDate" 				, InvoiceDate(Now()) 'Teslimat tarihi  (opsiyonel) 
	        	End With
	        	.Add "customer" 				, oJSON.Collection()
	        	With oJSON.data("customer")
	        		.Add "customerId" 			, BIZIM_USERID ' DB User ID
	        		.Add "title" 				, BIZIM_UNVAN 'DB User Ünvan
	        		.Add "taxOffice" 			, BIZIM_VERGI_DAIRESI 'DB User Vergi Dairesi 
	        		.Add "taxNo" 				, BIZIM_VERGI_NO 'DB User Vergi No
	        		.Add "email" 				, BIZIM_MUSTERI_EPOSTA 'DB User EPosta
	        		.Add "phone" 				, BIZIM_MUSTERI_TELEFON 'DB User Telefon 
	        		.Add "address" 				, BIZIM_MUSTERI_ADRES 'DB User Fatura Adresi
	        	End With
	        	.Add "amounts" 					, oJSON.Collection()
	        	With oJSON.data("amounts")
	        		.Add "currency" 			, BIZIM_HESAP_PARABIRIMI 		' Teslimat
	        		.Add "gross" 				, BIZIM_HESAP_SEPET_TOPLAM 		' Teslimat
	        		.Add "discount" 			, BIZIM_HESAP_INDIRIM_TUTARI 	' Teslimat
	        		.Add "net" 					, BIZIM_HESAP_ARATOPLAM 		' Teslimat
	        		.Add "tax" 					, BIZIM_HESAP_VERGITUTARI 		' Teslimat
	        		.Add "total" 				, BIZIM_HESAP_GENELTOPLAM 		' Teslimat
	        	End With

	        	.Add "details" 			, oJSON.Collection()
	        	Dim BasketItems
	        		BasketItems = GetCardData()

	        	With oJSON.data("details")
		            If (NOT IsEmpty(BasketItems)) Then
		                Dim pElements(), Cursor : Cursor = 0 : ReDim pElements(UBOUND(BasketItems))
		                Dim cartItem
		                For Each cartItem In BasketItems
		                    .Add Cursor, oJSON.Collection()
		                    With .item(Cursor)
		                        .Add "productId" 		, ""& cartItem.Data.Item("productId") &""
		                        .Add "productName"		, ""& cartItem.Data.Item("productName") &""
		                        .Add "note"				, ""& cartItem.Data.Item("note") &""
		                        .Add "barcode"			, ""& cartItem.Data.Item("barcode") &""
		                        .Add "taxRate"			, ""& cartItem.Data.Item("taxRate") &""
		                        .Add "quantity"			, ""& cartItem.Data.Item("quantity") &""
		                        .Add "unitPrice"		, ""& cartItem.Data.Item("unitPrice") &""
		                        .Add "grossPrice"		, ""& cartItem.Data.Item("grossPrice") &""
		                        .Add "discount"			, ""& cartItem.Data.Item("discount") &""
		                        .Add "net"				, ""& cartItem.Data.Item("net") &""
		                        .Add "tax"				, ""& cartItem.Data.Item("tax") &""
		                        .Add "total"			, ""& cartItem.Data.Item("total") &""
		                    End With
		                    Cursor = Cursor + 1
		                Next
		            End If
	        	End With

	        End With
	        BIZIM_FATURA_DATAPACK = oJSON.JSONoutput()
	    Set oJSON = Nothing

	    CreateInvoice = BIZIM_FATURA_DATAPACK
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' Return StatusCode, Title-Guid, Error-Message, Null-URL
	'---------------------------------------------------------------
	Public Property Get SendInvoice()
		'
		'---------------------------------------
		If InvoiceExist(SIPARIS_ID) = True Then
			CreateLog "BizimHesapEFatura.SendInvoice", jSONReplace(BIZIM_FATURA_DATAPACK), "Duplicate Invoice", 400, "POST"

			SendInvoice = Array(400, "Duplicate Invoice", "Bu Fatura Zaten Kesilmiş", "") 
			
			Exit Property
		End If

		' Send Reuest
		'---------------------------------------
	    Dim InvoiceResults
	    	InvoiceResults = XMLHttp(BIZIM_API_BASE & "b2b/addinvoice", "POST", BIZIM_FATURA_DATAPACK, "")

		' Return Result
		'---------------------------------------
	    If InvoiceResults(0) = 200 Then 
			Set parseJsonData = New aspJSON
				parseJsonData.loadJSON( InvoiceResults(1) )

			Dim RESP_ERR, RESP_GUID, RESP_URL
				RESP_ERR 	= JSONTurkish( parseJsonData.data("error") )
				RESP_GUID 	= Trim( parseJsonData.data("guid") )
				RESP_URL 	= Trim( parseJsonData.data("url") )
			
			If Len(RESP_ERR) = 0 Then
				SendInvoice = Array(200, "Server Error", RESP_GUID, RESP_URL)

				CreateInvoiceLog SIPARIS_ID, RESP_GUID, RESP_ERR, RESP_URL, SIPARIS_NO
			Else 
				SendInvoice = Array(203, "Invoice Error", RESP_ERR, Null)
			End If

			CreateLog "BizimHesapEFatura.SendInvoice", JSONTurkish(BIZIM_FATURA_DATAPACK), JSONTurkish( InvoiceResults(1) ), 200, "POST"

			Set parseJsonData = Nothing
		Else 
			CreateLog "BizimHesapEFatura.SendInvoice", JSONTurkish( BIZIM_FATURA_DATAPACK ), JSONTurkish( InvoiceResults(1) ), JSONTurkish( InvoiceResults(0) ), "POST"

			SendInvoice = Array(500, "Server Error", JSONTurkish( InvoiceResults(1) ), Null)
	    End If
	End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Public Function JSONTurkish(vData)
		JSONTurkish = Trim( jSonTurkceKarakter( jsEncode(vData) ) )
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
    Public Function AddToCart(productId, productName, note, barcode, taxRate, quantity, unitPrice, grossPrice, discount, net, tax, total)
        ReDim PRESERVE SEPET(SEPET_COUNT)
        SEPET_COUNT=SEPET_COUNT+1

        Set SEPET(SEPET_COUNT-1) = New CartMaker
			SEPET( SEPET_COUNT-1 ).productId 	= productId
			SEPET( SEPET_COUNT-1 ).productName 	= productName
			SEPET( SEPET_COUNT-1 ).note 		= note
			SEPET( SEPET_COUNT-1 ).barcode 		= barcode
			SEPET( SEPET_COUNT-1 ).taxRate 		= taxRate
			SEPET( SEPET_COUNT-1 ).quantity 	= quantity
			SEPET( SEPET_COUNT-1 ).unitPrice 	= unitPrice
			SEPET( SEPET_COUNT-1 ).grossPrice 	= grossPrice
			SEPET( SEPET_COUNT-1 ).discount 	= discount
			SEPET( SEPET_COUNT-1 ).net 			= net
			SEPET( SEPET_COUNT-1 ).tax 			= tax
			SEPET( SEPET_COUNT-1 ).total 		= total
    End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
    Public Property Get GetCardData()
        GetCardData = SEPET
    End Property
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------
	Private Function TurkceKarakter2HTML(Txt)
		Txt = Txt & ""
		If IsNull(Txt) OR IsEmpty(Txt) OR Txt = "" OR Len(Txt) < 1 Then 
			Txt = ""
			Exit Function
		End If

		Txt = Replace(Txt, "ğ" ,"\u011F" ,1,-1,0)  
		Txt = Replace(Txt, "Ğ" ,"\u011E" ,1,-1,0)
		Txt = Replace(Txt, "ı" ,"\u0131" ,1,-1,0)
		Txt = Replace(Txt, "İ" ,"\u0130" ,1,-1,0)
		Txt = Replace(Txt, "ö" ,"\u00F6" ,1,-1,0)
		Txt = Replace(Txt, "Ö" ,"\u00D6" ,1,-1,0)
		Txt = Replace(Txt, "ü" ,"\u00FC" ,1,-1,0)
		Txt = Replace(Txt, "Ü" ,"\u00DC" ,1,-1,0)
		Txt = Replace(Txt, "ş" ,"\u015F" ,1,-1,0)
		Txt = Replace(Txt, "Ş" ,"\u015E" ,1,-1,0)
		Txt = Replace(Txt, "ç" ,"\u00E7" ,1,-1,0)
		Txt = Replace(Txt, "Ç" ,"\u00C7" ,1,-1,0)
		TurkceKarakter2HTML = Txt  
	End Function
	'---------------------------------------------------------------
	' 
	'---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
    Public Sub CreateInvoiceLog(OrderID, GUIDCode, Msg, FaturaURL, OrderNo)
    	Conn.Execute("INSERT INTO tbl_plugin_bizimhesap(ORDER_ID, GUID, MSG, FATURA_URL, ORDER_NO) VALUES('"& OrderID &"', '"& GUIDCode &"', '"& Msg &"', '"& FaturaURL &"', '"& OrderNo &"')")
    End Sub
	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------

	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
    Public Property Get InvoiceExist(OrderID)
    	InvoiceExist = False
    	
    	Set InvoiceControl = Conn.Execute("SELECT ID FROM tbl_plugin_bizimhesap WHERE ORDER_ID="& OrderID &"")
    	If Not InvoiceControl.Eof Then 
    		InvoiceExist = True
    	End If
    	InvoiceControl.Close : Set InvoiceControl = Nothing
    End Property
	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
End Class 


Class CartMaker
	Private tmp_productId, tmp_productName, tmp_note, tmp_barcode, tmp_taxRate
	Private tmp_quantity, tmp_unitPrice, tmp_grossPrice, tmp_discount, tmp_net, tmp_tax, tmp_total

    Private Sub Class_Initialize()
    End Sub

	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
    Public Property Get productId                  	: productId = tmp_productId 		: End Property
    Public Property Let productId(pVal)            	: tmp_productId = pVal 				: End Property
    Public Property Get productName                	: productName = tmp_productName 	: End Property
    Public Property Let productName(pVal) 			: tmp_productName = pVal 			: End Property
    Public Property Get note           				: note = tmp_note 					: End Property
    Public Property Let note(pVal)     				: tmp_note = pVal 					: End Property
    Public Property Get barcode           			: barcode = tmp_barcode 			: End Property
    Public Property Let barcode(pVal)     			: tmp_barcode = pVal 				: End Property
    Public Property Get taxRate            			: taxRate = tmp_taxRate 			: End Property
    Public Property Let taxRate(pVal)      			: tmp_taxRate = pVal 				: End Property
    Public Property Get quantity        			: quantity = tmp_quantity 			: End Property
    Public Property Let quantity(pVal)         		: tmp_quantity = pVal 				: End Property
    Public Property Get unitPrice    				: unitPrice = tmp_unitPrice 		: End Property
    Public Property Let unitPrice(pVal) 			: tmp_unitPrice = pVal 				: End Property
    Public Property Get grossPrice      			: grossPrice = tmp_grossPrice 		: End Property
    Public Property Let grossPrice(pVal) 			: tmp_grossPrice = pVal 			: End Property
    Public Property Get discount      				: discount = tmp_discount 			: End Property
    Public Property Let discount(pVal) 				: tmp_discount = pVal 				: End Property
    Public Property Get net      					: net = tmp_net 					: End Property
    Public Property Let net(pVal) 					: tmp_net = pVal 					: End Property
    Public Property Get tax      					: tax = tmp_tax 					: End Property
    Public Property Let tax(pVal) 					: tmp_tax = pVal 					: End Property
    Public Property Get total      					: total = tmp_total 				: End Property
    Public Property Let total(pVal) 				: tmp_total = pVal 					: End Property
	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------


	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
    Public Property Get Data
        Set Data = Server.CreateObject("Scripting.Dictionary")
	        If (NOT IsEmpty(productId))		Then Data.Add "productId", productId
	        If (NOT IsEmpty(productName))	Then Data.Add "productName", productName
	        If (NOT IsEmpty(note))			Then Data.Add "note", note
	        If (NOT IsEmpty(barcode))		Then Data.Add "barcode", barcode
	        If (NOT IsEmpty(taxRate))		Then Data.Add "taxRate", taxRate
	        If (NOT IsEmpty(quantity))		Then Data.Add "quantity", quantity
	        If (NOT IsEmpty(unitPrice))		Then Data.Add "unitPrice", unitPrice
	        If (NOT IsEmpty(grossPrice))	Then Data.Add "grossPrice", grossPrice
	        If (NOT IsEmpty(discount))		Then Data.Add "discount", discount
	        If (NOT IsEmpty(net))			Then Data.Add "net", net
	        If (NOT IsEmpty(tax))			Then Data.Add "tax", tax
	        If (NOT IsEmpty(total)) 		Then Data.Add "total", total
    End Property
	'---------------------------------------------------------------
	' 
    '---------------------------------------------------------------
End Class
%>