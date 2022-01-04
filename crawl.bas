Attribute VB_Name = "Module1"
Sub crawl()
ProgramStart:
    Dim URL As String, processing As Worksheet, data As Worksheet
    Set processing = ThisWorkbook.Worksheets("processing")
    Set data = ThisWorkbook.Worksheets("crawl_data")
    URL = processing.Cells(2, "B").Value
    
    Dim ie As Object
    Set ie = CreateObject("InternetExplorer.Application")
    
    With ie
        .navigate URL
        ieBusy ie
        .Visible = False
        
        Dim pagebar As Object, pages As Integer, ind As Integer, page As Integer, p As Integer, pg As Integer
        Set pagebar = .document.getElementsByClassName(processing.Cells(3, "B").Value)(processing.Cells(3, "C").Value)
        pages = pagebar.childElementCount - 1
        processing.Cells(4, "B").Value = pages
        
        page = processing.Cells(2, "E").Value
        p = processing.Cells(2, "F").Value
        
        ind = (page - 1) * 12 + p + 1
        For i = page To pages
            .navigate URL & "?p=" & i
            ieBusy ie
            
            processing.Cells(2, "E").Value = i
            Dim cont As Object, products As Integer
            Set cont = .document.getElementsByClassName(processing.Cells(5, "B").Value)(processing.Cells(5, "C").Value)
            products = cont.childElementCount
            If i = 1 Then
            processing.Cells(6, "B").Value = products
            Else
            processing.Cells(6, "B").Value = processing.Cells(6, "B").Value & "," & products
            End If
            
            If i = page Then
            pg = p
            Else
            pg = 1
            End If
            
            For j = pg To products
                On Error GoTo ErrorHandler2
                processing.Cells(2, "F").Value = j
                Dim product As Object, href As String
                Set cont = .document.getElementsByClassName(processing.Cells(5, "B").Value)(processing.Cells(5, "C").Value)
                Set product = cont.Children(j - 1)
                href = product.querySelector("a")
                
                .navigate href & "?p=" & i
                ieBusy ie
                Dim handle As String, h() As String, title As String, body As String, vendor As String, pType As String, tags As String, published As String, optionName As String, optionValue As String, SKU As String, price As String, shipping As String, taxable As String, weight As String, stats As String, additional As String, benefits As String, FAQs As String, treatmentType As String, treatmentProblem As String, skin As String, size As String, brand As String, usage As String, ingredients As String, video As String, photos As String
                h = Split(href, "/")
                handle = Split(h(UBound(h)), ".")(0)
                title = .document.getElementsByClassName(processing.Cells(7, "B").Value)(processing.Cells(7, "C").Value).innerHTML
                body = .document.getElementsByClassName(processing.Cells(8, "B").Value)(processing.Cells(8, "C").Value).innerHTML
                vendor = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='Brand']").innerHTML
                pType = "Product"
                published = "TRUE"
                'optionName
                'optionValue
                Dim SKUh As Object, SKUhref() As String
                Set SKUh = .document.getElementsByClassName(processing.Cells(11, "B").Value)(processing.Cells(11, "C").Value)
                SKUhref = Split(SKUh.href, "/")
                On Error GoTo ErrorHandler1
                SKU = Split(.document.getElementsByClassName(processing.Cells(10, "B").Value)(processing.Cells(10, "C").Value).Children(0).Children(0).Children(0).ID, "-")(2)
                price = .document.getElementsByClassName(processing.Cells(10, "B").Value)(processing.Cells(10, "C").Value).Children(0).Children(0).Children(0).Children(0).innerHTML
                On Error GoTo ErrorHandler2
                shipping = "TRUE"
                taxable = "TRUE"
                weight = "lb"
                If price = "" Then
                stats = "inactive"
                Else
                stats = "active"
                End If
                additional = .document.getElementById(processing.Cells(12, "B").Value).innerHTML
                benefits = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='Benefits']").innerHTML
                FAQs = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='FAQs']").innerHTML
                'treatmentType
                'treatmentProblem
                skin = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='Skin Type']").innerHTML
                size = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='Size']").innerHTML
                brand = .document.getElementsByClassName(processing.Cells(9, "B").Value)(processing.Cells(9, "C").Value).querySelector("[data-th='Brand']").innerHTML
                usage = .document.getElementById(processing.Cells(13, "B").Value).innerHTML
                ingredients = .document.getElementById(processing.Cells(14, "B").Value).innerHTML
                'video
                photos = .document.getElementById(processing.Cells(15, "B").Value & "-" & SKU).href
                                
                data.Cells(ind, "A").Value = handle
                data.Cells(ind, "B").Value = title
                data.Cells(ind, "C").Value = body
                data.Cells(ind, "C").WrapText = False
                data.Cells(ind, "D").Value = vendor
                data.Cells(ind, "E").Value = pType
                data.Cells(ind, "G").Value = published
                data.Cells(ind, "J").Value = SKU
                data.Cells(ind, "K").Value = price
                data.Cells(ind, "L").Value = shipping
                data.Cells(ind, "M").Value = taxable
                data.Cells(ind, "N").Value = weight
                data.Cells(ind, "O").Value = stats
                data.Cells(ind, "R").Value = additional
                data.Cells(ind, "S").Value = benefits
                data.Cells(ind, "R").WrapText = False
                data.Cells(ind, "S").WrapText = False
                data.Cells(ind, "T").Value = FAQs
                data.Cells(ind, "W").Value = skin
                data.Cells(ind, "X").Value = size
                data.Cells(ind, "Y").Value = brand
                data.Cells(ind, "Z").Value = usage
                data.Cells(ind, "AA").Value = ingredients
                data.Cells(ind, "Z").WrapText = False
                data.Cells(ind, "AA").WrapText = False
                data.Cells(ind, "AC").Value = photos
                
                ind = ind + 1
                'MsgBox SKU
                .navigate URL & "?p=" & i
                ieBusy ie
            Next j
        Next i
    End With
    
Exit Sub
ErrorHandler1:
    SKU = SKUhref(UBound(SKUhref) - 3)
    price = ""
    Resume Next
ErrorHandler2:
    MsgBox "2"
    ie.Quit
    GoTo ProgramStart
End Sub

Sub ieBusy(ie As Object)
    Do While ie.Busy Or ie.readyState < 4
        DoEvents
    Loop
End Sub

