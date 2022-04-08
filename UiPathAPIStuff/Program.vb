Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Configuration

Module Module1

    Sub Main()

        Call GetQueueItem()

    End Sub


    Sub AddQueueItem()


        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create("https://cloud.uipath.com/TomsTestOrch/DefaultTenant/odata/Queues/UiPathODataSvc.AddQueueItem")
        myReq.Method = "POST"
        myReq.ContentType = "application/json"
        myReq.Timeout = 100000

        myReq.Headers.Add("Authorization", "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6IlJUTkVOMEl5T1RWQk1UZEVRVEEzUlRZNE16UkJPVU00UVRRM016TXlSalUzUmpnMk4wSTBPQSJ9.eyJodHRwczovL3VpcGF0aC9lbWFpbCI6InRqb25lc0Byb2JpcXVpdHkuY29tIiwiaHR0cHM6Ly91aXBhdGgvZW1haWxfdmVyaWZpZWQiOnRydWUsImlzcyI6Imh0dHBzOi8vYWNjb3VudC51aXBhdGguY29tLyIsInN1YiI6Im9hdXRoMnxVaVBhdGgtQUFEVjJ8ZDMyNjE4ZDQtMmIwMC00MWJhLThjMTYtYTc5NTRlY2Q0Mjg0IiwiYXVkIjpbImh0dHBzOi8vb3JjaGVzdHJhdG9yLmNsb3VkLnVpcGF0aC5jb20iLCJodHRwczovL3VpcGF0aC5ldS5hdXRoMC5jb20vdXNlcmluZm8iXSwiaWF0IjoxNjQ5MjM4NjU4LCJleHAiOjE2NDkzMjUwNTgsImF6cCI6IjhERXYxQU1OWGN6VzN5NFUxNUxMM2pZZjYyaks5M241Iiwic2NvcGUiOiJvcGVuaWQgcHJvZmlsZSBlbWFpbCBvZmZsaW5lX2FjY2VzcyJ9.ZbPgMCd5gdP5l0VKXTztxkgs3m60MRdg2cR7zy4Zr9XWqdVzB6A_Y9EMg3G4wCeS5vq4O8Ad4Q5s-aTQUUh3pjpfwW_Zh9rxVRzIPX-DOXIQcpc9c1p3-ki7iscwWHOaFzTG57gIonLd1w-P7NQzMYoNIRMUGC2lAKRMbTUh0spZ8HscNQrnQhdPqXR6lOh4QUzcU20ExII0iHkYFFP11ghBnBtiab1AQucCKRZC_6IjnIHlw4g1jOwAsmJm44YipRFWqY9bibZeHtqCbigunoiydatW9ayP5BWHf0zv3USHd1r8d8H2GfYWb2OEdgsuGvlFwpXrLkr5Ihl-iAEXQQ")
        myReq.Headers.Add("X-UIPATH-TenantName", "DefaultTenant")
        myReq.Headers.Add("X-UIPATH-FolderPath", "Toms Folder")
        myReq.Headers.Add("Cookie", "UiPathBrowserId=f75f2b58-4670-4525-bbc3-901dff64d904")

        Dim PostString As String '= JsonConvert.SerializeObject(NewData)

        PostString = "{
	""itemData"": {
		""Priority"": ""High"",
		""DeferDate"": ""2023-03-21T13:42:27.654Z"",
		""DueDate"": ""2022-03-25T13:42:27.654Z"",
		""Name"": ""CAD_ExportQueue"",
		""SpecificContent"": {
			""Email"": ""obrian@uipath.com"", 
			""Name"": ""O'Brian""
		}
	}
}"

        Dim byteArray As Byte() = Encoding.UTF8.GetBytes(PostString)
        myReq.ContentLength = byteArray.Length
        Dim dataStream As Stream = myReq.GetRequestStream()
        dataStream.Write(byteArray, 0, byteArray.Length)
        dataStream.Close() 'sends request

        Dim myResp As HttpWebResponse = myReq.GetResponse()


    End Sub

    Sub GetQueueItem()


        Dim myReq As Net.HttpWebRequest = HttpWebRequest.Create("https://cloud.uipath.com/TomsTestOrch/DefaultTenant/odata/QueueItems?$filter=Status eq 'New'")
        myReq.Method = "GET"
        myReq.ContentType = "application/json"
        myReq.Timeout = 100000

        myReq.Headers.Add("Authorization", "Bearer eyJhbGciOiJSUzI1NiIsInR5cCI6IkpXVCIsImtpZCI6IlJUTkVOMEl5T1RWQk1UZEVRVEEzUlRZNE16UkJPVU00UVRRM016TXlSalUzUmpnMk4wSTBPQSJ9.eyJodHRwczovL3VpcGF0aC9lbWFpbCI6InRqb25lc0Byb2JpcXVpdHkuY29tIiwiaHR0cHM6Ly91aXBhdGgvZW1haWxfdmVyaWZpZWQiOnRydWUsImlzcyI6Imh0dHBzOi8vYWNjb3VudC51aXBhdGguY29tLyIsInN1YiI6Im9hdXRoMnxVaVBhdGgtQUFEVjJ8ZDMyNjE4ZDQtMmIwMC00MWJhLThjMTYtYTc5NTRlY2Q0Mjg0IiwiYXVkIjpbImh0dHBzOi8vb3JjaGVzdHJhdG9yLmNsb3VkLnVpcGF0aC5jb20iLCJodHRwczovL3VpcGF0aC5ldS5hdXRoMC5jb20vdXNlcmluZm8iXSwiaWF0IjoxNjQ5NDEwNjM5LCJleHAiOjE2NDk0OTcwMzksImF6cCI6IjhERXYxQU1OWGN6VzN5NFUxNUxMM2pZZjYyaks5M241Iiwic2NvcGUiOiJvcGVuaWQgcHJvZmlsZSBlbWFpbCBvZmZsaW5lX2FjY2VzcyJ9.Nitzwx_KRJnWH2fcKUb-FlMbByLaMBlnUGCXrqmVQUgtz89pUZuQye2OFb032QnxRXHjeOavnsy6fYtAevlAZj0e_u95255DhNRTy0zkdVTf3IJ6ATiNJNSNh95Is0xEEQIomaDwMTQgZh7MSB4RmmfikH69fATQtMtKPK2kWnIOyEQKU9C2GCt_M3Yb14uCHtkV0ZYskVG4VohBjQggL_a7NYaGjBjPXvD1WidMHg4-h_WOC8AgsQ6EZxBGiWeNyOyUO8pK9Rv4ttiIPobxh8IfOVEOQfM6rmuHGwuHDdSqVwApADQc1yxHvS65X5_W5d17fd3DkuFGRtji7cdacw")

        Dim response As HttpWebResponse = CType(myReq.GetResponse(), HttpWebResponse)

        Console.WriteLine("Content length is {0}", response.ContentLength)
        Console.WriteLine("Content type is {0}", response.ContentType)

        Dim receiveStream As Stream = response.GetResponseStream()

        ' Pipes the stream to a higher level stream reader with the required encoding format. 
        Dim readStream As New StreamReader(receiveStream, Encoding.UTF8)

        Console.WriteLine("Response stream received.")
        Console.WriteLine(readStream.ReadToEnd())
        response.Close()
        readStream.Close()


    End Sub



End Module
