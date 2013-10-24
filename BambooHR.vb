'
' Copyright (c) 2011, 2012, 2013 Bamboo HR LLC
' All rights reserved.
'
' Redistribution and use in source and binary forms, with or without modification,
' are permitted provided that the following conditions are met:
'
' * Redistributions of source code must retain the above copyright notice, this
' list of conditions and the following disclaimer.
'
' * Redistributions in binary form must reproduce the above copyright notice,
' this list of conditions and the following disclaimer in the documentation
' and/or other materials provided with the distribution.
'
' * Neither the name of Bamboo HR nor the names of its contributors may be used
' to endorse or promote products derived from this software without specific
' prior written permission.
'
'
' THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
' ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
' WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
' DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
' FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
' DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
' SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
' CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
' OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
' OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
'

Imports System
Imports System.Collections.Specialized
Imports System.Net
Imports System.Text
Imports System.IO
Imports System.Xml
Imports System.Uri

Module BambooHR

    Public Class BambooHTTPRequest
        Public method As String = ""
        Public headers As NameValueCollection
        Public url As String = ""
        Public contents As String = ""

        Public Sub New()
            headers = New NameValueCollection()
        End Sub
    End Class

    Public Class BambooHTTPResponse
        Public responseCode As Integer
        Public headers As NameValueCollection
        Private content As HttpWebResponse

        Sub New(ByVal responseCode As Integer)
            Me.responseCode = responseCode
        End Sub

        Public Sub New(ByVal content As HttpWebResponse)
            Me.content = content
        End Sub


        Public Function getContentBytes() As Byte()

            If Not content Is Nothing Then
                Dim inStream As Stream = content.GetResponseStream()

                Dim buffer(32768) As Byte
                Using ms As MemoryStream = New MemoryStream()

                    While (True)
                        Dim read As Integer = inStream.Read(buffer, 0, buffer.Length)
                        If (read <= 0) Then
                            ms.Close()
                            content.Close()
                            Return ms.ToArray()
                        End If
                        ms.Write(buffer, 0, read)
                    End While

                End Using
                Return buffer

            Else
                Dim empty() As Byte = {}
                Return empty
            End If
        End Function


        Public Function getContentString() As String
            If (Not content Is Nothing) Then

                Dim contentReader As StreamReader = New StreamReader(content.GetResponseStream())
                Dim ret As String = contentReader.ReadToEnd()


                contentReader.Close()
                content.Close()
                Return ret
            Else
                Return ""
            End If
        End Function

    End Class


    Public Class BambooHTTPClient
        Private basicAuthUsername As String = ""
        Private basicAuthPassword As String
        Private verifyCert As Boolean = True

        Public Sub setBasicAuth(ByVal username As String, ByVal password As String)
            basicAuthUsername = username
            basicAuthPassword = password
        End Sub

        Public Sub setCertificateValidation(Optional ByVal verifyCert As Boolean = True)
            Me.verifyCert = verifyCert
        End Sub


        Public Function sendRequest(ByVal req As BambooHTTPRequest) As BambooHTTPResponse
            Dim request As HttpWebRequest = WebRequest.Create(req.url)
            request.KeepAlive = False
            request.Method = req.method

            Dim iCount As Integer = req.headers.Count
            Dim key As String
            Dim keyvalue As String

            Dim i As Integer
            For i = 0 To iCount - 1
                key = req.headers.Keys(i)
                keyvalue = req.headers(i)
                request.Headers.Add(key, keyvalue)
            Next

            Dim enc As System.Text.UTF8Encoding = New System.Text.UTF8Encoding()
            Dim bytes() As Byte = {}

            If (req.contents.Length > 0) Then
                bytes = enc.GetBytes(req.contents)
                request.ContentLength = bytes.Length
            End If

            request.AllowAutoRedirect = False

            If Not basicAuthUsername.Equals("") Then
                Dim authInfo As String = basicAuthUsername + ":" + basicAuthPassword
                authInfo = Convert.ToBase64String(Encoding.Default.GetBytes(authInfo))
                request.Headers("Authorization") = "Basic " + authInfo
            End If


            If req.contents.Length > 0 Then
                Dim outBound As Stream = request.GetRequestStream()
                outBound.Write(bytes, 0, bytes.Length)
            End If

            Dim resp As BambooHTTPResponse
            Try

                Dim webresponse As HttpWebResponse = request.GetResponse()
                resp = New BambooHTTPResponse(webresponse)
                resp.responseCode = webresponse.StatusCode
                resp.headers = webresponse.Headers
            Catch e As WebException

                If (e.Status = WebExceptionStatus.ProtocolError) Then
                    resp = New BambooHTTPResponse(DirectCast(e.Response, HttpWebResponse).StatusCode)
                Else
                    resp = New BambooHTTPResponse(0)
                End If
            End Try

            Return resp
        End Function


    End Class


    Public Class BambooAPIClient
        Private http As BambooHTTPClient
        Public baseUrl As String = "https://api.bamboohr.com/api/gateway.php"


        Public Sub New(ByVal companySubdomain As String, Optional ByVal client As BambooHTTPClient = Nothing)
            baseUrl += "/" + companySubdomain
            If Not client Is Nothing Then
                http = client
            Else
                http = New BambooHTTPClient()
            End If

        End Sub

        Public Sub setSecretKey(ByVal key As String)
            http.setBasicAuth(key, "x")
        End Sub



        Public Function getEmployee(ByVal employeeId As Integer, ByVal fields() As String) As BambooHTTPResponse
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "GET"
            request.url = String.Format("{0}/v1/employees/{1}/?fields={2}", _
                                        Me.baseUrl, employeeId, String.Join(",", fields))

            Return http.sendRequest(request)
        End Function

        Private Sub prepareFieldValues(ByRef post As XmlWriter, ByRef values As Hashtable)

            Dim denum As IDictionaryEnumerator = values.GetEnumerator()
            Dim dentry As DictionaryEntry

            Do While denum.MoveNext()
                dentry = DirectCast(denum.Current, DictionaryEntry)
                post.WriteStartElement("field")
                post.WriteAttributeString("id", dentry.Key)
                post.WriteString(dentry.Value)
                post.WriteEndElement()

            Loop            
        End Sub

        Public Function addEmployee(ByRef initialFieldValues As Hashtable)
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "POST"
            request.url = String.Format("{0}/v1/employees/", Me.baseUrl)

            Dim str As StringBuilder = New StringBuilder
            Dim post As XmlWriter = XmlWriter.Create(str)

            post.WriteProcessingInstruction("xml", "version=""1.0""")
            post.WriteStartElement("employee")
            Me.prepareFieldValues(post, initialFieldValues)
            post.WriteEndElement()

            post.WriteEndDocument()
            post.Flush()
            post.Close()
            request.contents = str.ToString()

            Return http.sendRequest(request)
        End Function



        Public Function updateEmployee(ByVal employeeId As Integer, ByRef fieldValues As Hashtable)
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()

            request.method = "POST"
            request.url = String.Format("{0}/v1/employees/{1}", Me.baseUrl, employeeId)

            Dim str As StringBuilder = New StringBuilder
            Dim post As XmlWriter = XmlWriter.Create(str)

            post.WriteProcessingInstruction("xml", "version=""1.0""")
            post.WriteStartElement("employee")
            post.WriteAttributeString("id", employeeId)
            Me.prepareFieldValues(post, fieldValues)
            post.WriteEndElement()
            post.WriteEndDocument()
            post.Flush()
            post.Close()
            request.contents = str.ToString()

            Return http.sendRequest(request)
        End Function

        Public Function getEmployeesReport(ByVal format As String, ByVal title As String, ByVal fields() As String) As BambooHTTPResponse

            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "POST"
            request.url = String.Format("{0}/v1/reports/custom?format={1}", Me.baseUrl, format)

            Dim str As StringBuilder = New StringBuilder()
            Dim post As XmlWriter = XmlWriter.Create(str)

            post.WriteProcessingInstruction("xml", "version=""1.0""")
            post.WriteStartElement("report")
            post.WriteElementString("title", title)
            post.WriteStartElement("fields")
            Dim i As Integer
            For i = 0 To fields.Length - 1
                post.WriteStartElement("field")
                post.WriteAttributeString("id", fields(i))
                post.WriteEndElement()
            Next
            post.WriteEndElement()
            post.WriteEndElement()
            post.Flush()
            post.Close()

            request.contents = str.ToString()
            Return http.sendRequest(request)
        End Function

        Public Function getEmployeeTable(ByVal employeeId As Integer, ByVal tableName As String) As BambooHTTPResponse
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "GET"
            request.url = String.Format("{0}/v1/employees/{1}/tables/{2}", Me.baseUrl, employeeId, tableName)
            Return http.sendRequest(request)
        End Function

        Private Function prepareTableRow(ByRef values As Hashtable)
            Dim str As StringBuilder = New StringBuilder
            Dim post As XmlWriter = XmlWriter.Create(str)

            post.WriteProcessingInstruction("xml", "version=""1.0""")
            post.WriteStartElement("row")
            prepareFieldValues(post, values)            
            post.WriteEndElement()
            post.Flush()
            Return str.ToString()
        End Function


        Public Function addTableRow(ByVal employeeId As Integer, ByVal tableName As String, ByRef values As Hashtable) As BambooHTTPResponse
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "POST"
            request.url = String.Format("{0}/v1/employees/{1}/tables/{2}", Me.baseUrl, employeeId, tableName)
            request.contents = prepareTableRow(values)
            Return http.sendRequest(request)
        End Function

        Public Function updateTableRow(ByVal employeeId As Integer, ByVal tableName As String, ByVal rowId As Integer, ByRef values As Hashtable)
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()
            request.method = "POST"
            request.url = String.Format("{0}/v1/employees/{1}/tables/{2}/{3}", Me.baseUrl, employeeId, tableName, rowId)
            request.contents = prepareTableRow(values)
            Return http.sendRequest(request)
        End Function


        Public Function getTimeOffRequests(ByRef filters As Hashtable)
            Dim request As BambooHTTPRequest = New BambooHTTPRequest()            
            request.method = "GET"
            request.url = Me.baseUrl + "/v1/time_off/requests?"

            Dim denum As IDictionaryEnumerator = filters.GetEnumerator()
            Dim dentry As DictionaryEntry

            Do While denum.MoveNext()
                dentry = DirectCast(denum.Current, DictionaryEntry)
                request.url = request.url + Uri.EscapeUriString(dentry.Key) + "=" + Uri.EscapeUriString(dentry.Value) + "&"
            Loop
            Console.Out.WriteLine(request.url)
            Return http.sendRequest(request)
        End Function


    End Class


    

End Module
