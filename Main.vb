
Module Test
    Public Sub Main()
        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12
        
        Dim companySubdomain As String = ""     ' Your company's subdomain here
        Dim userSecretKey As String = ""        ' Your user's api key here

        Dim cl As BambooAPIClient = New BambooAPIClient(companySubdomain)
        cl.setSecretKey(userSecretKey)

        Dim fields() As String = { _
            "id", _
            "ssn", _
            "employeeNumber", _
            "lastName", _
            "firstName", _
            "middleName", _
            "address1", _
            "address2", _
            "city", _
            "state", _
            "zipcode", _
            "bestEmail", _
            "dateOfBirth", _
            "exempt", _
            "jobTitle", _
            "homePhone", _
            "workPhonePlusExtension", _
            "mobilePhone", _
            "hireDate", _
            "terminationDate" _
        }

        Dim employeeId = 27

        Dim resp As BambooHTTPResponse

        System.Console.WriteLine("--- Custom Report ---")
        resp = cl.getEmployeesReport("csv", "Custom report", fields)
        System.Console.WriteLine(resp.getContentString())


        System.Console.WriteLine("--- Get time off requests ---")
        Dim filters As Hashtable = New Hashtable
        filters.Add("start", "2012-01-01")
        filters.Add("end", "2012-01-31")
        filters.Add("status", "approved")

        resp = cl.getTimeOffRequests(filters)
        System.Console.WriteLine(resp.getContentString())


        System.Console.WriteLine("--- Update employee ---")
        Dim values As Hashtable = New Hashtable
        values.Add("firstName", "VB")
        values.Add("lastName", "User 2")
        values.Add("status", "active")

        resp = cl.updateEmployee(41290, values)
        System.Console.WriteLine(resp.responseCode)
        System.Console.WriteLine(resp.getContentString())

    End Sub

End Module
