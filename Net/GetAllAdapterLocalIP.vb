Private Function GetAllAdapterLocalIP()
    Dim adapterInfoDT As New DataTable
    adapterInfoDT.Columns.Add("name", GetType(String))
    adapterInfoDT.Columns.Add("ip", GetType(String))

    Dim adapters As NetworkInterface() = NetworkInterface.GetAllNetworkInterfaces
    If adapters.Length < 0 Or adapters Is Nothing Then
        MsgBox("No network interfaces found.")
    End If

    For Each ni As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()

        If ni.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Or ni.NetworkInterfaceType = NetworkInterfaceType.Ethernet Then
            
            For Each ip As UnicastIPAddressInformation In ni.GetIPProperties().UnicastAddresses

                If ip.Address.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then
                    adapterInfoDT.Rows.Add(ni.Name.TrimEnd(), ip.Address.ToString)
                    
                End If
            Next
        End If
    Next
    Return adapterInfoDT
End Function