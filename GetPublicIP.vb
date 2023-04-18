Dim webClient As New System.Net.WebClient()
Dim utf8 As New System.Text.UTF8Encoding()
' Get public ip from URL by Https
' @params ip String in/out
ip = utf8.GetString(webClient.DownloadData("https://api.ipify.org"))
