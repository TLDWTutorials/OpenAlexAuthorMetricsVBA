# 📊 Automate Author Metrics Retrieval with OpenAlex API & VBA

## 🚀 Overview
This VBA script allows you to **automatically retrieve author metrics** (such as **H-index, i10-index, citation count, and work count**) from the **OpenAlex API** and store them in an **Excel sheet**.  

It supports **both direct Author ID lookups** and **name-based searches** when an ID is unavailable.

## 📌 Features
✅ **Fetch H-index, i10-index, total citations, and work count**  
✅ **Supports both Author ID and Name-based searches**  
✅ **Automatically formats & saves data in Excel**  
✅ **Includes error handling for missing or incorrect authors**  
✅ **Implements rate limiting to avoid API bans**  

## 🛠️ Requirements
- **Microsoft Excel** with **VBA enabled**
- **Internet connection** (to access the OpenAlex API)
- **Basic knowledge of VBA (optional)**

## 📄 How It Works
1. Enter author names in **Column A** of an Excel sheet.
2. (Optional) If you have OpenAlex Author IDs, enter them in **Column B**.
3. Run the VBA script.
4. The script queries the OpenAlex API and fills in **Columns C to H** with:
   - **H-index**
   - **i10-index**
   - **Citation count**
   - **Total work count**
   - **OpenAlex profile URL**
   - **Number of search results (if searching by name)**

## 📜 VBA Code
```vba
Sub GetAuthorMetrics_OpenAlex()
    Dim http As Object, url As String, authorName As String, authorID As String
    Dim lastRow As Integer, ws As Worksheet, i As Integer, responseText As String
    Dim hIndex As String, i10Index As String, citationCount As String, workCount As String
    Dim authorURL As String, searchResultsCount As String, startPos As Integer, endPos As Integer

    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets(1)

    ' Add headers if not already present
    ws.Cells(1, 3).Value = "H-index"
    ws.Cells(1, 4).Value = "i10-index"
    ws.Cells(1, 5).Value = "Citation Count"
    ws.Cells(1, 6).Value = "Work Count"
    ws.Cells(1, 7).Value = "OpenAlex URL"
    ws.Cells(1, 8).Value = "Search Results Count"

    ' Find the last row with data in column A
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Create HTTP Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    ' Loop through each author
    For i = 2 To lastRow
        authorName = Trim(ws.Cells(i, 1).Value)
        authorID = Trim(ws.Cells(i, 2).Value)

        ' Construct OpenAlex API URL
        If authorID <> "" Then
            url = "https://api.openalex.org/authors/" & authorID
        Else
            url = "https://api.openalex.org/authors?search=" & Replace(authorName, " ", "%20") & "&sort=works_count:desc"
        End If

        ' Make API request
        http.Open "GET", url, False
        http.Send

        ' Check response
        If http.Status = 200 Then
            responseText = http.responseText
            hIndex = ExtractJSONValue(responseText, """h_index"":")
            i10Index = ExtractJSONValue(responseText, """i10_index"":")
            citationCount = ExtractJSONValue(responseText, """cited_by_count"":")
            workCount = ExtractJSONValue(responseText, """works_count"":")
            
            ' Extract Author URL
            If authorID <> "" Then
                authorURL = "https://openalex.org/" & authorID
            Else
                startPos = InStr(responseText, """results"":[{""id"":""https://openalex.org/")
                If startPos > 0 Then
                    startPos = startPos + 23
                    endPos = InStr(startPos, responseText, """")
                    authorURL = "https://openalex.org/" & Mid(responseText, startPos, endPos - startPos)
                    authorURL = Replace(authorURL, "://openalex.org/://openalex.org/", "://openalex.org/")
                Else
                    authorURL = "N/A"
                End If
            End If

            ' Extract search results count
            If authorID = "" Then
                startPos = InStr(responseText, """count"":")
                If startPos > 0 Then
                    startPos = startPos + 8
                    endPos = InStr(startPos, responseText, ",")
                    searchResultsCount = Trim(Mid(responseText, startPos, endPos - startPos))
                Else
                    searchResultsCount = "0"
                End If
            Else
                searchResultsCount = "1"
            End If
        Else
            hIndex = "Error"
            i10Index = "Error"
            citationCount = "Error"
            workCount = "Error"
            authorURL = "Error"
            searchResultsCount = "Error"
        End If

        ' Write results to Excel
        ws.Cells(i, 3).Value = hIndex
        ws.Cells(i, 4).Value = i10Index
        ws.Cells(i, 5).Value = citationCount
        ws.Cells(i, 6).Value = workCount
        ws.Cells(i, 7).Value = authorURL
        ws.Cells(i, 8).Value = searchResultsCount

        ' Delay to avoid API rate limits
        Application.Wait (Now + TimeValue("00:00:01"))
    Next i

    MsgBox "Author metrics retrieval completed!", vbInformation, "Done"
End Sub

' Function to extract JSON values from API response
Function ExtractJSONValue(jsonString As String, key As String) As String
    Dim startPos As Integer, endPos As Integer, value As String

    startPos = InStr(jsonString, key)
    If startPos > 0 Then
        startPos = startPos + Len(key)
        endPos = InStr(startPos, jsonString, ",")
        If endPos = 0 Then endPos = InStr(startPos, jsonString, "}")
        value = Trim(Mid(jsonString, startPos, endPos - startPos))
        value = Replace(value, ":", "")
        value = Replace(value, """", "")
    Else
        value = "N/A"
    End If

    ExtractJSONValue = value
End Function
```

## 🛠 Troubleshooting
| **Issue** | **Solution** |
|-----------|-------------|
| **No results found for an author?** | Ensure their name is correctly spelled, or use their OpenAlex ID. |
| **Error in API response?** | OpenAlex might be down or rate-limiting requests. Try again later. |
| **VBA script not running?** | Enable macros in Excel (File → Options → Trust Center → Macro Settings). |
| **Slow performance?** | The script includes a **1-second delay** to avoid API rate limits. |

---

## 📢 Support & Contributions
💡 **Have ideas for improvements? Found a bug?** Open an issue or submit a pull request!  
⭐ **Enjoyed this project?** Give it a star on GitHub!  

### 📚 Learn More
- OpenAlex API Docs: [https://docs.openalex.org/](https://docs.openalex.org/)
- VBA for Beginners: [Microsoft VBA Guide](https://docs.microsoft.com/en-us/office/vba/api/overview/)

## License
This project is licensed under the MIT License.
