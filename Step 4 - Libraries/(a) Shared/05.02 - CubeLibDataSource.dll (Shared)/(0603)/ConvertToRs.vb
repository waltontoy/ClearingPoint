Imports System.Data
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Xsl
Imports System.IO
Imports System.Text

'**************************************************************************
'   Class Name  : ConvertToRs
'   Description : This class converts a DataSet to a ADODB Recordset.
'**************************************************************************
Public Class ConvertToRs

    '**************************************************************************
    '   Method Name : GetADORS
    '   Description : Takes a DataSet and converts into a Recordset. The converted 
    '                 ADODB recordset is saved as an XML file. The data is saved 
    '                 to the file path passed as parameter.
    '   Output      : The output of this method is long. Returns 1 if successfull. 
    '                 If not throws an exception. 
    '   Input parameters:
    '               1. DataSet object
    '               2. Database Name
    '               3. Output file - where the converted should be written.
    '**************************************************************************
    Public Function GetADORS(ByVal DS As DataSet, ByVal dbName As String) As String

        Try
            'Create a MemoryStream to contain the XML
            Dim mStream As New MemoryStream
            'Create an XmlWriter object, to write 
            'the formatted XML to the MemoryStream
            Dim xWriter As New XmlTextWriter(mStream, Nothing)

            'Additional formatting for XML
            xWriter.Indentation = 8
            xWriter.Formatting = Formatting.Indented
            'call this Sub to write the ADONamespaces
            WriteADONamespaces(xWriter)
            'call this Sub to write the ADO Recordset Schema
            WriteSchemaElement(DS, dbName, xWriter)
            'Call this sub to transform 
            'the data portion of the Dataset
            TransformData(DS, xWriter)
            'Flush all input to XmlWriter
            xWriter.Flush()

            'Prepare the return value
            mStream.Position = 0
            Dim Buffer As Array
            Buffer = Array.CreateInstance(GetType(Byte), mStream.Length)
            mStream.Read(Buffer, 0, mStream.Length)
            'mStream.Read(Buffer, 0, 350)
            Dim TextConverter As New UTF8Encoding
            'Debug.Print(TextConverter.GetString(Buffer))
            Return TextConverter.GetString(Buffer)

        Catch ex As Exception
            'Returns error message to the calling function.
            Err.Raise(100, ex.Source, ex.ToString)
            Return ""
        End Try

    End Function


    Private Sub WriteADONamespaces(ByRef xWriter As XmlTextWriter)
        'Uncomment the following line to change 
        'the encoding if special characters are required
        'xWriter.WriteProcessingInstruction("xml", "version='1.0' encoding='ISO-8859-1'")

        'Add XML start element
        xWriter.WriteStartElement("xml", "xml", "")

        'Append the ADO Recordset namespaces
        xWriter.WriteAttributeString("xmlns", "s", Nothing, _
                "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
        xWriter.WriteAttributeString("xmlns", "dt", Nothing, _
                "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882")
        xWriter.WriteAttributeString("xmlns", "rs", Nothing, _
                "urn:schemas-microsoft-com:rowset")
        xWriter.WriteAttributeString("xmlns", "z", _
                Nothing, "#RowsetSchema")
        xWriter.Flush()
    End Sub


    Private Sub WriteSchemaElement(ByVal DS As DataSet, _
        ByVal dbName As String, ByRef xWriter As  _
        XmlTextWriter)

        'write element Schema
        xWriter.WriteStartElement("s", "Schema", _
                "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
        xWriter.WriteAttributeString("id", "RowsetSchema")

        'write element ElementType
        xWriter.WriteStartElement("s", "ElementType", _
                "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")

        'write the attributes for ElementType
        xWriter.WriteAttributeString("name", "", "row")
        xWriter.WriteAttributeString("content", "", "eltOnly")
        xWriter.WriteAttributeString("rs", "updatable", _
                "urn:schemas-microsoft-com:rowset", "true")

        WriteSchema(DS, dbName, xWriter)
        'write the end element for ElementType
        xWriter.WriteFullEndElement()

        'write the end element for Schema
        xWriter.WriteFullEndElement()
        xWriter.Flush()
    End Sub


    Private Sub WriteSchema(ByVal DS As DataSet, ByVal dbName _
            As String, ByRef xWriter As XmlTextWriter)

        Dim i As Int32 = 1
        Dim DC As DataColumn

        For Each DC In DS.Tables(0).Columns

            DC.ColumnMapping = MappingType.Attribute

            xWriter.WriteStartElement("s", "AttributeType", _
                    "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            'write all the attributes
            xWriter.WriteAttributeString("name", "", DC.ToString)
            xWriter.WriteAttributeString("rs", "number", _
                     "urn:schemas-microsoft-com:rowset", i.ToString)
            xWriter.WriteAttributeString("rs", "baseCatalog", _
                     "urn:schemas-microsoft-com:rowset", dbName)
            xWriter.WriteAttributeString("rs", "baseTable", _
                     "urn:schemas-microsoft-com:rowset", _
                     DC.Table.TableName.ToString)
            xWriter.WriteAttributeString("rs", "keycolumn", _
                     "urn:schemas-microsoft-com:rowset", _
                     DC.Unique.ToString)
            xWriter.WriteAttributeString("rs", "autoincrement", _
                     "urn:schemas-microsoft-com:rowset", _
                     DC.AutoIncrement.ToString)
            'write child element
            xWriter.WriteStartElement("s", "datatype", _
                    "uuid:BDC6E3F0-6DA3-11d1-A2A3-00AA00C14882")
            'write attributes
            xWriter.WriteAttributeString("dt", "type", _
                     "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882", _
                     GetDatatype(DC.DataType.ToString))
            xWriter.WriteAttributeString("dt", "maxlength", _
                     "uuid:C2F41010-65B3-11d1-A29F-00AA00C14882", _
                     DC.MaxLength.ToString)
            xWriter.WriteAttributeString("rs", "maybenull", _
                    "urn:schemas-microsoft-com:rowset", _
                     DC.AllowDBNull.ToString)
            'write end element for datatype
            xWriter.WriteEndElement()
            'end element for AttributeType
            xWriter.WriteEndElement()
            xWriter.Flush()
            i = i + 1
        Next
        DC = Nothing

    End Sub


    Private Function GetDatatype(ByVal DType As String) As String
        Select Case (DType)
            Case "System.Int32", "System.Int16", "System.Integer"
                Return "int"
            Case "System.DateTime"
                Return "dateTime.iso8601tz"
            Case "System.String"
                Return "string"
            Case "System.Byte[]"
                Return "bin.hex"
            Case "System.Boolean"
                Return "boolean"
            Case "System.Guid"
                Return "guid"
            Case Else
                Return "string"
        End Select
    End Function


    Private Sub TransformData(ByVal DS As DataSet, _
            ByRef xWriter As XmlTextWriter)

        'Loop through DataSet and add data to XML
        xWriter.WriteStartElement("", "rs:data", "")
        Dim i As Long
        Dim j As Integer
        'For each row...
        For i = 0 To DS.Tables(0).Rows.Count - 1
            'Write the start element for the row
            xWriter.WriteStartElement("", "z:row", "")
            'For each field in the row...
            For j = 0 To DS.Tables(0).Columns.Count - 1
                'Write the attribute that describes 
                'this field and it's value
                If DS.Tables(0).Columns(j).DataType.ToString = "System.Byte[]" Then
                    'Binary data must be properly encoded (bin.hex)
                    If Not IsDBNull(DS.Tables(0).Rows(i).Item(
                           DS.Tables(0).Columns(j).ColumnName)) Then
                        xWriter.WriteAttributeString(DS.Tables(0).
                           Columns(j).ColumnName, _
                           DataToBinHex(DS.Tables(0).Rows(i).Item(
                           DS.Tables(0).Columns(j).ColumnName)))
                    End If
                Else
                    If Not IsDBNull(DS.Tables(0).Rows(i).Item(
                           DS.Tables(0).Columns(j).ColumnName)) Then
                        xWriter.WriteAttributeString(
                                DS.Tables(0).Columns(j).ColumnName, _
               CType( _
                  DS.Tables(0).Rows(i).Item(DS.Tables(0).
                               Columns(j).ColumnName), String))
                    End If
                End If
            Next
            'End the row element
            xWriter.WriteEndElement()
        Next
        'Write the end element for rs:data
        xWriter.WriteEndElement()
        'Write the end element for xml
        xWriter.WriteEndElement()
        xWriter.Flush()

    End Sub

    Private Function DataToBinHex(ByVal thisData As Byte()) As String
        Dim sb As New StringBuilder
        Dim i As Integer = 0
        For i = 0 To thisData.Length - 1
            'First nibble of byte (4 most significant bits)
            sb.Append(Hex((thisData(i) And &HF0) / 2 ^ 4))
            'Second nibble of byte (4 least significant bits)
            sb.Append(Hex(thisData(i) And &HF))
        Next
        Return sb.ToString
    End Function

End Class