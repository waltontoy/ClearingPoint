Attribute VB_Name = "MADOXCreateRelationships"
' #VBIDEUtils#************************************************************
' * Author           :  Larry Rebich
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           : waty.thierry@vbdiamond.com
' * Date             : 12/10/2003
' * Purpose          :
' * Project Name     : DBUpdateADO
' * Module Name      : modCreateRelationshipUsingADOX
' **********************************************************************
' * Comments         :
' *
' *
' * Example          :
' *
' * History          : Updated by Waty Thierry
' * 2003/01/31 Copyright © 2003, Larry Rebich, using the DELL7500
' * 2003/01/31 larry@buygold.net, www.buygold.net, 760-771-4730
' * 2003/02/28 Modified by Larry Rebich to support an array of keys and cascade rules
' * 2003/05/13 Used in projOMT
' *
' * See Also         :
' *
' *
' **********************************************************************

Option Explicit
DefLng A-Z

'From:
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/odeopg/html/deovrcreatingrelationshipsintegrityconstraints.asp
'
'Creating Relationships and Integrity Constraints
'
'In a relational database, a relationship is created between a foreign key in one table and typically a primary key _
(or some other field that contains unique values) in another table. For simplicity's sake, we'll assume that the foreign _
key and primary key are both single fields, but it is possible that either or both can be made up of more than one field. _
   To get a handle on the tables that contain the foreign and primary keys, we'll use the term foreign table for the table _
that contains the foreign key, and the term related table for the table that contains the primary key.
'
'A foreign table is most typically on the "many" side of a relationship. For example, in the Northwind sample database _
there is a one-to-many relationship between the Categories table and the Products table, _
   so the Categories table is the related table and the Products table is the foreign table, _
   and the relationship is established between the primary key field in the Categories table, CategoryID, _
   and the foreign key field in the Products table, which is also named CategoryID.
'
'To create a relationship by using ADOX, you use a Key object to create an object that defines the relationship _
by specifying the related table and key fields. You then open a Table object on the foreign table and _
   add the new Key object to the Keys collection of the table. _
   The following code sample shows how to create a relationship by using ADOX.

'For example, to use this procedure to create the relationship described above between the Categories and Products _
tables in the Northwind database, you can use a line of code like this:
'
'CreateRelationship _
'    "c:\Program Files\Microsoft Office\Office\Samples\Northwind.mdb", _
'    "CategoriesProducts", "Products",  "CategoryID", "Categories", "CategoryID"
'
'The CreateRelationship procedure can be found in the CreateDatabase module in the DataAccess.mdb sample file, _
which is available in the ODETools\V9\Samples\OPG\Samples\CH14 subfolder on the Office 2000 Developer CD-ROM.
'
'The Key object also supports two additional properties that are used to define whether related records will be _
automatically updated or deleted if the value in the primary key in the related table is changed. _
   These features of the Jet database engine are called cascading updates and cascading deletions. _
   By default, cascading updates and deletions are not active when you create a new relationship. _
   To turn on cascading updates for the relationship, set the UpdateRule property for the Key object to adRICascade; _
   turn on cascading deletions for the relationship, set the DeleteRule property for the Key object to adRICascade.
'
'Note   When you use OLE DB, there is no way to create a relationship that is not enforced, therefore there is no _
equivalent in ADOX to the DAO dbRelationDontEnforce setting of the DAO Attributes property of a Relation object. _
   Also, ADOX and the Microsoft Jet 4.0 OLE DB Provider don't provide a way to specify the default join type that will _
be used in the Access query Design view window, as can be done by using the dbRelationRight and dbRelationLeft _
   settings of the Attributes property.

Public Function CreateRelationshipUsingADOX(catDB As ADOX.Catalog, sRelationshipName As String, _
   sForeignTable As String, arysForeignTableKeys() As String, _
   sRelatedTable As String, arysRelatedTableKeys() As String, _
   lUpdateRule As ADOX.RuleEnum, lDeleteRule As ADOX.RuleEnum) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           : Waty Thierry
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modCreateRelationshipUsingADOX
   ' * Procedure Name   : CreateRelationshipUsingADOX
   ' * Parameters       :
   ' *                    catDB As ADOX.Catalog
   ' *                    sRelationshipName As String
   ' *                    sForeignTable As String
   ' *                    arysForeignTableKeys() As String
   ' *                    sRelatedTable As String
   ' *                    arysRelatedTableKeys() As String
   ' *                    lUpdateRule As ADOX.RuleEnum
   ' *                    lDeleteRule As ADOX.RuleEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/03/01 Function modified by Larry Rebich while in La Quinta, CA.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************

   Dim tbl              As ADOX.Table
   Dim kee              As New ADOX.Key
   Dim lLB              As Long
   Dim lUB              As Long
   Dim l                As Long
   Dim bOK              As Boolean  'bounds the same

   On Error GoTo CreateRelationshipUsingADOXError

   'Get the array boundries - both must be the same
   lLB = LBound(arysForeignTableKeys())
   lUB = UBound(arysForeignTableKeys())

   If lLB = LBound(arysRelatedTableKeys()) Then        'are they the same?
      If lUB = UBound(arysRelatedTableKeys()) Then
         bOK = True
      End If
   End If

   If bOK Then
      ' Create the foreign key to define the relationship.
      With kee
         ' Specify name for the relationship in the Keys collection.
         .Name = sRelationshipName
         ' Specify the related table's name.
         .RelatedTable = sRelatedTable
         .Type = adKeyForeign
         .UpdateRule = lUpdateRule
         .DeleteRule = lDeleteRule
         For l = lLB To lUB
            ' Add the foreign key field(s) to the Columns collection.
            .Columns.Append arysForeignTableKeys(l)
            ' Specify the field the foreign key is related to.
            .Columns(arysForeignTableKeys(l)).RelatedColumn = arysRelatedTableKeys(l)
         Next
      End With

      Set tbl = New ADOX.Table
      ' Open the table and add the foreign key.

      Set tbl = catDB.Tables(sForeignTable)
      tbl.Keys.Append kee
      CreateRelationshipUsingADOX = True
   Else
      Err.Raise vbObjectError + 1, "modCreateRelationshipUsingADOX:CreateRelationshipUsingADOX", "Array bounds differ, arysForeignTableKeys(), arysRelatedTableKeys()"
   End If
   Exit Function

CreateRelationshipUsingADOXError:
   Err.Raise Err.Number, "modCreateRelationshipUsingADOX:CreateRelationshipUsingADOX", Err.Description
End Function

Public Function DeleteRelationshipUsingADOX(catDB As ADOX.Catalog, sTable As String, sRelationshipName As String) As Boolean
   ' #VBIDEUtils#************************************************************
   ' * Author           :  Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modCreateRelationshipUsingADOX
   ' * Procedure Name   : DeleteRelationshipUsingADOX
   ' * Parameters       :
   ' *                    catDB As ADOX.Catalog
   ' *                    sTable As String
   ' *                    sRelationshipName As String
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/03/01 If the relationship exists under this name then delete it
   ' * 2003/04/19 Made a public function and moved to this module.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   Dim kee              As ADOX.Key
   Dim tbl              As ADOX.Table
   
   Set tbl = catDB.Tables(sTable)
   For Each kee In tbl.Keys
      With kee
         If .Name = sRelationshipName Then
            tbl.Keys.Delete .Name   'delete it
            DeleteRelationshipUsingADOX = True
            Exit Function           'bye
         End If
      End With
   Next
   
End Function

Public Function FormatADOXRuleEnum(lRule As ADOX.RuleEnum) As String
   ' #VBIDEUtils#************************************************************
   ' * Author           :  Larry Rebich
   ' * Web Site         : http://www.vbdiamond.com
   ' * E-Mail           : waty.thierry@vbdiamond.com
   ' * Date             : 12/10/2003
   ' * Purpose          :
   ' * Project Name     : DBUpdateADO
   ' * Module Name      : modCreateRelationshipUsingADOX
   ' * Procedure Name   : FormatADOXRuleEnum
   ' * Parameters       :
   ' *                    lRule As ADOX.RuleEnum
   ' **********************************************************************
   ' * Comments         :
   ' *
   ' *
   ' * Example          :
   ' *
   ' * History          : Updated by Waty Thierry
   ' * 2003/03/01 Format the ADOX Foreign Key Enum
   ' * adRINone 0     Default. No action is taken.
   ' * adRICascade 1  Cascade changes.
   ' * adRISetNull 2  Foreign key value is set to null.
   ' * adRISetDefault 3 Foreign key value is set to the default.
   ' *
   ' * See Also         :
   ' *
   ' *
   ' **********************************************************************
   
   Select Case lRule
      Case adRINone
         FormatADOXRuleEnum = "adRINone"
      
      Case adRICascade
         FormatADOXRuleEnum = "adRICascade"
      
      Case adRISetNull
         FormatADOXRuleEnum = "adRISetNull"
      
      Case adRISetDefault
         FormatADOXRuleEnum = "adRISetDefault"
   End Select

End Function


