Attribute VB_Name = "Module3"
' based on IFCsvrR300 sample7

' -----------------------------------------------------------------
' Global Setting
' -----------------------------------------------------------------

Option Explicit
  Public objIFCsvr As IFCsvr.R300
  Public objDesign As IFCsvr.Design
  
  Const STR_ENT_REL_ASSIGNS_PROPERTIES_2X = "IfcRelDefinesByProperties"
  Const STR_ATT_RELATED_OBJECTS_2X = "RelatedObjects"
  Const STR_ATT_RELATING_PROPERTY_2X = "RelatingPropertyDefinition"
  Const STR_ENT_PROPERTY_SET_2X = "IfcPropertySet"
  Const STR_ENT_PROPERTY_SINBLEVALUE_2X = "IfcPropertySingleValue"
  Const STR_ATT_PROPERTY_VALUE_2X = "NominalValue"

  Dim str_ent_rel_assigns_properties As String
  Dim str_att_related_objects As String
  Dim str_att_relating_property As String
  Dim str_ent_property_set As String

' -----------------------------------------------------------------
' procPsetViewer in Grid Form for IFC2xx
' -----------------------------------------------------------------

Public Function procPsetViewerGrid()
  Dim filename As String
  Dim strClassName As String
  Dim objEntity As IFCsvr.Entity
  Dim objEntities As IFCsvr.Entities
  Dim objRelAssignsProperties As IFCsvr.Entity
  Dim objPropertySet As IFCsvr.Entity
  Dim objProperty As IFCsvr.Entity
  Dim r1 As Excel.Range
  Dim strPropertySetName As String
  Dim strPropertyType As String
  Dim strPropertyName As String
  
  Dim strSchemaName As String
  
  filename = ActiveSheet.Range("C7").Value

  Set objIFCsvr = New IFCsvr.R300
  
  If objIFCsvr Is Nothing Then
    MsgBox "IFCsvr is not installed."
    Exit Function
  End If
  
  Set objDesign = objIFCsvr.OpenDesign(filename)
  If objDesign Is Nothing Then
    MsgBox "IFC file not found."
    Set objDesign = Nothing
    Set objIFCsvr = Nothing
    Exit Function
  End If
  
  strClassName = ActiveSheet.Range("C8").Value
  If Len(strClassName) = 0 Then
    MsgBox "IFC Class name not defined."
    Exit Function
  End If
  
  strSchemaName = UCase(objDesign.SchemaName)
  ActiveSheet.Range("C9").Value = strSchemaName
   
  Set r1 = ActiveSheet.Range("A12")
    
    
    
' 属性名の列挙
  Dim i As Integer
  Dim j As Integer
    i = 0
    j = 1
  ' IfcSpaceの繰り返し　1回だけ
  For Each objEntity In objDesign.FindObjects(strClassName)
    If i > 0 Then
        Exit For
    End If
    i = i + 1
    'r1.Value = objEntity.Type
    'r1.Offset(0, 1).Value = "{" & objEntity.GUID & "}"
    'Set r1 = r1.Offset(1, 0)
    
    ' Psetの繰り返し
    For Each objRelAssignsProperties In _
          objEntity.GetUsedIn(STR_ENT_REL_ASSIGNS_PROPERTIES_2X, STR_ATT_RELATED_OBJECTS_2X)
      Set objPropertySet = objRelAssignsProperties.Attributes(STR_ATT_RELATING_PROPERTY_2X).Value
      If Not objPropertySet Is Nothing Then
        ' Psetが空でないとき
        If objPropertySet.Type Like STR_ENT_PROPERTY_SET_2X Then
        j = j + 1
          ' PsetがIfcPropertySetのとき
          strPropertySetName = objPropertySet.Attributes("Name").Value
          'r1.Offset(0, 1).Value = "PropertySet Name:"
          r1.Offset(0, j).Value = strPropertySetName
          'Set r1 = r1.Offset(1, 0)

          For Each objProperty In objPropertySet.Attributes("HasProperties").Value
            ' Pset内のHasPropertiesに対して繰り返し
            j = j + 1
            strPropertyType = objProperty.Type
            strPropertyName = objProperty.Attributes("Name").Value
            'r1.Offset(0, 2).Value = strPropertyType
            r1.Offset(0, j).Value = strPropertyName
            If strPropertyType Like STR_ENT_PROPERTY_SINBLEVALUE_2X Then
              ' プロパティがIfcPropertySingleValueのとき
              'r1.Offset(0, 4).Value = objProperty.Attributes("NominalValue").GetSelectType
              'r1.Offset(0, 5).Value = objProperty.Attributes("NominalValue").Value
            End If
            'Set r1 = r1.Offset(1, 0)
          Next objProperty
          
         
        End If
      End If
       
    Next objRelAssignsProperties
    Set r1 = r1.Offset(1, 0)
  Next objEntity
  



'属性の列挙
  ' IfcSpaceの繰り返し
  For Each objEntity In objDesign.FindObjects(strClassName)
  j = 0
    r1.Value = objEntity.Type
    j = j + 1
    r1.Offset(0, j).Value = "{" & objEntity.GUID & "}"
    'Set r1 = r1.Offset(1, 0)
    
    ' Psetの繰り返し
    For Each objRelAssignsProperties In _
          objEntity.GetUsedIn(STR_ENT_REL_ASSIGNS_PROPERTIES_2X, STR_ATT_RELATED_OBJECTS_2X)
      Set objPropertySet = objRelAssignsProperties.Attributes(STR_ATT_RELATING_PROPERTY_2X).Value
      If Not objPropertySet Is Nothing Then
        ' Psetが空でないとき
        If objPropertySet.Type Like STR_ENT_PROPERTY_SET_2X Then
        j = j + 1
          ' PsetがIfcPropertySetのとき
          strPropertySetName = objPropertySet.Attributes("Name").Value
          'r1.Offset(0, 1).Value = "PropertySet Name:"
          'r1.Offset(0, j).Value = strPropertySetName
          'Set r1 = r1.Offset(1, 0)

          For Each objProperty In objPropertySet.Attributes("HasProperties").Value
            ' Pset内のHasPropertiesに対して繰り返し
            j = j + 1
            strPropertyType = objProperty.Type
            strPropertyName = objProperty.Attributes("Name").Value
            'r1.Offset(0, 2).Value = strPropertyType
            'r1.Offset(0, j).Value = strPropertyName
            If strPropertyType Like STR_ENT_PROPERTY_SINBLEVALUE_2X Then
              ' プロパティがIfcPropertySingleValueのとき
              'r1.Offset(0, 4).Value = objProperty.Attributes("NominalValue").GetSelectType
              r1.Offset(0, j).Value = objProperty.Attributes("NominalValue").Value
            End If
            'Set r1 = r1.Offset(1, 0)
          Next objProperty
          
         
        End If
      End If
       
    Next objRelAssignsProperties
    Set r1 = r1.Offset(1, 0)
  Next objEntity
  
  
  
  Set objDesign = Nothing
  Set objIFCsvr = Nothing
  
End Function





