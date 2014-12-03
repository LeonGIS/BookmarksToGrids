Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.Geodatabase
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.Editor




Public Class frmInput
    Private m_WS As IWorkspace
    Private m_PageNum As String = "PageNumber"
    Private m_PageName As String = "PageName"

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles btnBrowse.Click
        Dim pGxDialog As IGxDialog = New GxDialogClass()
        pGxDialog.ObjectFilter = New GxFilterFileGeodatabases
        pGxDialog.AllowMultiSelect = False

        '  Dim pWS As IWorkspace = Nothing
        Dim pSelectWS As IEnumGxObject = Nothing
        m_WS = Nothing

        If (pGxDialog.DoModalOpen(0, pSelectWS)) And (Not pSelectWS Is Nothing) Then
            Dim pGxObj As IGxObject = pSelectWS.Next

            If TypeOf (pGxObj) Is IGxDatabase Then
                Dim pGxDb As IGxDatabase = CType(pGxObj, IGxDatabase)
                If Not pGxDb Is Nothing Then
                    m_WS = pGxDb.Workspace
                End If

            End If

            System.Runtime.InteropServices.Marshal.ReleaseComObject(pSelectWS)
        End If


        txtWorkspace.Text = m_WS.PathName


        pGxDialog.InternalCatalog.Close()
        System.Runtime.InteropServices.Marshal.ReleaseComObject(pGxDialog)

    End Sub




    Private Sub btnCreate_Click(sender As System.Object, e As System.EventArgs) Handles btnCreate.Click
        Try
            If m_WS Is Nothing Then
                MsgBox("Need to pick a file geodatabase", MsgBoxStyle.Information, "Bookmarks to Grid")
                Exit Sub
            End If

            Dim strFC As String = txtFC.Text
            If strFC = "" Then
                MsgBox("Need to enter feature class name", MsgBoxStyle.Information, "Bookmarks to Grid")
                Exit Sub
            End If

            '** Create Grid Feature Class
            Dim pWS2 As IWorkspace2 = CType(m_WS, IWorkspace2)
            Dim pFC = CreateFeatureClass(pWS2, Nothing, txtFC.Text, Nothing, Nothing, Nothing, "")

            '** Add to map
            Dim pFeatLayer As IFeatureLayer2 = New FeatureLayer
            pFeatLayer.FeatureClass = pFC
            Dim pLayer As ILayer2
            pLayer = CType(pFeatLayer, ILayer2)
            pLayer.Name = txtFC.Text

            My.ArcMap.Document.FocusMap.AddLayer(pLayer)


            '** Convert Bookmarks to grid features
            Dim pEditor As IEditor
            Dim pUID As New UID
            pUID.Value = "esriEditor.Editor"
            pEditor = My.ArcMap.Application.FindExtensionByCLSID(pUID)


            Dim pDS As IDataset
            pDS = pFC
            pEditor.StartEditing(pDS.Workspace)
            pEditor.StartOperation()

            '** Get Bookmarks
            Dim pMapBookmarks As IMapBookmarks
            Dim pMarks As IEnumSpatialBookmark
            Dim pBookmark As ISpatialBookmark
            Dim pAOI As IAOIBookmark
            Dim pEnv As IEnvelope
            Dim pNewFeat As IFeature

            Dim pFeatBuff As IFeatureBuffer = pFC.CreateFeatureBuffer()
            Dim pFeatCursor As IFeatureCursor = pFC.Insert(True)




            pMapBookmarks = My.ArcMap.Document.FocusMap
            pMarks = pMapBookmarks.Bookmarks
            pBookmark = pMarks.Next
            Dim intPage As Integer = 1

            Do While Not pBookmark Is Nothing

                If TypeOf pBookmark Is AOIBookmark Then
                    pAOI = pBookmark
                    pEnv = pAOI.Location

                    pFeatBuff.Shape = PointsToPoly(pEnv)
                    pFeatBuff.Value(pFC.FindField(m_PageNum)) = intPage
                    pFeatBuff.Value(pFC.FindField(m_PageName)) = pBookmark.Name

                    pFeatCursor.InsertFeature(pFeatBuff)

                    intPage = intPage + 1

                End If



                pBookmark = pMarks.Next
            Loop



            pEditor.StopOperation("Bookmarks to Grid")
            pEditor.StopEditing(True)

           
            '** Release objects

            '** Refresh view
            My.ArcMap.Document.ActiveView.Refresh()
            Me.Close()

            Exit Sub

        Catch ex As Exception
            MsgBox("Unhandled Error: " & ex.Message, MsgBoxStyle.Exclamation, "Bookmarks to Grid")
        End Try
    End Sub


    Private Function PointsToPoly(pEnv As IEnvelope) As Polygon
        Dim pPointsColl As IPointCollection
        pPointsColl = New Polygon

        Dim pPoint1 As IPoint
        pPoint1 = New Point
        Dim pPoint2 As IPoint
        pPoint2 = New Point
        Dim pPoint3 As IPoint
        pPoint3 = New Point
        Dim pPoint4 As IPoint
        pPoint4 = New Point
        pPoint1.PutCoords(pEnv.XMin, pEnv.YMin)
        pPoint2.PutCoords(pEnv.XMax, pEnv.YMin)
        pPoint3.PutCoords(pEnv.XMax, pEnv.YMax)
        pPoint4.PutCoords(pEnv.XMin, pEnv.YMax)

        pPointsColl.AddPoint(pPoint1)
        pPointsColl.AddPoint(pPoint2)
        pPointsColl.AddPoint(pPoint3)
        pPointsColl.AddPoint(pPoint4)

        Dim pTopo As ITopologicalOperator
        pTopo = pPointsColl
        pTopo.Simplify()

        PointsToPoly = pPointsColl
    End Function


#Region "Create FeatureClass"
    ' ArcGIS Snippet Title:
    ' Create FeatureClass
    ' 
    ' Long Description:
    ' Simple helper to create a featureclass in a geodatabase.
    ' 
    ' Add the following references to the project:
    ' ESRI.ArcGIS.Geodatabase
    ' ESRI.ArcGIS.Geometry
    ' ESRI.ArcGIS.System
    ' System
    ' 
    ' Intended ArcGIS Products for this snippet:
    ' ArcGIS Desktop (ArcEditor, ArcInfo, ArcView)
    ' ArcGIS Engine
    ' ArcGIS Server
    ' 
    ' Applicable ArcGIS Product Versions:
    ' 9.2
    ' 9.3
    ' 9.3.1
    ' 10.0
    ' 
    ' Required ArcGIS Extensions:
    ' (NONE)
    ' 
    ' Notes:
    ' This snippet is intended to be inserted at the base level of a Class.
    ' It is not intended to be nested within an existing Function or Sub.
    ' 

    '''<summary>Simple helper to create a featureclass in a geodatabase.</summary>
    ''' 
    '''<param name="workspace">An IWorkspace2 interface</param>
    '''<param name="featureDataset">An IFeatureDataset interface or Nothing</param>
    '''<param name="featureClassName">A System.String that contains the name of the feature class to open or create. Example: "states"</param>
    '''<param name="fields">An IFields interface</param>
    '''<param name="CLSID">A UID value or Nothing. Example "esriGeoDatabase.Feature" or Nothing</param>
    '''<param name="CLSEXT">A UID value or Nothing (this is the class extension if you want to reference a class extension when creating the feature class).</param>
    '''<param name="strConfigKeyword">An empty System.String or RDBMS table string for ArcSDE. Example: "myTable" or ""</param>
    '''  
    '''<returns>An IFeatureClass interface or a Nothing</returns>
    '''  
    '''<remarks>
    '''  (1) If a 'featureClassName' already exists in the workspace a reference to that feature class 
    '''      object will be returned.
    '''  (2) If an IFeatureDataset is passed in for the 'featureDataset' argument the feature class
    '''      will be created in the dataset. If a Nothing is passed in for the 'featureDataset'
    '''      argument the feature class will be created in the workspace.
    '''  (3) When creating a feature class in a dataset the spatial reference is inherited 
    '''      from the dataset object.
    '''  (4) If an IFields interface is supplied for the 'fields' collection it will be used to create the
    '''      table. If a Nothing value is supplied for the 'fields' collection, a table will be created using 
    '''      default values in the method.
    '''  (5) The 'strConfigurationKeyword' parameter allows the application to control the physical layout 
    '''      for this table in the underlying RDBMS—for example, in the case of an Oracle database, the 
    '''      configuration keyword controls the tablespace in which the table is created, the initial and 
    '''     next extents, and other properties. The 'strConfigurationKeywords' for an ArcSDE instance are 
    '''      set up by the ArcSDE data administrator, the list of available keywords supported by a workspace 
    '''      may be obtained using the IWorkspaceConfiguration interface. For more information on configuration 
    '''      keywords, refer to the ArcSDE documentation. When not using an ArcSDE table use an empty 
    '''      string (ex: "").
    '''</remarks>
    Public Function CreateFeatureClass(ByVal workspace As ESRI.ArcGIS.Geodatabase.IWorkspace2, ByVal featureDataset As ESRI.ArcGIS.Geodatabase.IFeatureDataset, ByVal featureClassName As System.String, ByVal fields As ESRI.ArcGIS.Geodatabase.IFields, ByVal CLSID As ESRI.ArcGIS.esriSystem.UID, ByVal CLSEXT As ESRI.ArcGIS.esriSystem.UID, ByVal strConfigKeyword As System.String) As ESRI.ArcGIS.Geodatabase.IFeatureClass

        If featureClassName = "" Then
            Return Nothing ' name was not passed in
        End If

        Dim featureClass As ESRI.ArcGIS.Geodatabase.IFeatureClass
        Dim featureWorkspace As ESRI.ArcGIS.Geodatabase.IFeatureWorkspace = CType(workspace, ESRI.ArcGIS.Geodatabase.IFeatureWorkspace) ' Explicit Cast

        If workspace.NameExists(ESRI.ArcGIS.Geodatabase.esriDatasetType.esriDTFeatureClass, featureClassName) Then
            featureClass = featureWorkspace.OpenFeatureClass(featureClassName) ' feature class with that name already exists
            Return featureClass
        End If

        ' assign the class id value if not assigned
        If CLSID Is Nothing Then
            CLSID = New ESRI.ArcGIS.esriSystem.UIDClass
            CLSID.Value = "esriGeoDatabase.Feature"
        End If

        Dim objectClassDescription As ESRI.ArcGIS.Geodatabase.IObjectClassDescription = New ESRI.ArcGIS.Geodatabase.FeatureClassDescription



        ' if a fields collection is not passed in then supply our own
        If fields Is Nothing Then


            fields = CreateFieldsCollectionForFeatureClass(My.ArcMap.Document.FocusMap.SpatialReference)


            '' create the fields using the required fields method
            'fields = objectClassDescription.RequiredFields
            'Dim fieldsEdit As ESRI.ArcGIS.Geodatabase.IFieldsEdit = CType(fields, ESRI.ArcGIS.Geodatabase.IFieldsEdit) ' Explict Cast
            'Dim field As ESRI.ArcGIS.Geodatabase.IField = New ESRI.ArcGIS.Geodatabase.Field

            '' create a user defined text field
            'Dim fieldEdit As ESRI.ArcGIS.Geodatabase.IFieldEdit = CType(Field, ESRI.ArcGIS.Geodatabase.IFieldEdit) ' Explict Cast

            '' setup field properties
            'fieldEdit.Name_2 = "SampleField"
            'fieldEdit.Type_2 = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeString
            'fieldEdit.IsNullable_2 = True
            'fieldEdit.AliasName_2 = "Sample Field Column"
            'fieldEdit.DefaultValue_2 = "test"
            'fieldEdit.Editable_2 = True
            'fieldEdit.Length_2 = 100

            '' add field to field collection
            'fieldsEdit.AddField(Field)
            'fields = CType(fieldsEdit, ESRI.ArcGIS.Geodatabase.IFields) ' Explicit Cast

        End If

        Dim strShapeField As System.String = ""

        ' locate the shape field
        Dim j As System.Int32
        For j = 0 To fields.FieldCount - 1
            If fields.Field(j).Type = ESRI.ArcGIS.Geodatabase.esriFieldType.esriFieldTypeGeometry Then
                strShapeField = fields.Field(j).Name
            End If
        Next j

        ' Use IFieldChecker to create a validated fields collection.
        Dim fieldChecker As ESRI.ArcGIS.Geodatabase.IFieldChecker = New ESRI.ArcGIS.Geodatabase.FieldChecker
        Dim enumFieldError As ESRI.ArcGIS.Geodatabase.IEnumFieldError = Nothing
        Dim validatedFields As ESRI.ArcGIS.Geodatabase.IFields = Nothing
        fieldChecker.ValidateWorkspace = CType(workspace, ESRI.ArcGIS.Geodatabase.IWorkspace)
        fieldChecker.Validate(fields, enumFieldError, validatedFields)

        ' The enumFieldError enumerator can be inspected at this point to determine 
        ' which fields were modified during validation.


        ' finally create and return the feature class
        If featureDataset Is Nothing Then
            ' if no feature dataset passed in, create at the workspace level
            featureClass = featureWorkspace.CreateFeatureClass(featureClassName, validatedFields, CLSID, CLSEXT, ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimple, strShapeField, strConfigKeyword)
        Else
            featureClass = featureDataset.CreateFeatureClass(featureClassName, validatedFields, CLSID, CLSEXT, ESRI.ArcGIS.Geodatabase.esriFeatureType.esriFTSimple, strShapeField, strConfigKeyword)
        End If

        Return featureClass

    End Function

    Public Function CreateFieldsCollectionForFeatureClass(ByVal spatialReference As ISpatialReference) As IFields

        ' Use the feature class description to return the required fields in a fields collection.
        Dim fcDesc As IFeatureClassDescription = New FeatureClassDescriptionClass()
        Dim ocDesc As IObjectClassDescription = CType(fcDesc, IObjectClassDescription)

        ' Create the fields using the required fields method.
        Dim fields As IFields = ocDesc.RequiredFields

        ' Locate the shape field with the name from the feature class description.
        Dim shapeFieldIndex As Integer = fields.FindField(fcDesc.ShapeFieldName)
        Dim shapeField As IField = fields.Field(shapeFieldIndex)

        ' Modify the GeometryDef object before using the fields collection to create a feature class.
        Dim geometryDef As IGeometryDef = shapeField.GeometryDef
        Dim geometryDefEdit As IGeometryDefEdit = CType(geometryDef, IGeometryDefEdit)

        ' Alter the feature class geometry type to lines (default is polygons).
        geometryDefEdit.GeometryType_2 = esriGeometryType.esriGeometryPolygon
        geometryDefEdit.HasM_2 = False
        geometryDefEdit.GridCount_2 = 1

        ' Set the first grid size to zero and allow ArcGIS to determine a valid grid size.
        geometryDefEdit.GridSize_2(0) = 0
        geometryDefEdit.SpatialReference_2 = spatialReference

        ' Because the fields collection already exists, the AddField method on the IFieldsEdit
        ' interface will be used to add a field that is not required to the fields collection.
        Dim fieldsEdit As IFieldsEdit = CType(fields, IFieldsEdit)
        Dim pNameField As IField = New FieldClass()
        Dim pNameFieldEdit As IFieldEdit = CType(pNameField, IFieldEdit)

        ' Create a user-defined double field.

        pNameFieldEdit.Editable_2 = True
        pNameFieldEdit.IsNullable_2 = False
        pNameFieldEdit.Name_2 = m_PageName
        pNameFieldEdit.Length_2 = 255
        pNameFieldEdit.Type_2 = esriFieldType.esriFieldTypeString
        fieldsEdit.AddField(pNameField)

        Dim pNumberField As IField = New FieldClass()
        Dim pNumFieldEdit As IFieldEdit = CType(pNumberField, IFieldEdit)
        pNumFieldEdit.Editable_2 = True
        pNumFieldEdit.IsNullable_2 = False
        pNumFieldEdit.Name_2 = m_PageNum
        pNumFieldEdit.Type_2 = esriFieldType.esriFieldTypeInteger
        fieldsEdit.AddField(pNumberField)

        Return fields

    End Function



#End Region



End Class