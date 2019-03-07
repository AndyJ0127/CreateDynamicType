Imports Microsoft.VisualBasic
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Threading

Public Class CreateDynaType

    '當需要使用時才 New 出實體物件
    Private _getJsonCache As GetUriCache = Nothing

    Private ReadOnly Property GetJsonCache() As GetUriCache
        Get
            If (Me._getJsonCache Is Nothing) = True Then
                Me._getJsonCache = New GetUriCache()
            End If
            Return Me._getJsonCache
        End Get
    End Property

    Protected Overrides Sub Finalize()
        Me._getJsonCache = Nothing
        MyBase.Finalize()
    End Sub


    '建立 AssemblyBuilder
    Public Function CreateAssemblyBuilder(Optional assemblyName As String = "") As System.Reflection.Emit.AssemblyBuilder

        If assemblyName.Trim.Length = 0 Then
            assemblyName = Guid.NewGuid.ToString.Trim.Replace("-", "")
        End If

        Dim assName As New System.Reflection.AssemblyName(assemblyName:=assemblyName)

        Return Thread.GetDomain.DefineDynamicAssembly(name:=assName, access:=AssemblyBuilderAccess.RunAndSave)

    End Function

    '建立 ModuleBuilder 01
    Public Function CreateModuleBuilder(Optional assemblyName As String = "") As System.Reflection.Emit.ModuleBuilder

        If assemblyName.Trim.Length = 0 Then
            assemblyName = Guid.NewGuid.ToString.Trim.Replace("-", "")
        End If

        Dim assName As New System.Reflection.AssemblyName(assemblyName:=assemblyName)

        Dim ab As System.Reflection.Emit.AssemblyBuilder = Thread.GetDomain.DefineDynamicAssembly(name:=assName, access:=AssemblyBuilderAccess.RunAndSave)

        Return ab.DefineDynamicModule(name:=assName.Name, fileName:=assName.Name & ".dll")

    End Function

    '建立 ModuleBuilder 02
    Public Function CreateModuleBuilder(ByVal assemblyBuilderer As System.Reflection.Emit.AssemblyBuilder) As System.Reflection.Emit.ModuleBuilder

        If (assemblyBuilderer Is Nothing) = False Then
            Return assemblyBuilderer.DefineDynamicModule(name:=assemblyBuilderer.GetName.Name, fileName:=assemblyBuilderer.GetName.Name & ".dll")
        End If

        Return Nothing

    End Function

    '建立 TypeBuilder 01
    ''' <summary>
    ''' 建立一個 TypeBuilder物件。
    ''' </summary>
    ''' <param name="typeName">傳入 TypeBuilder物件的名稱</param>
    ''' <param name="nameSpaceName">選擇項。傳入要建立 TypeBuilder物件的所屬 Namespace名稱</param>
    ''' <param name="parentType">選擇項。傳入要建立 TypeBuilder物件的所屬 父類別的T Type物件</param>
    ''' <param name="interfaceTypes">選擇項。陣列值。傳入要建立 TypeBuilder物件的所屬 介面的 Type物件陣列</param>
    ''' <param name="propertyNames">選擇項。傳入要建立 TypeBuilder物件的屬性(Property)的名稱字串。字串用逗號分開。例如："propA,propB,propC,propD"</param>
    ''' <param name="propertyTypes">選擇項。傳入要建立 TypeBuilder物件的屬性(Property)的型別名稱字串。字串用逗號分開。例如："string,int,bool,double"</param>
    ''' <param name="customAttributeObj">選擇項。傳入要建立 TypeBuilder物件的 Attribute屬性物件。例如：customAttributeObj:= New Newtonsoft.Json.JsonArrayAttribute()</param>
    ''' <param name="customAttributeValues">選擇項。傳入要建立 TypeBuilder物件的 Attribute屬性名稱字串。字串用逗號分開。例如："propA,propB,propC,propD"。因為某些原本的屬性字串包含不允許的字元故需要在 Attribute中進行對應。</param>
    ''' <returns>傳回已建立的 TypeBuilder物件</returns>
    Public Function CreatTypeBuilder(typeName As String,
                                     Optional nameSpaceName As String = "",
                                     Optional parentType As System.Type = Nothing,
                                     Optional interfaceTypes() As System.Type = Nothing,
                                     Optional propertyNames As String = "",
                                     Optional propertyTypes As String = "",
                                     Optional customAttributeObj As Object = Nothing,
                                     Optional customAttributeValues As String = "") As System.Reflection.Emit.TypeBuilder

        If typeName.Trim.Length = 0 Then
            typeName = Guid.NewGuid.ToString.Trim.Replace("-", "")
        End If

        If nameSpaceName.Trim.Length > 0 Then
            typeName = nameSpaceName & "." & typeName
        End If

        Dim mb As System.Reflection.Emit.ModuleBuilder = Me.CreateModuleBuilder()

        Dim tb As TypeBuilder = Nothing

        tb = mb.DefineType(name:=typeName,
                           attr:=TypeAttributes.Public Or
                                 TypeAttributes.Class Or
                                 TypeAttributes.AutoClass Or
                                 TypeAttributes.AutoLayout Or
                                 TypeAttributes.Sealed,
                           parent:=parentType,
                           interfaces:=interfaceTypes)

        '建立 Property 及 Property's CustomAttribute
        Me.Process_Property(tb:=tb,
                            propertyNames:=propertyNames.Trim,
                            propertyTypes:=propertyTypes.Trim,
                            customAttributeObj:=customAttributeObj,
                            customAttributeValues:=customAttributeValues.Trim)


        Return tb

    End Function


    ''' <summary>
    ''' 建立一個 TypeBuilder物件。
    ''' </summary>
    ''' <param name="typeName">傳入 TypeBuilder物件的名稱</param>
    ''' <param name="nameSpaceName">選擇項。傳入要建立 TypeBuilder物件的所屬 Namespace名稱，例如"NamespaceA.NamespaceB"</param>
    ''' <param name="parentType">選擇項。傳入要建立 TypeBuilder物件的所屬 父類別的T Type物件</param>
    ''' <param name="interfaceTypes">選擇項。陣列值。傳入要建立 TypeBuilder物件的所屬 介面的 Type物件陣列</param>
    ''' <param name="aPropertyNames">選擇項。字串陣列。傳入要建立 TypeBuilder物件的屬性(Property)的名稱字串。字串用逗號分開。例如：New String() {"PropName01", "PropName02", "..."}</param>
    ''' <param name="aPropertyTypes">選擇項。Type陣列。傳入要建立 TypeBuilder物件的屬性(Property)的型別名稱字串。若沒有提供則預設會以 New Type() {GetType(String)}為預設型別。字串用逗號分開。例如：New Type() {GetType(String), GetType(Int32), ....}。若陣列內只提供一個元素則所有的Property全部以此元素當型別，例如：New Type() {GetType(String)}</param>
    ''' <param name="customAttributeObj">選擇項。傳入要建立 TypeBuilder物件的 Attribute屬性物件。例如：customAttributeObj:= New Newtonsoft.Json.JsonArrayAttribute()</param>
    ''' <param name="aCustomAttributeValues">選擇項。字串陣列。傳入要建立 TypeBuilder物件的 Attribute屬性名稱字串。字串用逗號分開。例如：New String() {"PropName01", "PropName02", "..."}。因為某些原本的屬性字串包含不允許的字元故需要在 Attribute中進行對應。</param>
    ''' <returns>傳回已建立的 TypeBuilder物件</returns>
    Public Function CreatTypeBuilder(typeName As String,
                                         Optional nameSpaceName As String = "",
                                         Optional parentType As System.Type = Nothing,
                                         Optional interfaceTypes() As System.Type = Nothing,
                                         Optional aPropertyNames() As String = Nothing,
                                         Optional aPropertyTypes() As Type = Nothing,
                                         Optional customAttributeObj As Object = Nothing,
                                         Optional aCustomAttributeValues() As String = Nothing) As System.Reflection.Emit.TypeBuilder

        If typeName.Trim.Length = 0 Then
            typeName = Guid.NewGuid.ToString.Trim.Replace("-", "")
        End If

        If nameSpaceName.Trim.Length > 0 Then
            typeName = nameSpaceName & "." & typeName
        End If

        Dim mb As System.Reflection.Emit.ModuleBuilder = Me.CreateModuleBuilder()

        Dim tb As TypeBuilder = Nothing

        tb = mb.DefineType(name:=typeName,
                               attr:=TypeAttributes.Public Or
                                     TypeAttributes.Class Or
                                     TypeAttributes.AutoClass Or
                                     TypeAttributes.AutoLayout Or
                                     TypeAttributes.Sealed,
                               parent:=parentType,
                               interfaces:=interfaceTypes)

        '建立 Property 及 Property's CustomAttribute
        Me.Process_Property(tb:=tb,
                                aPropertyNames:=aPropertyNames,
                                aPropertyTypes:=aPropertyTypes,
                                customAttributeObj:=customAttributeObj,
                                aCustomAttributeValues:=aCustomAttributeValues)

        Return tb
    End Function


    ''' <summary>
    ''' 傳回已經將 TypeBuilder物件 轉換成泛型的物件型別
    ''' </summary>
    ''' <param name="tb">傳入已經建立完成的 TypeBuilder 物件</param>
    ''' <param name="genericType">傳入泛型型態的Type，例如 GetType(IList(Of)), GetType(List(Of)), GetType(IEnumerable(Of))</param>
    ''' <returns>傳回已經將 TypeBuilder物件 轉換成泛型的物件型別</returns>
    Public Function Create_GenericCollectionType(tb As TypeBuilder, genericType As Type) As System.Type

        If (tb Is Nothing) = True OrElse (genericType Is Nothing) = True Then
            Return Nothing
        End If

        Return genericType.MakeGenericType(New Type() {tb.CreateType})

    End Function

    '建立 Porperty
    Private Sub CreateProperty(ByRef typeBuilder As System.Reflection.Emit.TypeBuilder,
                                   propertyName As String,
                                   propertyDataType As String,
                                   ByRef customAttributeObj As Object,
                                   customAttributeValue As String)

        Dim newPropType As System.Type = Me.GetUserType(strType:=propertyDataType)

        Dim fieldBuilder As System.Reflection.Emit.FieldBuilder = Nothing
        Dim propBuilder As System.Reflection.Emit.PropertyBuilder = Nothing
        Dim method_Get_Accessor As System.Reflection.Emit.MethodBuilder = Nothing
        Dim method_Set_Accessor As System.Reflection.Emit.MethodBuilder = Nothing

        fieldBuilder = typeBuilder.DefineField(fieldName:="m_" & propertyName,
                                                   type:=newPropType,
                                                   attributes:=FieldAttributes.Private)

        propBuilder = typeBuilder.DefineProperty(name:=propertyName,
                                                      returnType:=newPropType,
                                                      attributes:=PropertyAttributes.HasDefault,
                                                      parameterTypes:=Nothing)


        Dim getSetAttr As MethodAttributes = MethodAttributes.Public Or
                                                 MethodAttributes.SpecialName Or
                                                 MethodAttributes.HideBySig

        '用 TypeBuilder 與 FieldBuilder 設定 Get Method Builder========================================
        method_Get_Accessor = typeBuilder.DefineMethod(name:="get_" & propertyName,
                                                           attributes:=getSetAttr,
                                                           returnType:=newPropType,
                                                           parameterTypes:=System.Type.EmptyTypes)

        Dim getIL As ILGenerator = method_Get_Accessor.GetILGenerator()
        getIL.Emit(opcode:=OpCodes.Ldarg_0)
        getIL.Emit(opcode:=OpCodes.Ldfld, field:=fieldBuilder)
        getIL.Emit(opcode:=OpCodes.Ret)
        '用 TypeBuilder 與  FieldBuilder 設定 Get Method========================================


        '用 TypeBuilder 與  FieldBuilder 設定 Set Method========================================
        method_Set_Accessor = typeBuilder.DefineMethod(name:="set_" & propertyName,
                                                           attributes:=getSetAttr,
                                                           returnType:=Nothing,
                                                           parameterTypes:=New System.Type() {newPropType})

        Dim setIL As ILGenerator = method_Set_Accessor.GetILGenerator()
        setIL.Emit(opcode:=OpCodes.Ldarg_0)
        setIL.Emit(opcode:=OpCodes.Ldarg_1)
        setIL.Emit(opcode:=OpCodes.Stfld, field:=fieldBuilder)
        setIL.Emit(opcode:=OpCodes.Ret)
        '用 TypeBuilder 與  FieldBuilder 設定 Set Method========================================

        '將 Get Method 與 Set Method 加入 Property Builder
        propBuilder.SetGetMethod(method_Get_Accessor)
        propBuilder.SetSetMethod(method_Set_Accessor)


        '將CustomAttributes 加入 Property
        If (customAttributeObj Is Nothing) = False AndAlso customAttributeValue.Trim.Length > 0 Then

            Dim propertyCtorInfo As System.Reflection.ConstructorInfo =
                                    customAttributeObj.GetType().GetConstructor(types:=New Type() {GetType(String)})

            Dim AttrBuilder As New System.Reflection.Emit.CustomAttributeBuilder(con:=propertyCtorInfo,
                                                                                 constructorArgs:=New Object() {customAttributeValue})

            propBuilder.SetCustomAttribute(AttrBuilder)
        End If

    End Sub

    '建立 Porperty
    Private Sub CreateProperty(ByRef typeBuilder As System.Reflection.Emit.TypeBuilder,
                               propertyName As String,
                               propertyDataType As Type,
                               ByRef customAttributeObj As Object,
                               customAttributeValue As String)

        'HttpContext.Current.Response.Write("<BR>propertyDataType=" & propertyDataType.ToString & "<BR>")

        Dim fieldBuilder As System.Reflection.Emit.FieldBuilder = Nothing
        Dim propBuilder As System.Reflection.Emit.PropertyBuilder = Nothing
        Dim method_Get_Accessor As System.Reflection.Emit.MethodBuilder = Nothing
        Dim method_Set_Accessor As System.Reflection.Emit.MethodBuilder = Nothing

        '建立Field
        fieldBuilder = typeBuilder.DefineField(fieldName:="m_" & propertyName,
                                               type:=propertyDataType,
                                               attributes:=FieldAttributes.Private)

        '建立Property
        propBuilder = typeBuilder.DefineProperty(name:=propertyName,
                                                 returnType:=propertyDataType,
                                                 attributes:=PropertyAttributes.HasDefault,
                                                 parameterTypes:=Nothing)

        Dim getSetAttr As MethodAttributes = MethodAttributes.Public Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig

        'Begin 用 TypeBuilder 與 FieldBuilder 設定 Get Method Builder========================================
        method_Get_Accessor = typeBuilder.DefineMethod(name:="get_" & propertyName,
                                                           attributes:=getSetAttr,
                                                           returnType:=propertyDataType,
                                                           parameterTypes:=System.Type.EmptyTypes)

        Dim getIL As ILGenerator = method_Get_Accessor.GetILGenerator()
        getIL.Emit(opcode:=OpCodes.Ldarg_0)
        getIL.Emit(opcode:=OpCodes.Ldfld, field:=fieldBuilder)
        getIL.Emit(opcode:=OpCodes.Ret)
        'End 用 TypeBuilder 與  FieldBuilder 設定 Get Method========================================

        'Begin 用 TypeBuilder 與  FieldBuilder 設定 Set Method========================================
        method_Set_Accessor = typeBuilder.DefineMethod(name:="set_" & propertyName, attributes:=getSetAttr, returnType:=Nothing, parameterTypes:=New System.Type() {propertyDataType})

        Dim setIL As ILGenerator = method_Set_Accessor.GetILGenerator()
        setIL.Emit(opcode:=OpCodes.Ldarg_0)
        setIL.Emit(opcode:=OpCodes.Ldarg_1)
        setIL.Emit(opcode:=OpCodes.Stfld, field:=fieldBuilder)
        setIL.Emit(opcode:=OpCodes.Ret)
        'End 用 TypeBuilder 與  FieldBuilder 設定 Set Method========================================

        '將 Get Method 與 Set Method 加入 Property Builder
        propBuilder.SetGetMethod(method_Get_Accessor)
        propBuilder.SetSetMethod(method_Set_Accessor)

        '將CustomAttributes 加入 Property
        If (customAttributeObj Is Nothing) = False AndAlso customAttributeValue.Trim.Length > 0 Then

            Dim propertyCtorInfo As System.Reflection.ConstructorInfo =
                                        customAttributeObj.GetType().GetConstructor(types:=New Type() {GetType(String)})

            Dim AttrBuilder As New System.Reflection.Emit.CustomAttributeBuilder(con:=propertyCtorInfo,
                                                                                 constructorArgs:=New Object() {customAttributeValue})

            propBuilder.SetCustomAttribute(AttrBuilder)
        End If

    End Sub

    'Get_Json2Object_Async 02
    ''' <summary>
    ''' 非同步方法。直接建立DTO物件並且依據目標網址取得遠端JSON資料後，直接轉換並傳回泛型的物件資料
    ''' </summary>
    ''' <param name="typeName">指定要</param>
    ''' <param name="jsonTargetURL">遠端JSON資源的網址</param>
    ''' <param name="genericType">傳入要轉換的泛型型別。例如：GetType(IList(Of)) or GetType(List(Of)) or GetType(IEnumalbe(Of))</param>
    ''' <param name="cacheMinTime"></param>
    ''' <param name="nameSpaceName"></param>
    ''' <param name="parentType"></param>
    ''' <param name="interfaceTypes"></param>
    ''' <param name="aPropertyNames"></param>
    ''' <param name="aPropertyTypes"></param>
    ''' <param name="customAttributeObj"></param>
    ''' <param name="aCustomAttributeValues"></param>
    ''' <returns>傳回已經內含JSON資料的泛型集合物件</returns>
    Public Async Function Get_Json2Object_Async(typeName As String,
                                                jsonTargetURL As String,
                                                genericType As Type,
                                                Optional cacheMinTime As Integer = 10,
                                                Optional nameSpaceName As String = "",
                                                Optional parentType As System.Type = Nothing,
                                                Optional interfaceTypes() As System.Type = Nothing,
                                                Optional aPropertyNames() As String = Nothing,
                                                Optional aPropertyTypes() As Type = Nothing,
                                                Optional customAttributeObj As Object = Nothing,
                                                Optional aCustomAttributeValues() As String = Nothing) As Tasks.Task(Of Object)


        Dim tb As TypeBuilder = Me.CreatTypeBuilder(typeName:=typeName,
                                                    nameSpaceName:=nameSpaceName,
                                                    parentType:=parentType,
                                                    interfaceTypes:=interfaceTypes,
                                                    aPropertyNames:=aPropertyNames,
                                                    aPropertyTypes:=aPropertyTypes,
                                                    customAttributeObj:=customAttributeObj,
                                                    aCustomAttributeValues:=aCustomAttributeValues)


        Dim genericTbType As Type = Me.Create_GenericCollectionType(tb:=tb, genericType:=genericType)

        Dim methodInfo As MethodInfo = Me.GetJsonCache.GetType().GetMethod("GetJsonCacheAsync")

        Dim genericMethodInfo As MethodInfo = methodInfo.MakeGenericMethod(genericTbType)

        'GetJsonCacheXXXXX(cacheName As String, targetURI As String, cacheMinTime As Integer)
        Dim listResults As Object = Await genericMethodInfo.Invoke(obj:=Me.GetJsonCache,
                                                                   parameters:=New Object() {typeName, jsonTargetURL, cacheMinTime})

        Return listResults

    End Function


    'Get_Json2Object  04
    ''' <summary>
    ''' 同步方法。直接建立DTO物件並且依據目標網址取得遠端JSON資料後，直接轉換並傳回泛型的物件資料
    ''' </summary>
    ''' <param name="typeName"></param>
    ''' <param name="jsonTargetURL"></param>
    ''' <param name="genericType"></param>
    ''' <param name="cacheMinTime"></param>
    ''' <param name="nameSpaceName"></param>
    ''' <param name="parentType"></param>
    ''' <param name="interfaceTypes"></param>
    ''' <param name="aPropertyNames"></param>
    ''' <param name="aPropertyTypes"></param>
    ''' <param name="customAttributeObj"></param>
    ''' <param name="aCustomAttributeValues"></param>
    ''' <returns></returns>
    Public Function Get_Json2Object(typeName As String,
                                    jsonTargetURL As String,
                                    genericType As Type,
                                    Optional cacheMinTime As Integer = 10,
                                    Optional nameSpaceName As String = "",
                                    Optional parentType As System.Type = Nothing,
                                    Optional interfaceTypes() As System.Type = Nothing,
                                    Optional aPropertyNames() As String = Nothing,
                                    Optional aPropertyTypes() As Type = Nothing,
                                    Optional ByRef customAttributeObj As Object = Nothing,
                                    Optional aCustomAttributeValues() As String = Nothing) As Object


        Dim tb As TypeBuilder = Me.CreatTypeBuilder(typeName:=typeName,
                                                    nameSpaceName:=nameSpaceName,
                                                    parentType:=parentType,
                                                    interfaceTypes:=interfaceTypes,
                                                    aPropertyNames:=aPropertyNames,
                                                    aPropertyTypes:=aPropertyTypes,
                                                    customAttributeObj:=customAttributeObj,
                                                    aCustomAttributeValues:=aCustomAttributeValues)

        Dim genericTbType As Type = Me.Create_GenericCollectionType(tb:=tb, genericType:=genericType)

        Dim methodInfo As MethodInfo = Me.GetJsonCache.GetType().GetMethod("GetJsonCache")

        Dim genericMethodInfo As MethodInfo = methodInfo.MakeGenericMethod(genericTbType)

        'GetJsonCacheXXXXX(cacheName As String, targetURI As String, cacheMinTime As Integer)
        Dim listResults = genericMethodInfo.Invoke(obj:=Me.GetJsonCache,
                                                   parameters:=New Object() {typeName, jsonTargetURL, cacheMinTime})

        Return listResults

    End Function


    Private Sub Process_Property(ByRef tb As TypeBuilder,
                                 propertyNames As String,
                                 propertyTypes As String,
                                 ByRef customAttributeObj As Object,
                                 customAttributeValues As String)

        Dim aPropertyNames() As String = Nothing
        Dim aPropertyTypes() As String = Nothing
        Dim aCustomAttributeValues() As String = Nothing

        If propertyNames.Trim.Length > 0 Then
            '代表要求要增加Property

            aPropertyNames = propertyNames.Trim.Split(",")

            If propertyTypes.Trim.Length > 0 Then

                aPropertyTypes = propertyTypes.Trim.Split(",")

                If aPropertyNames.Length = aPropertyTypes.Length Then
                    '此區段的 propertyTypes 要取出來

                    If (customAttributeObj Is Nothing) = False Then

                        If customAttributeValues.Trim.Length > 0 Then

                            aCustomAttributeValues = customAttributeValues.Split(",")

                            If aPropertyNames.Length = aCustomAttributeValues.Length Then

                                For i As Integer = 0 To aPropertyNames.Length - 1

                                    Me.CreateProperty(typeBuilder:=tb,
                                                      propertyName:=aPropertyNames(i),
                                                      propertyDataType:=aPropertyTypes(i),
                                                      customAttributeObj:=customAttributeObj,
                                                      customAttributeValue:=aCustomAttributeValues(i))
                                Next i
                            Else
                                For i As Integer = 0 To aPropertyNames.Length - 1

                                    Me.CreateProperty(typeBuilder:=tb,
                                                      propertyName:=aPropertyNames(i),
                                                      propertyDataType:=aPropertyTypes(i),
                                                      customAttributeObj:=customAttributeObj,
                                                      customAttributeValue:=aPropertyNames(i)) 'aPropertyNames
                                Next i
                            End If 'If propertyNames.Trim.Length = customAttributeValues.Trim.Length Then

                        Else
                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:=aPropertyTypes(i),
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aPropertyNames(i)) 'aPropertyNames
                            Next i
                        End If 'If customAttributeValues.Trim.Length > 0 Then

                    Else
                        For i As Integer = 0 To aPropertyNames.Length - 1

                            Me.CreateProperty(typeBuilder:=tb,
                                              propertyName:=aPropertyNames(i),
                                              propertyDataType:=aPropertyTypes(i),
                                              customAttributeObj:=Nothing,
                                              customAttributeValue:="")
                        Next i

                    End If 'If (customAttributeObj Is Nothing) = False Then
                Else
                    For i As Integer = 0 To aPropertyNames.Length - 1

                        Me.CreateProperty(typeBuilder:=tb,
                                          propertyName:=aPropertyNames(i),
                                          propertyDataType:="string",
                                          customAttributeObj:=Nothing,
                                          customAttributeValue:="")
                    Next i
                End If 'If (customAttributeObj Is Nothing) = False Then

            Else 'If propertyTypes.Trim.Length > 0 Then
                '以下此區段的 propertyDataType 都是 "string"

                If (customAttributeObj Is Nothing) = False Then

                    If customAttributeValues.Trim.Length > 0 Then

                        aCustomAttributeValues = customAttributeValues.Trim.Split(",")

                        If aPropertyNames.Length = aCustomAttributeValues.Length Then

                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:="string",
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aCustomAttributeValues(i))
                            Next i

                        Else
                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:="string",
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aPropertyNames(i)) 'aPropertyNames
                            Next i

                        End If 'If propertyNames.Trim.Length = customAttributeValues.Trim.Length Then

                    Else
                        For i As Integer = 0 To aPropertyNames.Length - 1

                            Me.CreateProperty(typeBuilder:=tb,
                                              propertyName:=aPropertyNames(i),
                                              propertyDataType:="string",
                                              customAttributeObj:=customAttributeObj,
                                              customAttributeValue:=aPropertyNames(i))
                        Next i
                    End If 'If customAttributeValues.Trim.Length > 0 Then
                Else
                    For i As Integer = 0 To aPropertyNames.Length - 1

                        Me.CreateProperty(typeBuilder:=tb,
                                          propertyName:=aPropertyNames(i),
                                          propertyDataType:="string",
                                          customAttributeObj:=Nothing,
                                          customAttributeValue:="")
                    Next i
                End If 'If (customAttributeObj Is Nothing) = False Then

            End If 'If propertyTypes.Trim.Length > 0 Then

        End If 'If propertyNames.Trim.Length > 0 Then

    End Sub 'Process_Property


    Private Sub Process_Property(ByRef tb As TypeBuilder,
                                 aPropertyNames() As String,
                                 aPropertyTypes() As Type,
                                 ByRef customAttributeObj As Object,
                                 aCustomAttributeValues() As String)

        If (aPropertyNames Is Nothing) = False Then

            If (aPropertyTypes Is Nothing) = False Then

                If (aPropertyNames.Length = aPropertyTypes.Length) Then

                    If (customAttributeObj Is Nothing) = False Then

                        If (aCustomAttributeValues Is Nothing) = False Then

                            If aPropertyNames.Length = aCustomAttributeValues.Length Then
                                For i As Integer = 0 To aPropertyNames.Length - 1

                                    Me.CreateProperty(typeBuilder:=tb,
                                                      propertyName:=aPropertyNames(i),
                                                      propertyDataType:=aPropertyTypes(i),
                                                      customAttributeObj:=customAttributeObj,
                                                      customAttributeValue:=aCustomAttributeValues(i))
                                Next i
                            Else
                                For i As Integer = 0 To aPropertyNames.Length - 1

                                    Me.CreateProperty(typeBuilder:=tb,
                                                      propertyName:=aPropertyNames(i),
                                                      propertyDataType:=aPropertyTypes(i),
                                                      customAttributeObj:=customAttributeObj,
                                                      customAttributeValue:=aPropertyNames(i)) 'aPropertyNames
                                Next i

                            End If 'If aPropertyNames.Length = aCustomAttributeValues.Length Then
                        Else
                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:=aPropertyTypes(i),
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aPropertyNames(i))  'aPropertyNames
                            Next i
                        End If 'If (aCustomAttributeValues Is Nothing) = False Then

                    Else
                        For i As Integer = 0 To aPropertyNames.Length - 1

                            Me.CreateProperty(typeBuilder:=tb,
                                              propertyName:=aPropertyNames(i),
                                              propertyDataType:=aPropertyTypes(i),
                                              customAttributeObj:=Nothing,
                                              customAttributeValue:="")
                        Next i
                    End If 'If (customAttributeObj Is Nothing) = False Then

                Else

                    For i As Integer = 0 To aPropertyNames.Length - 1

                        Me.CreateProperty(typeBuilder:=tb,
                                          propertyName:=aPropertyNames(i),
                                          propertyDataType:=aPropertyTypes(0),
                                          customAttributeObj:=Nothing,
                                          customAttributeValue:="") 'propertyDataType:=aPropertyTypes(0) 只取第一個
                    Next i

                End If 'If (aPropertyNames.Length = aPropertyTypes.Length) Then

            Else

                If (customAttributeObj Is Nothing) = False Then

                    If (aCustomAttributeValues Is Nothing) = False Then

                        If aPropertyNames.Length = aCustomAttributeValues.Length Then
                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:=GetType(String),
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aCustomAttributeValues(i))
                            Next i
                        Else
                            For i As Integer = 0 To aPropertyNames.Length - 1

                                Me.CreateProperty(typeBuilder:=tb,
                                                  propertyName:=aPropertyNames(i),
                                                  propertyDataType:=GetType(String),
                                                  customAttributeObj:=customAttributeObj,
                                                  customAttributeValue:=aPropertyNames(i)) 'aPropertyNames
                            Next i

                        End If 'If aPropertyNames.Length = aCustomAttributeValues.Length Then
                    Else
                        For i As Integer = 0 To aPropertyNames.Length - 1

                            Me.CreateProperty(typeBuilder:=tb,
                                              propertyName:=aPropertyNames(i),
                                              propertyDataType:=GetType(String),
                                              customAttributeObj:=customAttributeObj,
                                              customAttributeValue:=aPropertyNames(i))
                        Next i
                    End If 'If (aCustomAttributeValues Is Nothing) = False Then
                Else
                    For i As Integer = 0 To aPropertyNames.Length - 1

                        Me.CreateProperty(typeBuilder:=tb,
                                          propertyName:=aPropertyNames(i),
                                          propertyDataType:=GetType(String),
                                          customAttributeObj:=Nothing,
                                          customAttributeValue:="")
                    Next i
                End If 'If (customAttributeObj Is Nothing) = False Then

            End If 'If (aPropertyTypes Is Nothing) = False Then

        End If 'If (aPropertyNames Is Nothing) = False Then

    End Sub

    ''' <summary>
    ''' 由給定參數的字串，傳回正確的 System.Type；主要是取得泛型集合的 System.Type
    ''' </summary>
    ''' <param name="strType">字串。例如：string, int, int32, long, int64, double, bool, IList(of AAA), IList/<AAA/>, List(Of AAA), List/<AAA/>, IEnumerable(Of AAA), IEnumerable/<AAA/>。</param>
    ''' <returns>傳回正確的 System.Type</returns>
    Private Function GetUserType(ByVal strType As String) As System.Type
        '由給定參數的字串，傳回正確的 System.Type
        '主要是取得泛型集合的 System.Type

        Dim iListGeneric As String = "System.Collections.Generic.IList`1"
        Dim listGeneric As String = "System.Collections.Generic.List`1"
        Dim iEnumerableGeneric As String = "System.Collections.Generic.IEnumerable`1"

        Dim newType As Type = Nothing

        Dim startPos As Integer = 0
        Dim stopPos As Integer = 0
        Dim getStrLength As Integer = 0

        Dim objTypeName As String = String.Empty

        Select Case strType.Trim.ToLower
            Case "string"
                newType = System.Type.GetType(typeName:="System.String", throwOnError:=False, ignoreCase:=True)
            Case "int"
                newType = System.Type.GetType(typeName:="System.Int32", throwOnError:=False, ignoreCase:=True)
            Case "int32"
                newType = System.Type.GetType(typeName:="System.Int32", throwOnError:=False, ignoreCase:=True)
            Case "int64"
                newType = System.Type.GetType(typeName:="System.Int64", throwOnError:=False, ignoreCase:=True)
            Case "long"
                newType = System.Type.GetType(typeName:="System.Int64", throwOnError:=False, ignoreCase:=True)
            Case "double"
                newType = System.Type.GetType(typeName:="System.Double", throwOnError:=False, ignoreCase:=True)
            Case "bool"
                newType = System.Type.GetType(typeName:="System.Boolean", throwOnError:=False, ignoreCase:=True)
            Case Else
                Try
                    newType = System.Type.GetType(typeName:=strType.Trim.ToLower, throwOnError:=True, ignoreCase:=True)
                Catch ex As Exception
                    'HttpContext.Current.Response.Write("ex=" & ex.Message & "<BR>")
                    newType = System.Type.GetType(typeName:="System.String", throwOnError:=False, ignoreCase:=True)
                End Try
        End Select

        'Response.Write("<BR>strType.Trim.ToLower=" & strType.Trim.ToLower.ToString & "<BR>")


        '判斷是不是泛型集合
        If strType.Trim.ToLower.IndexOf("ilist<") >= 0 Then
            '符合 IList<
            startPos = strType.Trim.ToLower.IndexOf("ilist<")
            stopPos = strType.Trim.ToLower.IndexOf(">")
            objTypeName = iListGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 6, length:=stopPos - (startPos + 6)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)
        ElseIf strType.Trim.ToLower.IndexOf("ilist(of") >= 0 Then
            '符合 IList(Of
            '再取出 Of 右邊型別的字串
            startPos = strType.Trim.ToLower.IndexOf("ilist(of")
            stopPos = strType.Trim.ToLower.IndexOf(")")
            objTypeName = iListGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 8, length:=stopPos - (startPos + 8)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)

        ElseIf strType.Trim.ToLower.IndexOf("list<") >= 0 Then
            '符合 List<
            startPos = strType.Trim.ToLower.IndexOf("list<")
            stopPos = strType.Trim.ToLower.IndexOf(">")
            objTypeName = listGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 5, length:=stopPos - (startPos + 5)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)
        ElseIf strType.Trim.ToLower.IndexOf("list(of") >= 0 Then
            '符合 List(Of
            '再取出 Of 右邊型別的字串
            startPos = strType.Trim.ToLower.IndexOf("list(of")
            stopPos = strType.Trim.ToLower.IndexOf(")")
            objTypeName = listGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 7, length:=stopPos - (startPos + 7)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)

        ElseIf strType.Trim.ToLower.IndexOf("ienumerable<") >= 0 Then
            '符合 IEnumerable<
            startPos = strType.Trim.ToLower.IndexOf("ienumerable<")
            stopPos = strType.Trim.ToLower.IndexOf(">")
            objTypeName = iEnumerableGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 12, length:=stopPos - (startPos + 12)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)
        ElseIf strType.Trim.ToLower.IndexOf("ienumerable(of") >= 0 Then
            '符合 IEnumerable(Of
            '再取出 Of 右邊型別的字串
            startPos = strType.Trim.ToLower.IndexOf("ienumerable(of")
            stopPos = strType.Trim.ToLower.IndexOf(")")
            objTypeName = iEnumerableGeneric & "[" & strType.Trim.ToLower.Substring(startIndex:=startPos + 14, length:=stopPos - (startPos + 14)).Trim & "]"
            newType = System.Type.GetType(typeName:=objTypeName, throwOnError:=False, ignoreCase:=True)
        End If

        Return newType

    End Function

End Class
