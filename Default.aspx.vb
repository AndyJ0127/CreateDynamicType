Imports System.Reflection.Emit
Imports Newtonsoft.Json

Partial Class _Default
    Inherits System.Web.UI.Page

    Private Async Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load

        Dim targetURL As String = "http://opendata.ilepb.gov.tw/data/ilepb02020"

        Dim dynaType As New CreateDynaType()

        Dim newTB01 As TypeBuilder = dynaType.CreatTypeBuilder(typeName:="ILEPB02020",
                                                               nameSpaceName:="DTO",
                                                               parentType:=GetType(DTO.Base_DTO),
                                                               interfaceTypes:=Nothing,
                                                               propertyNames:="項次,補助對象,補助項目,新舊車類別,淘汰或換新購類別,換新購之車型,補助總額,補助限制",
                                                               propertyTypes:="string,string,string,string,string,string,int,string",
                                                               customAttributeObj:=New Newtonsoft.Json.JsonPropertyAttribute(),
                                                               customAttributeValues:="項次,補助對象,補助項目,新舊車類別,淘汰或換、新購類別,換、新購之車型,補助總額,補助限制")


        Dim newTB02 As TypeBuilder = dynaType.CreatTypeBuilder(typeName:="ILEPB02020",
                                                               nameSpaceName:="DTO",
                                                               parentType:=GetType(DTO.Base_DTO),
                                                               interfaceTypes:=Nothing,
                                                               aPropertyNames:=New String() {"項次", "補助對象", "補助項目", "新舊車類別", "淘汰或換新購類別", "換新購之車型", "補助總額,補助限制"},
                                                               aPropertyTypes:=New Type() {GetType(String), GetType(String), GetType(String), GetType(String), GetType(String), GetType(String), GetType(Int32), GetType(String)},
                                                               customAttributeObj:=New Newtonsoft.Json.JsonPropertyAttribute(),
                                                               aCustomAttributeValues:=New String() {"項次", "補助對象", "補助項目", "新舊車類別", "淘汰或換、新購類別", "換、新購之車型", "補助總額,補助限制"})


        Dim jsonList01 As Object = Await dynaType.Get_Json2Object_Async(typeName:="ILEPB02020",
                                                                        nameSpaceName:="DTO",
                                                                        parentType:=GetType(DTO.Base_DTO),
                                                                        interfaceTypes:=Nothing,
                                                                        jsonTargetURL:=targetURL,
                                                                        genericType:=GetType(IList(Of)),
                                                                        cacheMinTime:=10,
                                                                        propertyNames:="項次,補助對象,補助項目,新舊車類別,淘汰或換新購類別,換新購之車型,補助總額,補助限制",
                                                                        propertyTypes:="string,string,string,string,string,string,int,string",
                                                                        customAttributeObj:=New Newtonsoft.Json.JsonPropertyAttribute(),
                                                                        customAttributeValues:="項次,補助對象,補助項目,新舊車類別,淘汰或換、新購類別,換、新購之車型,補助總額,補助限制")


        Dim jsonList02 As Object = Await dynaType.Get_Json2Object_Async(typeName:="ILEPB02020",
                                                                        nameSpaceName:="DTO",
                                                                        parentType:=GetType(DTO.Base_DTO),
                                                                        interfaceTypes:=Nothing,
                                                                        jsonTargetURL:=targetURL,
                                                                        genericType:=GetType(IList(Of)),
                                                                        cacheMinTime:=10,
                                                                        aPropertyNames:=New String() {"項次", "補助對象", "補助項目", "新舊車類別", "淘汰或換新購類別", "換新購之車型", "補助總額,補助限制"},
                                                                        aPropertyTypes:=New Type() {GetType(String), GetType(String), GetType(String), GetType(String), GetType(String), GetType(String), GetType(Int32), GetType(String)},
                                                                        customAttributeObj:=New Newtonsoft.Json.JsonPropertyAttribute(),
                                                                        aCustomAttributeValues:=New String() {"項次", "補助對象", "補助項目", "新舊車類別", "淘汰或換、新購類別", "換、新購之車型", "補助總額,補助限制"})

    End Sub
End Class
