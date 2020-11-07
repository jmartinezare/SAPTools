Imports System.Globalization

Module General

    Public SEP_DECIMAL_UI As String = "." ' Decimal separator when a value is get from a form

    NotInheritable Class Memoria
        Public Shared Sub LiberarRecordSet(rs As SAPbobsCOM.Recordset)
            rs = Nothing
            GC.Collect()
        End Sub
    End Class


    ''' <summary>
    ''' Tools class with functions to help in development process
    ''' </summary>
    ''' <remarks></remarks>
    NotInheritable Class Tools

        ''' <summary>
        ''' Get a decimal number from a string
        ''' </summary>
        ''' <param name="s">string</param>
        ''' <param name="def_val">default value</param>
        ''' <param name="sep_decimal_from">decimal character</param>
        ''' <returns>decimal number</returns>
        ''' <remarks></remarks>
        Public Shared Function ToDecimal(s As String, Optional def_val As Decimal = 0, Optional sep_decimal_from As String = ",") As Decimal
            Dim resu As Decimal = def_val
            If Not (s Is Nothing Or Trim(s) = "") Then
                Try
                    If sep_decimal_from = "." Then
                        resu = CDbl(s.Replace(",", "").Replace(".", ","))
                    Else
                        resu = CDbl(s)
                    End If
                Catch ex As Exception
                    resu = def_val
                End Try
            End If
            Return resu
        End Function

        Public Shared Function ToInteger(s As String, Optional def_val As Integer = 0) As Integer
            Dim resu As Integer = def_val
            If Not (s Is Nothing Or Trim(s) = "") Then
                Try
                    resu = CInt(s)
                Catch ex As Exception
                    resu = def_val
                End Try
            End If
            Return resu
        End Function

        Public Shared Function DecToStr(d As Decimal, Optional sep_decimal_to As String = ",") As String
            Dim resu As String = "0"

            Try
                If Trim(sep_decimal_to) = "." Then
                    resu = d.ToString().Replace(".", "").Replace(",", ".")
                Else
                    resu = d.ToString()
                End If
            Catch ex As Exception
                resu = "0"
            End Try

            Return resu
        End Function
        
        Public Shared Function ToDateTime(s As String, Optional def_val As DateTime? = Nothing, Optional format As String = "dd/MM/yyyy h:mm:ss") As DateTime?
            Dim resuIni As DateTime? = Nothing
            If Not def_val Is Nothing Then
                resuIni = def_val
            End If
            Dim resu As DateTime? = resuIni

            If Not (s Is Nothing Or Trim(s) = "") Then
                Try
                    resu = DateTime.ParseExact(s, format, System.Globalization.CultureInfo.InvariantCulture)
                Catch ex As Exception
                    resu = resuIni
                End Try
            End If
            Return resu
        End Function

        Public Shared Function ToDate(s As String, Optional def_val As Date? = Nothing, Optional format As String = "yyyyMMdd") As Date?
            Dim resu As Date? = Nothing
            If Not resu Is Nothing Then
                resu = def_val
            End If
            If Not (s Is Nothing Or Trim(s) = "") Then
                Try
                    resu = Date.ParseExact(s, format, System.Globalization.CultureInfo.InvariantCulture)
                Catch ex As Exception
                    resu = def_val
                End Try
            End If
            Return resu
        End Function

        Public Shared Function ToDateStr(d As Date, Optional def_val As String = "", Optional format As String = "yyyyMMdd") As String
            If Trim(def_val) = "" Then
                def_val = Date.Today
                Date.Today.ToString(format, CultureInfo.InvariantCulture)
            End If
            Dim resu As String = def_val

            Try
                resu = d.ToString(format, CultureInfo.InvariantCulture)
            Catch ex As Exception
                resu = def_val
            End Try

            Return resu
        End Function

        ''' <summary>
        ''' Set a matrix field as editable. First, the field must be editable in the form settings and maybe changed as not editable in other moment like the load event.
        ''' </summary>        
        Public Shared Function MatrixSetColEditable(m As SAPbouiCOM.Matrix, colName As String, editable As Boolean) As Integer
            Dim resu As Integer = 1
            Try
                If m.Columns.Item(colName).Editable = Not editable Then
                    m.Columns.Item(colName).Editable = editable
                End If
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a matrix field value by it's name.
        ''' </summary>        
        Public Shared Function MatrixValue(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As Object
            Dim resu As Object
            Try
                resu = m.Columns.Item(colName).Cells.Item(row).Specific.Value
            Catch ex As Exception
                resu = New Object
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' From a matrix, get a combobox value.
        ''' </summary> 
        Public Shared Function MatrixValueCombo(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As String
            Dim resu As String = ""
            Try
                Dim Combo = m.Columns.Item(colName).Cells.Item(row).Specific
                If Not Combo Is Nothing Then
                    resu = Combo.Selected.Value.ToString()
                End If
            Catch ex As Exception
                resu = ""
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' From a matrix, get a field value as string.
        ''' </summary>        
        Public Shared Function MatrixValueStr(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As String
            Return Tools.MatrixValue(m, colName, row).ToString()
        End Function

        ''' <summary>
        ''' From a matrix, get a field value as decimal.
        ''' </summary> 
        Public Shared Function MatrixValueDecimal(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As Decimal
            Return Tools.ToDecimal(Tools.MatrixValue(m, colName, row).ToString(), sep_decimal_from:=SEP_DECIMAL_UI)
        End Function

        ''' <summary>
        ''' From a matrix, get a currency field value as Decimal. Filter the currency name from de value. Ej: "15,00 EUR" --> 15,00
        ''' </summary>        
        Public Shared Function MatrixValueCurrency(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As Decimal
            Dim value As String = Tools.MatrixValueStr(m, colName, row).ToString()
            Dim lstValues() As String = Split(value, " ")
            Dim resu As Decimal = 0
            If lstValues.Length > 0 Then
                resu = Tools.ToDecimal(lstValues(0))
            End If
            Return resu
        End Function

        ''' <summary>
        ''' From a matrix, get a check field value as Boolean.
        ''' </summary>
        ''' <param name="m"></param>
        ''' <param name="colName"></param>
        ''' <param name="row"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function MatrixValueChk(m As SAPbouiCOM.Matrix, colName As String, row As Integer) As Boolean
            Dim resu As Boolean = False
            Try
                resu = m.Columns.Item(colName).Cells.Item(row).Specific.Checked
            Catch ex As Exception
                resu = False
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' In a matrix, set a field value. If it is needed, deal with the editable status of the field. The field must be editable in the form settings.
        ''' </summary>        
        Public Shared Function MatrixSetValue(m As SAPbouiCOM.Matrix, colName As String, row As Integer, value As String, Optional NotChangeFocus As Boolean = True) As Integer
            Dim resu As Integer = 1
            Dim CelFocus = m.GetCellFocus()
            Dim EsEditable = True
            Try
                If Not m.Columns.Item(colName).Editable Then
                    EsEditable = False
                    m.Columns.Item(colName).Editable = True
                End If

                m.Columns.Item(colName).Cells.Item(row).Specific.Value = value

                'NotChangeFocus = True means that it keeps the focus in the same position where it was before modify field value.
                If NotChangeFocus Then
                    m.SetCellFocus(CelFocus.rowIndex, CelFocus.ColumnIndex)
                End If

                If Not EsEditable Then
                    m.Columns.Item(colName).Editable = False
                End If
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' In a matrix, set a decimal field value. The field must be editable in the form settings.
        ''' </summary>    
        Public Shared Function MatrixSetValueDecimal(m As SAPbouiCOM.Matrix, colName As String, row As Integer, value As Decimal, Optional colIsNotEditable As Boolean = False) As Integer
            Dim resu As Integer = 1
            Try
                Dim valTxt = Tools.DecToStr(value, ".")
                If colIsNotEditable Then 'Set as editable before modify the value
                    Tools.MatrixSetColEditable(m, colName, True)
                End If

                m.Columns.Item(colName).Cells.Item(row).Specific.Value = valTxt

                If colIsNotEditable Then
                    Tools.MatrixSetColEditable(m, colName, False)
                End If
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function


        ''' <summary>
        ''' In a matrix, set a combobox field value. The field must be editable in the form settings.
        ''' </summary>        
        Public Shared Function MatrixSetCombo(m As SAPbouiCOM.Matrix, colName As String, row As Integer, value As String) As Integer
            Dim resu As Integer = 1
            Try
                Dim obj As SAPbouiCOM.ComboBox = m.Columns.Item(colName).Cells.Item(row).Specific
                obj.Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue)                
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a Recordset value by the name of the field. 
        ''' </summary>        
        Public Shared Function RSetValue(rs As SAPbobsCOM.Recordset, colName As String, Optional def_val As String = "") As Object
            Dim resu As Object
            Try
                resu = rs.Fields.Item(colName).Value
            Catch ex As Exception
                resu = def_val
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a Recordset string value by the name of the field.
        ''' </summary>        
        Public Shared Function RSetValueStr(rs As SAPbobsCOM.Recordset, colName As String, Optional def_val As String = "") As Object
            Return Tools.RSetValue(rs, colName, def_val).ToString()
        End Function

        ''' <summary>
        ''' Get a Recordset decimal value by the name of the field.
        ''' </summary>        
        Public Shared Function RSetValueDecimal(rs As SAPbobsCOM.Recordset, colName As String, Optional def_val As String = "") As Decimal
            If UCase(CultureInfo.CurrentCulture.Name) = "ES-ES" Then
                Return Tools.ToDecimal(Tools.RSetValue(rs, colName, def_val).ToString()) ' Al hacer ToString() cambia la "," por "." como separador decimal
            Else
                Return Tools.ToDecimal(Tools.RSetValue(rs, colName, def_val).ToString(), sep_decimal_from:=".")
            End If
        End Function

        ''' <summary>
        ''' Get a Recordset integer value by the name of the field.
        ''' </summary>        
        Public Shared Function RSetValueInteger(rs As SAPbobsCOM.Recordset, colName As String, Optional def_val As String = "") As Integer
            Return Tools.ToInteger(Tools.RSetValue(rs, colName, def_val).ToString())
        End Function

        ''' <summary>
        ''' Get a Recordset Date value by the name of the field.
        ''' </summary> 
        Public Shared Function RSetValueDate(rs As SAPbobsCOM.Recordset, colName As String, Optional def_val As Date? = Nothing, Optional format As String = "dd/MM/yyyy") As Date?
            Return Tools.ToDate(Tools.RSetValue(rs, colName, ""), def_val, format)
        End Function

        ''' <summary>
        ''' Checks if a form field exists
        ''' </summary>
        Public Shared Function FormExistItem(f As SAPbouiCOM.Form, itemName As String) As Boolean
            Dim resu As Boolean = False
            Try
                If f.Items.Item(itemName).UniqueID = itemName Then
                    resu = True
                End If
            Catch ex As Exception
                resu = False
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a form field, not it's value. 
        ''' </summary>
        Public Shared Function FormItem(f As SAPbouiCOM.Form, colName As String) As Object
            Dim resu As Object
            Try
                resu = f.Items.Item(colName).Specific
            Catch ex As Exception
                resu = New Object
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a form field value.
        ''' </summary>
        Public Shared Function FormValue(f As SAPbouiCOM.Form, colName As String) As Object
            Dim resu As Object
            Try
                resu = f.Items.Item(colName).Specific.Value
            Catch ex As Exception
                resu = New Object
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Get a form field value as string.
        ''' </summary>        
        Public Shared Function FormValueStr(f As SAPbouiCOM.Form, colName As String) As Object
            Return Tools.FormValue(f, colName).ToString()
        End Function

        ''' <summary>
        ''' Get a form field value as decimal
        ''' </summary>        
        Public Shared Function FormValueDecimal(f As SAPbouiCOM.Form, colName As String) As Decimal
            Return Tools.ToDecimal(Tools.FormValue(f, colName).ToString(), sep_decimal_from:=SEP_DECIMAL_UI)
        End Function

        ''' <summary>
        ''' Get a form field value as integer.
        ''' </summary>        
        Public Shared Function FormValueInteger(f As SAPbouiCOM.Form, colName As String) As Integer
            Return Tools.ToInteger(Tools.FormValue(f, colName).ToString())
        End Function

        ''' <summary>
        ''' Get a form field value as date.
        ''' </summary>        
        Public Shared Function FormValueDate(f As SAPbouiCOM.Form, colName As String, Optional format As String = "yyyyMMdd", Optional def_val As Date? = Nothing) As Date?
            Dim resu As Date? = Tools.ToDate(Tools.FormValueStr(f, colName), format:=format, def_val:=def_val)
            Return resu
            'Return Tools.ToDate(Tools.FormValueStr(f, colName), format:=format)
        End Function

        ''' <summary>
        ''' Get a form field value as dateTime.
        ''' </summary>        
        Public Shared Function FormValueDateTime(f As SAPbouiCOM.Form, colName As String) As DateTime?
            Return Tools.ToDateTime(Tools.FormValueStr(f, colName), format:="yyyyMMdd")
        End Function

        ''' <summary>
        ''' Set a form value. The field must be editable in the form settings.
        ''' </summary>        
        ''' <returns>0: no errors, 1: other cases</returns>
        ''' <remarks></remarks>
        Public Shared Function FormSetValue(f As SAPbouiCOM.Form, colName As String, value As String) As Integer
            Dim resu As Integer = 1
            Try
                f.Items.Item(colName).Specific.Value = value
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function

        ''' <summary>
        ''' Set combobox value in a form. The field must be editable in the form settings.
        ''' </summary>        
        Public Shared Function FormSetCombo(f As SAPbouiCOM.Form, colName As String, value As String) As Integer
            Dim resu As Integer = 1
            Try
                f.Items.Item(colName).Specific.Select(value, SAPbouiCOM.BoSearchKey.psk_ByValue)
            Catch ex As Exception
                resu = 0
            End Try
            Return resu
        End Function

    End Class

End Module