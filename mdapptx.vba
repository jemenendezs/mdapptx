Sub ImportarMarkdownAPowerPoint()
    Dim fd As FileDialog
    Dim archivoSeleccionado As String
    Dim contenidoMarkdown As String
    Dim lineas As Variant
    Dim i As Integer
    Dim presentacion As Presentation
    Dim diapositiva As Slide
    Dim diapositivaActual As Integer
    Dim contadorTitulos As Integer
    Dim contadorVinetas As Integer
    
    ' Crear cuadro de diálogo para seleccionar archivo
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    With fd
        .Title = "Seleccionar archivo Markdown"
        .Filters.Clear
        .Filters.Add "Archivos Markdown", "*.md;*.markdown;*.txt"
        .Filters.Add "Todos los archivos", "*.*"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            archivoSeleccionado = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ningún archivo", vbInformation
            Exit Sub
        End If
    End With
    
    ' Leer contenido del archivo con soporte UTF-8
    contenidoMarkdown = LeerArchivoUTF8(archivoSeleccionado)
    
    ' Verificar que el archivo no esté vacío
    If Len(contenidoMarkdown) = 0 Then
        MsgBox "El archivo seleccionado está vacío", vbWarning
        Exit Sub
    End If
    
    ' Dividir en líneas
    contenidoMarkdown = Replace(contenidoMarkdown, vbCrLf, vbLf)
    contenidoMarkdown = Replace(contenidoMarkdown, vbCr, vbLf)
    lineas = Split(contenidoMarkdown, vbLf)
    
    ' Obtener presentación activa
    Set presentacion = ActivePresentation
    diapositivaActual = 0
    contadorTitulos = 0
    contadorVinetas = 0
    Set diapositiva = Nothing
    
    ' Procesar cada línea
    For i = 0 To UBound(lineas)
        Dim linea As String
        linea = Trim(lineas(i))
        
        ' Saltar líneas vacías
        If Len(linea) > 0 Then
            ' Verificar si es un título
            If EsTitulo(linea) Then
                contadorTitulos = contadorTitulos + 1
                
                ' Crear nueva diapositiva
                Set diapositiva = CrearNuevaDiapositiva(presentacion)
                
                If Not diapositiva Is Nothing Then
                    diapositivaActual = diapositivaActual + 1
                    
                    ' Agregar título usando la nueva función segura
                    Dim titulo As String
                    titulo = ObtenerTitulo(linea)
                    AgregarTitulo diapositiva, titulo
                End If
                
            ' Verificar si es una viñeta
            ElseIf EsVineta(linea) Then
                contadorVinetas = contadorVinetas + 1
                If Not diapositiva Is Nothing Then
                    AgregarVineta diapositiva, ObtenerTextoVineta(linea)
                End If
            End If
        End If
    Next i
    
    ' Mensaje de resultado detallado
    MsgBox "Procesamiento completado:" & vbCrLf & _
           "- Líneas procesadas: " & (UBound(lineas) + 1) & vbCrLf & _
           "- Títulos encontrados: " & contadorTitulos & vbCrLf & _
           "- Viñetas encontradas: " & contadorVinetas & vbCrLf & _
           "- Diapositivas creadas: " & diapositivaActual, vbInformation
End Sub

Function LeerArchivoUTF8(rutaArchivo As String) As String
    Dim stream As Object
    
    On Error GoTo ErrorHandler
    
    ' Crear objeto ADODB.Stream para leer UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Texto
    stream.Charset = "utf-8"
    stream.Open
    stream.LoadFromFile rutaArchivo
    LeerArchivoUTF8 = stream.ReadText
    stream.Close
    Set stream = Nothing
    Exit Function
    
ErrorHandler:
    ' Si falla UTF-8, intentar con el método tradicional
    If Not stream Is Nothing Then
        stream.Close
        Set stream = Nothing
    End If
    
    ' Fallback al método tradicional
    Dim numeroArchivo As Integer
    Dim contenido As String
    
    numeroArchivo = FreeFile
    Open rutaArchivo For Input As #numeroArchivo
    contenido = Input$(LOF(numeroArchivo), numeroArchivo)
    Close #numeroArchivo
    
    LeerArchivoUTF8 = contenido
End Function

Function LimpiarTexto(texto As String) As String
    Dim textoLimpio As String
    textoLimpio = texto
    
    ' Eliminar referencias como [^1], [^2], [^abc], etc.
    Dim i As Integer
    Dim patron As String
    Dim posicion As Integer
    
    ' Buscar y eliminar patrones de referencias [^...]
    Do
        posicion = InStr(textoLimpio, "[^")
        If posicion > 0 Then
            Dim finReferencia As Integer
            finReferencia = InStr(posicion, textoLimpio, "]")
            If finReferencia > 0 Then
                ' Eliminar la referencia completa
                textoLimpio = Left(textoLimpio, posicion - 1) & Mid(textoLimpio, finReferencia + 1)
            Else
                Exit Do
            End If
        End If
    Loop While posicion > 0
    
    ' Eliminar espacios dobles que puedan quedar
    Do While InStr(textoLimpio, "  ") > 0
        textoLimpio = Replace(textoLimpio, "  ", " ")
    Loop
    
    ' Limpiar espacios al inicio y final
    textoLimpio = Trim(textoLimpio)
    
    LimpiarTexto = textoLimpio
End Function

Function AgregarPuntoFinal(texto As String) As String
    ' *** AGREGAR PUNTOS FINALES SOLO A VIÑETAS (NO A TÍTULOS) ***
    ' Para deshabilitar esta función, cambia "True" por "False" en la siguiente línea:
    Dim agregarPuntos As Boolean
    agregarPuntos = True
    
    If Not agregarPuntos Then
        AgregarPuntoFinal = texto
        Exit Function
    End If
    
    Dim textoConPunto As String
    textoConPunto = Trim(texto)
    
    ' Verificar si el texto no está vacío
    If Len(textoConPunto) > 0 Then
        ' Obtener el último carácter
        Dim ultimoCaracter As String
        ultimoCaracter = Right(textoConPunto, 1)
        
        ' Si no termina en punto, signo de interrogación o exclamación, agregar punto
        If ultimoCaracter <> "." And ultimoCaracter <> "?" And ultimoCaracter <> "!" And ultimoCaracter <> ":" Then
            textoConPunto = textoConPunto & "."
        End If
    End If
    
    AgregarPuntoFinal = textoConPunto
End Function

Function EsTitulo(linea As String) As Boolean
    Dim lineaTrimmed As String
    lineaTrimmed = Trim(linea)
    
    If Len(lineaTrimmed) >= 2 Then
        If Left(lineaTrimmed, 2) = "##" Then
            EsTitulo = True
            Exit Function
        End If
        
        If Left(lineaTrimmed, 1) = "#" And Len(lineaTrimmed) > 1 Then
            If Mid(lineaTrimmed, 2, 1) = " " Or Mid(lineaTrimmed, 2, 1) = "#" Then
                EsTitulo = True
                Exit Function
            End If
        End If
    End If
    
    EsTitulo = False
End Function

Function ObtenerTitulo(linea As String) As String
    Dim titulo As String
    Dim lineaTrimmed As String
    lineaTrimmed = Trim(linea)
    
    Dim i As Integer
    For i = 1 To Len(lineaTrimmed)
        If Mid(lineaTrimmed, i, 1) <> "#" Then
            Exit For
        End If
    Next i
    
    titulo = Trim(Mid(lineaTrimmed, i))
    
    ' Limpiar referencias PERO NO agregar punto final a los títulos
    titulo = LimpiarTexto(titulo)
    
    ' If titulo = "" Then
    '     titulo = "Diapositiva " & (ActivePresentation.Slides.Count + 1)
    ' End If
   
    ObtenerTitulo = titulo
End Function

Function EsVineta(linea As String) As Boolean
    Dim lineaTrimmed As String
    lineaTrimmed = LTrim(linea)
    
    If Len(lineaTrimmed) >= 1 Then
        If Left(lineaTrimmed, 1) = "-" Then
            If Len(lineaTrimmed) = 1 Or Mid(lineaTrimmed, 2, 1) = " " Then
                EsVineta = True
            End If
        End If
    End If
End Function

Function ObtenerTextoVineta(linea As String) As String
    Dim lineaTrimmed As String
    Dim textoVineta As String
    
    lineaTrimmed = LTrim(linea)
    
    If Len(lineaTrimmed) > 1 Then
        textoVineta = Trim(Mid(lineaTrimmed, 2))
    Else
        textoVineta = ""
    End If
    
    ' Limpiar referencias Y agregar punto final SOLO a las viñetas
    textoVineta = LimpiarTexto(textoVineta)
    textoVineta = AgregarPuntoFinal(textoVineta)
    
    If textoVineta = "" Or textoVineta = "." Then
        textoVineta = "Elemento de lista."
    End If
    
    ObtenerTextoVineta = textoVineta
End Function

Function CrearNuevaDiapositiva(pres As Presentation) As Slide
    Dim nuevaDisp As Slide
    
    On Error GoTo ErrorHandler
    
    ' Crear nueva diapositiva con diseño de título y contenido (valor 2)
    Set nuevaDisp = pres.Slides.Add(pres.Slides.Count + 1, 2)
    
    ' Verificar que la diapositiva se creó correctamente
    If nuevaDisp Is Nothing Then
        GoTo ErrorHandler
    End If
    
    Set CrearNuevaDiapositiva = nuevaDisp
    Exit Function
    
ErrorHandler:
    MsgBox "Error al crear nueva diapositiva: " & Err.Description, vbCritical
    Set CrearNuevaDiapositiva = Nothing
End Function

Sub AgregarTitulo(diap As Slide, textoTitulo As String)
    Dim formaTitulo As Shape
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Buscar la forma de título
    For i = 1 To diap.Shapes.Count
        Set formaTitulo = diap.Shapes(i)
        
        ' Verificar si la forma tiene TextFrame y es un marcador de posición de título
        If formaTitulo.HasTextFrame Then
            If formaTitulo.TextFrame.HasText Or formaTitulo.Type = 14 Then ' 14 = msoPlaceholder
                ' Verificar si es un título (generalmente la primera forma o con nombre específico)
                If i = 1 Or InStr(LCase(formaTitulo.Name), "title") > 0 Then
                    formaTitulo.TextFrame.TextRange.Text = textoTitulo
                    Exit Sub
                End If
            End If
        End If
    Next i
    
    ' Si no se encontró forma de título, usar la primera forma disponible
    If diap.Shapes.Count > 0 Then
        Set formaTitulo = diap.Shapes(1)
        If formaTitulo.HasTextFrame Then
            formaTitulo.TextFrame.TextRange.Text = textoTitulo
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al agregar título: " & Err.Description & vbCrLf & "Título: " & textoTitulo, vbWarning
End Sub

Sub AgregarVineta(diap As Slide, textoVineta As String)
    Dim formaContenido As Shape
    Dim rangoTexto As TextRange
    Dim i As Integer
    
    On Error GoTo ErrorHandler
    
    ' Buscar la forma de contenido (generalmente la segunda forma o con "content" en el nombre)
    For i = 1 To diap.Shapes.Count
        Set formaContenido = diap.Shapes(i)
        
        If formaContenido.HasTextFrame Then
            ' Buscar forma de contenido (no título)
            If i > 1 Or InStr(LCase(formaContenido.Name), "content") > 0 Or InStr(LCase(formaContenido.Name), "text") > 0 Then
                Set rangoTexto = formaContenido.TextFrame.TextRange
                
                ' Agregar el texto de la viñeta
                If Len(Trim(rangoTexto.Text)) > 0 Then
                    rangoTexto.Text = rangoTexto.Text & vbCrLf & textoVineta
                Else
                    rangoTexto.Text = textoVineta
                End If
                
                ' Configurar formato de viñetas
                With rangoTexto.ParagraphFormat
                    .Bullet.Visible = -1  ' msoTrue = -1
                    .Bullet.Type = 1      ' ppBulletUnnumbered = 1
                End With
                
                Exit Sub
            End If
        End If
    Next i
    
    ' Si no se encontró forma de contenido, usar la última forma disponible
    If diap.Shapes.Count > 1 Then
        Set formaContenido = diap.Shapes(diap.Shapes.Count)
        If formaContenido.HasTextFrame Then
            Set rangoTexto = formaContenido.TextFrame.TextRange
            If Len(Trim(rangoTexto.Text)) > 0 Then
                rangoTexto.Text = rangoTexto.Text & vbCrLf & textoVineta
            Else
                rangoTexto.Text = textoVineta
            End If
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al agregar viñeta: " & Err.Description & vbCrLf & "Viñeta: " & textoVineta, vbWarning
End Sub
