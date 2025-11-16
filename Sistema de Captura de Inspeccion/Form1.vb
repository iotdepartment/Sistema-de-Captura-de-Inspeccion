Imports System.Data.SqlClient
Imports System.Runtime.InteropServices
Public Class Form1
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        SetForegroundWindow(Me.Handle)
        If TextBox1.Enabled Then
            TextBox1.Focus()
        ElseIf TextBox2.Enabled Then
            TextBox2.Focus()
        End If
        'Coment Test 
    End Sub
    Dim cadenaConexion As String = "Server=RMX-D4LZZV2;Database=ScanSystemDB;User Id=Manu ;Password=2022.Tgram2;"
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Mayuscul()
        Limpiar()
    End Sub
    Private Sub Enter2(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyData = Keys.Enter Then
            ' Obtener el texto del TextBox2
            Dim inputText As String = TextBox2.Text
            ' Verificar si el texto ingresado es numérico
            If IsNumeric(inputText) Then
                ' Llamar a la función numemp (si la tienes definida)
                numemp()
                Using connection As New SqlConnection(cadenaConexion)
                    connection.Open()
                    ' Usar corchetes [] para escapar la palabra reservada 'User'
                    Dim query As String = "SELECT COUNT(*) FROM [User] WHERE NumerodeEmpleado = @NumerodeEmpleado"
                    Using cmd As New SqlCommand(query, connection)
                        Dim Busqueda As String
                        Busqueda = TextBox2.Text
                        cmd.Parameters.AddWithValue("@NumerodeEmpleado", Busqueda)
                        Dim result As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                        If result > 0 Then
                            ' Si el número de empleado existe, mostrar el número de empleado en Label3
                            Label3.Text = TextBox2.Text
                            TextBox2.Text = "" ' Limpiar el TextBox2
                            TextBox2.Select() ' Volver a seleccionar TextBox2 para la siguiente entrada

                            ' Obtener el nombre del empleado y asignarlo a Label4
                            Dim NombreQuery As String = "SELECT Nombre FROM [User] WHERE NumerodeEmpleado = @NumerodeEmpleado"
                            Using NombreCmd As New SqlCommand(NombreQuery, connection)
                                NombreCmd.Parameters.AddWithValue("@NumerodeEmpleado", Busqueda)
                                Dim Nombre As String = Convert.ToString(NombreCmd.ExecuteScalar())
                                Label4.Text = Nombre ' Coloca el nombre en Label4
                            End Using

                            ' Verificar si la mesa ya está ingresada
                            If Label1.Text = "Mesa:" Or Label1.Text = "" Then
                                ' Si no se ha ingresado la mesa, pedir al usuario que ingrese una
                                Label5.BackColor = Color.Orange
                                Label5.ForeColor = Color.Black
                                Label5.Text = "INGRESE LA MESA"
                            Else
                                ' Si la mesa está ingresada, deshabilitar TextBox2 y habilitar TextBox1
                                TextBox2.Enabled = False
                                TextBox1.Enabled = True
                                TextBox1.Select() ' Seleccionar TextBox1 para ingresar la siguiente información
                                Label5.Text = "INGRESA EL MANDRIL A INSPECCIONAR"
                            End If
                        Else
                            ' Si el número de empleado no existe en la base de datos
                            Label5.Text = "Usted ha ingresado un número" & vbCrLf & "que no se encuentra en la Base de datos"
                            Label5.BackColor = Color.DarkRed
                            Label5.ForeColor = Color.White
                        End If
                    End Using
                End Using
            Else
                ' Si el texto no es un número, verificar si es una mesa válida
                Using connection As New SqlConnection(cadenaConexion)
                    connection.Open()
                    Dim query As String = "SELECT COUNT(*) FROM Mesas WHERE Mesas = @Mesas"
                    Using cmd As New SqlCommand(query, connection)
                        cmd.Parameters.AddWithValue("@Mesas", TextBox2.Text)
                        Dim result As Integer = Convert.ToInt32(cmd.ExecuteScalar())

                        If result > 0 Then
                            ' Si la mesa existe, mostrarla en Label1
                            Label1.Text = TextBox2.Text
                            TextBox2.Text = "" ' Limpiar el TextBox2
                            TextBox2.Select() ' Volver a seleccionar TextBox2 para la siguiente entrada

                            ' Verificar si el nombre ya está ingresado
                            If Label4.Text = "Nombre:" Or Label4.Text = "" Then
                                ' Si no se ha ingresado el nombre, pedir al usuario que ingrese el número de empleado
                                Label5.Text = "INGRESE EL NUMERO DE EMPLEADO"
                            Else
                                ' Si el nombre está ingresado, deshabilitar TextBox2 y habilitar TextBox1
                                TextBox2.Enabled = False
                                TextBox1.Enabled = True
                                TextBox1.Select() ' Seleccionar TextBox1 para ingresar la siguiente información
                                Label5.Text = "INGRESA EL MANDRIL A INSPECCIONAR"
                            End If
                        Else
                            ' Si la mesa no existe en la base de datos
                            Label5.Text = "Usted ha ingresado una mesa" & vbCrLf & "que no se encuentra en la Base de datos"
                            Label5.BackColor = Color.DarkRed
                            Label5.ForeColor = Color.White
                        End If
                    End Using
                End Using
            End If
            ' Limpiar el TextBox2 al final de la validación
            TextBox2.Text = ""
            ' Verificar si la mesa no se ha ingresado correctamente
            If Label1.Text = "" Then
                ' Mostrar mensaje de error si la mesa es incorrecta
                Label5.Text = "Usted ha ingresado una mesa INCORRECTA!!!"
                Label5.BackColor = Color.DarkRed
                Label5.ForeColor = Color.White
            End If
            ' Verificar si el número de empleado no se ha ingresado correctamente
            If Label3.Text = "" Then
                ' Mostrar mensaje de error si el número de empleado no está en la base de datos
                Label5.Text = "Usted ha ingresado un número" & vbCrLf & "que no se encuentra en la Base de datos"
                Label5.BackColor = Color.DarkRed
                Label5.ForeColor = Color.White
            End If
        End If
    End Sub
    Sub Insertardefecto()
        Using connection As New SqlConnection(cadenaConexion)
            If Label2.Text = "Mandril" OrElse Label2.Text = "Mandril no encontrado" Then
            Else
                connection.Open()
                GetDefectoData(connection)
                GetCountDefectosPorFechaYTurno()
                CargarRegistroDeDefectos()
            End If
        End Using
    End Sub
    Sub BuscarMandrel()
        Using connection As New SqlConnection(cadenaConexion)
            connection.Open()
            If Label2.Text = "Mandril" OrElse Label2.Text = "Mandril no encontrado" Then
                GetMandrelData(connection)
                GetTotalPiezasPorFechaYTurnoConNuMesa()
                GetTotalPiezasPorFechaYTurnoConMandrelYNuMesa()

                GetCountDefectosPorFechaYTurno()
                CargarRegistroDeDefectos()
                CargarDatosConFiltros()
            Else
                If TextBox1.Text = Label57.Text Then
                    ' Si el texto coincide con Label57, insertar en RegistrodePiezasEscaneadas
                    InsertRecord(connection)
                    GetTotalPiezasPorFechaYTurnoConNuMesa()
                    GetTotalPiezasPorFechaYTurnoConMandrelYNuMesa()

                    GetCountDefectosPorFechaYTurno()
                    CargarRegistroDeDefectos()
                    CargarDatosConFiltros()
                Else
                    ' Obtener datos de Mandrels si no coincide
                    GetMandrelData(connection)
                    GetTotalPiezasPorFechaYTurnoConNuMesa()
                    GetTotalPiezasPorFechaYTurnoConMandrelYNuMesa()

                    GetCountDefectosPorFechaYTurno()
                    CargarRegistroDeDefectos()
                    CargarDatosConFiltros()
                End If
            End If
        End Using
    End Sub
    Sub InsertarParciales()
        ' Operaciones cuando el texto comienza con "+"
        Using connection As New SqlConnection(cadenaConexion)
            ' Extraer el número después del "+"
            connection.Open()
            Dim texto As String = TextBox1.Text
            Dim numeroAdicional As Integer
            If Label2.Text = "Mandril" Then
            Else
                If Integer.TryParse(texto.Substring(1), numeroAdicional) Then
                    ' Si se pudo convertir el texto en un número, guardarlo en la variable
                    'MessageBox.Show($"Número adicional: {numeroAdicional}", "Número extraído", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    ' Aquí puedes realizar las operaciones necesarias con el número
                    'ProcessNumberAddition(numeroAdicional)
                    NumerodeParciales = numeroAdicional
                    InsertRecordParcial(connection)
                    GetTotalPiezasPorFechaYTurnoConNuMesa()
                    GetTotalPiezasPorFechaYTurnoConMandrelYNuMesa()
                    CargarDatosConFiltros()
                Else
                    'MessageBox.Show("El número después del '+' no es válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End If

            End If

        End Using
    End Sub
    Private UltimoEscaneo As DateTime = DateTime.MinValue
    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown

        If e.KeyData = Keys.Enter Then
            ' Dim ahora As DateTime = DateTime.Now ' Esta línea ya no se necesita aquí
            ' If (ahora - UltimoEscaneo).TotalSeconds < 3 Then ' Esta validación se mueve abajo

            ' ... otras comprobaciones (CAMBIODEMODELO, números, etc.) no tienen retardo
            Try
                If String.Equals(TextBox1.Text, "CAMBIODEMODELO", StringComparison.OrdinalIgnoreCase) Then
                    ManejarCambioDeModelo()
                ElseIf String.Equals(TextBox1.Text, "CAMBIODEMODELO01", StringComparison.OrdinalIgnoreCase) Then
                    ManejarCambioDeModelo01()
                ElseIf Not String.IsNullOrEmpty(TextBox1.Text) AndAlso Char.IsDigit(TextBox1.Text(0)) Then
                    ' Operaciones cuando el texto comienza con un número
                    Insertardefecto()
                ElseIf TextBox1.Text.StartsWith("F", StringComparison.OrdinalIgnoreCase) Then
                    ' **AQUÍ EMPIEZA LA VALIDACIÓN DE TIEMPO SOLO PARA "F"**
                    Dim ahora As DateTime = DateTime.Now
                    If (ahora - UltimoEscaneo).TotalSeconds < 3 Then
                        Label5.Text = "Espere 3 segundos antes de escanear nuevamente"
                        Label5.BackColor = Color.Yellow
                        Label5.ForeColor = Color.Black
                        TextBox1.Text = ""
                        Exit Sub ' Sale del sub porque la validación falló
                    End If
                    UltimoEscaneo = ahora
                    ' **AQUÍ TERMINA LA VALIDACIÓN DE TIEMPO SOLO PARA "F"**

                    ' Operaciones cuando el texto comienza con "F"
                    BuscarMandrel() ' Solo se ejecuta si pasaron 3 segundos
                ElseIf TextBox1.Text.StartsWith("+", StringComparison.OrdinalIgnoreCase) Then
                    InsertarParciales()
                ElseIf TextBox1.Text.StartsWith("P", StringComparison.OrdinalIgnoreCase) Then
                    EliminarUltimoRegistroConFiltros()
                ElseIf TextBox1.Text.StartsWith("/", StringComparison.OrdinalIgnoreCase) Then
                    InsertDownTime(Label4.Text, TextBox1.Text, Label1.Text)

                End If
            Catch ex As Exception
                ' Manejo de errores
                'MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
            TextBox1.Text = ""
        End If

    End Sub
    Private Sub CargarDatosConFiltros()
        Dim query As String = "
    SELECT 
        [Mandrel],
        SUM(CAST([NDPiezas] AS INT)) AS TotalPiezas
    FROM 
        [ScanSystemDB].[dbo].[RegistrodePiezasEscaneadas]
    WHERE 
        [Fecha] = @Fecha AND
        [Turno] = @Turno AND
        [NuMesa] = @NuMesa AND
        [TM] = @TM
    GROUP BY 
            [Mandrel];"
        Using conn As New SqlConnection(cadenaConexion)
            Using cmd As New SqlCommand(query, conn)
                ' Agregar parámetros desde los controles del formulario
                cmd.Parameters.AddWithValue("@Fecha", DateTime.Now.ToString("yyyy-MM-dd"))
                cmd.Parameters.AddWithValue("@Turno", LabelTurno.Text)
                cmd.Parameters.AddWithValue("@NuMesa", Label1.Text)
                cmd.Parameters.AddWithValue("@TM", Label4.Text)
                Dim adapter As New SqlDataAdapter(cmd)
                Dim dt As New DataTable()
                Try
                    conn.Open()
                    adapter.Fill(dt)
                    DataGridView1.DataSource = dt
                Catch ex As Exception
                    ' MessageBox.Show("Error al cargar datos: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub
    ' Metodo para agregar ceros a la izquierda si el texto en TextBox2 es un número
    Sub numemp()
        Dim texto As String = TextBox2.Text
        ' Verifica si el texto no está vacío y si el texto es un dígito
        If Not String.IsNullOrEmpty(texto) AndAlso Char.IsDigit(texto(0)) Then
            TextBox2.Text = texto.PadLeft(6, "0"c)
        End If
    End Sub
    ' Metodo para convertir el texto en mayúsculas en los TextBox
    Sub Mayuscul()
        TextBox1.CharacterCasing = CharacterCasing.Upper
        TextBox2.CharacterCasing = CharacterCasing.Upper
    End Sub
    ' Metodo para limpiar los Label
    Sub Limpiar()
        Label1.Text = "Mesa:"
        Label4.Text = "Nombre:"
        Label3.Text = "Numero de Empleado"
        Label2.Text = "Mandril"
        Label13.Text = ""
    End Sub
    Private Sub HorayFecha_Tick(sender As Object, e As EventArgs) Handles HorayFecha.Tick
        ' Actualiza la hora y la fecha en los controles
        FHORA.Text = DateTime.Now.ToLongTimeString
        FFecha.Text = DateTime.Now.ToString("yyyy-MM-dd")

        ' Obtener la hora actual
        Dim currentTime As DateTime = DateTime.Now

        ' Definir los rangos de tiempo para los turnos
        Dim morningStartTime As DateTime = DateTime.Today.AddHours(7.1667) ' 7:10 AM
        Dim afternoonEndTime As DateTime = DateTime.Today.AddHours(15.75) ' 3:40 PM
        Dim eveningEndTime As DateTime = DateTime.Today.AddHours(23.99) ' 11:50 PM (medianoche)
        Dim nightStartTime As DateTime = DateTime.Today.AddHours(23.999) ' 11:50 PM (medianoche)
        Dim nightEndTime As DateTime = DateTime.Today.AddDays(1).AddHours(7.1666) ' 7:00 AM del día siguiente

        ' Verificar en qué rango de tiempo está la hora actual
        If currentTime >= morningStartTime AndAlso currentTime < afternoonEndTime Then
            LabelTurno.Text = "1" ' Primer turno (7:00 AM - 3:30 PM)
        ElseIf currentTime >= afternoonEndTime AndAlso currentTime < eveningEndTime Then
            LabelTurno.Text = "2" ' Segundo turno (3:30 PM - 11:50 PM)
        ElseIf currentTime >= nightStartTime AndAlso currentTime < nightEndTime Then
            LabelTurno.Text = "3" ' Tercer turno (11:50 PM - 7:00 AM del día siguiente)
        Else
            LabelTurno.Text = "3" ' Tercer turno (11:50 PM - 7:00 AM del día siguiente)
        End If
    End Sub
    Dim NumerodeParciales As String
    ' Sub para realizar el INSERT en la tabla RegistrodePiezasEscaneadas
    Private Sub InsertRecord(connection As SqlConnection)

        Dim insertQuery As String = "
        INSERT INTO [dbo].[RegistrodePiezasEscaneadas] 
        (Mandrel, NDPiezas, Turno, NuMesa, TM)
        VALUES (@Mandrel, @NDPiezas, @Turno, @NuMesa, @TM)"

        Using insertCommand As New SqlCommand(insertQuery, connection)
            ' Parámetros para el INSERT
            insertCommand.Parameters.AddWithValue("@Mandrel", Label2.Text) ' Mandril obtenido
            insertCommand.Parameters.AddWithValue("@NDPiezas", Label13.Text) ' Ejemplo, puedes ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@Turno", LabelTurno.Text) ' Ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@NuMesa", Label1.Text) ' Ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@TM", Label4.Text) ' Ajustar según sea necesario

            ' Ejecutar el INSERT
            insertCommand.ExecuteNonQuery()

            '' Mostrar mensaje de éxito
            'MessageBox.Show("Registro insertado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
        connection.Close()
    End Sub
    ' Sub para realizar el INSERT de parciales en la tabla RegistrodePiezasEscaneadas
    Private Sub InsertRecordParcial(connection As SqlConnection)
        Dim insertQuery As String = "
        INSERT INTO [dbo].[RegistrodePiezasEscaneadas] 
        (Mandrel, NDPiezas, Turno, NuMesa, TM)
        VALUES (@Mandrel, @NDPiezas, @Turno, @NuMesa, @TM)"

        Using insertCommand As New SqlCommand(insertQuery, connection)
            ' Parámetros para el INSERT
            insertCommand.Parameters.AddWithValue("@Mandrel", Label2.Text) ' Mandril obtenido
            insertCommand.Parameters.AddWithValue("@NDPiezas", NumerodeParciales) ' Ejemplo, puedes ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@Turno", LabelTurno.Text) ' Ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@NuMesa", Label1.Text) ' Ajustar según sea necesario
            insertCommand.Parameters.AddWithValue("@TM", Label4.Text) ' Ajustar según sea necesario

            ' Ejecutar el INSERT
            insertCommand.ExecuteNonQuery()

            '' Mostrar mensaje de éxito
            'MessageBox.Show("Registro insertado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End Using
        connection.Close()
    End Sub
    ' Sub para obtener los datos de Mandrel, BarCode y CANTIDADEMPAQUEINSPECCION desde Mandrels
    Private Sub GetMandrelData(connection As SqlConnection)
        ' Consulta SQL para obtener datos de Mandrels
        Dim query As String = "
        SELECT Mandril, BarCode, CantidaddeEmpaque 
        FROM [dbo].[Mandriles] 
        WHERE BarCode = @BarCode"

        Using command As New SqlCommand(query, connection)
            command.Parameters.AddWithValue("@BarCode", TextBox1.Text)

            ' Usar un bloque Using para manejar el SqlDataReader
            Using reader As SqlDataReader = command.ExecuteReader()
                If reader.Read() Then
                    ' Asignar valores a los labels
                    Label2.Text = reader("Mandril").ToString()
                    Label57.Text = reader("BarCode").ToString()
                    Label13.Text = reader("CantidaddeEmpaque").ToString()
                Else
                    ' Manejo cuando no hay resultados
                    Label2.Text = "Mandril no encontrado"
                    Label57.Text = ""
                    Label5.Text = "NO SE ENCONTRO EL MANDRIL INTENTE DE NUEVO"
                    Label13.Text = ""
                End If
            End Using ' Esto asegura que el reader se cierra automáticamente
        End Using
    End Sub
    ' Sub para obtener los datos de Defecto, desde Defectos
    Private Sub GetDefectoData(connection As SqlConnection)
        Try
            Dim query As String = "
        SELECT Defecto, CodigodeDefecto
        FROM [dbo].[Defectos] 
        WHERE CodigodeDefecto = @CodigodeDefecto"

            Using command As New SqlCommand(query, connection)
                ' Parámetros de consulta
                command.Parameters.Add("@CodigodeDefecto", SqlDbType.NVarChar).Value = TextBox1.Text.Trim()


                ' Ejecutar la consulta y obtener el resultado
                Using reader As SqlDataReader = command.ExecuteReader()
                    If reader.Read() Then
                        ' Asignar Defecto, CodigodeDefecto, Fecha, Turno, NuMesa y Mandrel a las etiquetas correspondientes
                        Label18.Text = reader("Defecto").ToString()
                        Label19.Text = reader("CodigodeDefecto").ToString()
                    Else
                        ' Si no se encuentra el registro, mostrar un mensaje adecuado
                        Label18.Text = "Defecto no encontrado"
                        Label19.Text = ""
                    End If
                End Using ' Aquí el DataReader se cierra antes de realizar cualquier otra operación
            End Using

            ' Ahora que el DataReader está cerrado, puedes insertar el registro
            If Not String.IsNullOrEmpty(Label18.Text) AndAlso Label18.Text <> "Defecto no encontrado" Then
                InsertDefectoRecord(connection)
            End If

        Catch ex As Exception
            ' Manejar errores y mostrar mensaje si ocurre un problema
            '  MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Sub para insertar el registro en la tabla RegistrodeDefectos
    Private Sub InsertDefectoRecord(connection As SqlConnection)
        Try
            ' Consulta SQL para insertar el registro en RegistrodeDefectos
            Dim insertQuery As String = "
            INSERT INTO [dbo].[RegistrodeDefectos] 
            (Mandrel, CodigodeDefecto, Defecto, NuMesa, Turno, TM)
            VALUES (@Mandrel, @CodigodeDefecto, @Defecto, @NuMesa, @Turno, @TM)"
            Using insertCommand As New SqlCommand(insertQuery, connection)
                ' Parámetros para el INSERT
                insertCommand.Parameters.AddWithValue("@Mandrel", Label2.Text) ' Mandril obtenido
                insertCommand.Parameters.AddWithValue("@CodigodeDefecto", Label19.Text) ' Código de defecto obtenido
                insertCommand.Parameters.AddWithValue("@Defecto", Label18.Text) ' Defecto obtenido
                insertCommand.Parameters.AddWithValue("@NuMesa", Label1.Text) ' Número de mesa
                insertCommand.Parameters.AddWithValue("@Turno", LabelTurno.Text) ' Turno
                insertCommand.Parameters.AddWithValue("@TM", Label4.Text) ' TM
                ' Ejecutar el INSERT
                insertCommand.ExecuteNonQuery()
                '' Mostrar mensaje de éxito
            End Using
        Catch ex As Exception
        End Try
    End Sub

    Private Sub GetTotalPiezasPorFechaYTurnoConNuMesa()
        Try
            ' Consulta SQL con parámetros, incluyendo el filtro de NuMesa
            Dim query As String = "
        SELECT SUM(TRY_CAST(NDPiezas AS INT)) AS TotalPiezas
        FROM [ScanSystemDB].[dbo].[RegistrodePiezasEscaneadas]
        WHERE CONVERT(DATE, Fecha) = @Fecha 
          AND Turno = @Turno 
          AND TM = @TM"
            ' Crear y abrir la conexión
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()
                ' Crear el comando SQL
                Using command As New SqlCommand(query, connection)
                    ' Agregar parámetros
                    command.Parameters.Add("@Fecha", SqlDbType.Date).Value = DateTime.Now.Date
                    command.Parameters.Add("@Turno", SqlDbType.NVarChar).Value = LabelTurno.Text.Trim()
                    command.Parameters.Add("@TM", SqlDbType.NVarChar).Value = Label4.Text.Trim() ' Agregar filtro para NuMesa
                    ' Ejecutar la consulta
                    Dim result As Object = command.ExecuteScalar()
                    ' Mostrar el resultado en Label16
                    If result IsNot DBNull.Value Then
                        Label16.Text = Convert.ToInt32(result).ToString()
                        Label5.ForeColor = Color.Green ' Cambiar a verde si funciona
                    Else
                        Label16.Text = "0"
                        Label5.ForeColor = Color.Red ' Cambiar a rojo si no hay datos
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' Cambiar a rojo si ocurre un error
            Label16.Text = "0"
            Label5.ForeColor = Color.Red
        End Try
    End Sub
    Private Sub GetTotalPiezasPorFechaYTurnoConMandrelYNuMesa()
        Try
            ' Consulta SQL con parámetros, incluyendo los filtros de Mandrel y NuMesa
            Dim query As String = "
        SELECT SUM(TRY_CAST(NDPiezas AS INT)) AS TotalPiezas
        FROM [ScanSystemDB].[dbo].[RegistrodePiezasEscaneadas]
        WHERE CONVERT(DATE, Fecha) = @Fecha 
          AND Turno = @Turno 
          AND NuMesa = @NuMesa
          AND Mandrel = @Mandrel"
            ' Crear y abrir la conexión
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()
                ' Crear el comando SQL
                Using command As New SqlCommand(query, connection)
                    ' Agregar parámetros
                    command.Parameters.Add("@Fecha", SqlDbType.Date).Value = DateTime.Now.Date
                    command.Parameters.Add("@Turno", SqlDbType.NVarChar).Value = LabelTurno.Text.Trim()
                    command.Parameters.Add("@NuMesa", SqlDbType.NVarChar).Value = Label1.Text.Trim() ' Filtro para NuMesa
                    command.Parameters.Add("@Mandrel", SqlDbType.NVarChar).Value = Label2.Text.Trim() ' Filtro para Mandrel
                    ' Ejecutar la consulta
                    Dim result As Object = command.ExecuteScalar()

                    Label15.Text = If(result IsNot Nothing, result.ToString(), "0")
                End Using
            End Using
        Catch ex As Exception
            ' Cambiar a rojo si ocurre un error
            Label16.Text = "0"
            Label5.ForeColor = Color.Red
        End Try
    End Sub
    ' Contar todos los defectos del turno
    Private Sub GetCountDefectosPorFechaYTurno()
        Try
            ' Definir la consulta SQL con filtro por mesa
            Dim query As String = "
        SELECT COUNT(*) 
        FROM [ScanSystemDB].[dbo].[RegistrodeDefectos]
        WHERE CONVERT(DATE, Fecha) = @Fecha 
          AND Turno = @Turno
          AND NuMesa = @NuMesa"

            ' Crear una conexión a la base de datos
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()
                ' Crear el comando SQL
                Using command As New SqlCommand(query, connection)
                    ' Parámetros para la consulta
                    command.Parameters.AddWithValue("@Fecha", DateTime.Now.Date) ' Fecha actual
                    command.Parameters.AddWithValue("@Turno", LabelTurno.Text)    ' Turno de la etiqueta
                    command.Parameters.AddWithValue("@NuMesa", Label1.Text)  ' Mesa desde un TextBox
                    ' Ejecutar la consulta y obtener el resultado
                    Dim result As Object = command.ExecuteScalar()
                    ' Mostrar el resultado
                    If result IsNot DBNull.Value Then
                        Dim countDefectos As Integer = Convert.ToInt32(result)
                        Label14.Text = countDefectos.ToString()
                    Else
                        ' No hay registros
                    End If
                End Using
            End Using
        Catch ex As Exception
            ' MessageBox.Show($"Ocurrió un error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub
    ' Colocar los defectos en el datagridview
    Private Sub CargarRegistroDeDefectos()
        Dim fecha As String = DateTime.Today.ToString("yyyy-MM-dd")
        Dim turno As String = LabelTurno.Text.Trim()
        Dim nuMesa As String = Label1.Text.Trim()

        Try
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()

                ' 1. Obtener lista dinámica de mandriles
                Dim mandriles As New List(Of String)
                Dim queryMandriles As String = $"
                SELECT DISTINCT QUOTENAME(Mandrel) AS Mandril
                FROM RegistrodeDefectos
                WHERE Fecha = '{fecha}' AND Turno = '{turno}' AND NuMesa = '{nuMesa}'"

                Using cmdMandriles As New SqlCommand(queryMandriles, connection)
                    Using reader As SqlDataReader = cmdMandriles.ExecuteReader()
                        While reader.Read()
                            mandriles.Add(reader("Mandril").ToString())
                        End While
                    End Using
                End Using

                If mandriles.Count = 0 Then

                    Exit Sub
                End If

                Dim columnList As String = String.Join(",", mandriles)

                ' 2. Construir consulta dinámica tipo PIVOT
                Dim sqlPivot As String = $"
                SELECT Defecto, {columnList}
                FROM (
                    SELECT Defecto, Mandrel
                    FROM RegistrodeDefectos
                    WHERE Fecha = '{fecha}' AND Turno = '{turno}' AND NuMesa = '{nuMesa}'
                ) AS SourceTable
                PIVOT (
                    COUNT(Mandrel) FOR Mandrel IN ({columnList})
                ) AS PivotTable
                ORDER BY Defecto"

                ' 3. Ejecutar y mostrar en DataGridView
                Using cmdPivot As New SqlCommand(sqlPivot, connection)
                    Using adapter As New SqlDataAdapter(cmdPivot)
                        Dim dt As New DataTable()
                        adapter.Fill(dt)
                        RegistrodeDefectosDataGridView.DataSource = dt
                    End Using
                End Using

            End Using
        Catch ex As Exception
            '  MessageBox.Show($"Error al cargar la tabla cruzada: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
    ' Cambio de Modelo y Limpiar Mesa
    Private Sub ManejarCambioDeModelo()
        TextBox2.Enabled = True
        Label1.Text = "Mesa:"
        Label3.Text = "Numero de Empleado:"
        Label2.Text = "Mandril"
        TextBox1.Text = ""
        Label5.Text = "INGRESE NUMERO DE EMPLEADO"
        Me.Label5.Font = New Font("Century Gothic", 26.25!, FontStyle.Bold)
        Label4.Text = "Nombre:"
        TextBox2.Select()
        TextBox1.Enabled = False

    End Sub
    Private Sub ManejarCambioDeModelo01()
        Label2.Text = "Mandril"
        TextBox1.Text = ""
        Label5.Text = "INGRESA EL MANDRIL A INSPECCIONAR"
        Me.Label5.Font = New System.Drawing.Font("Century Gothic", 26.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        TextBox1.Select()
    End Sub

    Private Sub ActualizarDatos(connection As SqlConnection)
        Try
            ' Contar defectos
            Dim queryContarDefectos As String = "
            SELECT COUNT(*) 
            FROM RegistrodeDefectos 
            WHERE Fecha = @Fecha AND Turno = @Turno AND Mandril = @Mandril"
            Dim command As New SqlCommand(queryContarDefectos, connection)
            command.Parameters.AddWithValue("@Fecha", FFecha.Text)
            command.Parameters.AddWithValue("@Turno", LabelTurno.Text)
            command.Parameters.AddWithValue("@Mandril", Label2.Text)
            Label14.Text = command.ExecuteScalar().ToString()
            ' Sumar piezas
            Dim querySumarPiezas As String = "
            SELECT ISNULL(SUM(Piezas), 0) 
            FROM RegistrodePiezasEscaneadas 
            WHERE Mandril = @Mandril AND Fecha = @Fecha AND Turno = @Turno"
            command = New SqlCommand(querySumarPiezas, connection)
            command.Parameters.AddWithValue("@Mandril", Label2.Text)
            command.Parameters.AddWithValue("@Fecha", FFecha.Text)
            command.Parameters.AddWithValue("@Turno", LabelTurno.Text)
            ' Total de piezas
            Dim queryTotalPiezas As String = "
            SELECT ISNULL(SUM(Piezas), 0) 
            FROM RegistrodePiezasEscaneadas 
            WHERE Fecha = @Fecha AND Turno = @Turno"
            command = New SqlCommand(queryTotalPiezas, connection)
            command.Parameters.AddWithValue("@Fecha", FFecha.Text)
            command.Parameters.AddWithValue("@Turno", LabelTurno.Text)
            Label16.Text = command.ExecuteScalar().ToString()
        Catch ex As Exception
            '  MessageBox.Show("Error al actualizar los datos: " & ex.Message)
        End Try
    End Sub
    Private Sub ActualizarDatos()
        Try
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()
                Dim query As String = "SELECT SUM(Resultado) FROM RegistrodePiezasEscaneadas WHERE Fecha = @Fecha AND Turno = @Turno AND Tipo = @Tipo"
                Using command As New SqlCommand(query, connection)
                    command.Parameters.AddWithValue("@Fecha", FFecha.Text)
                    command.Parameters.AddWithValue("@Turno", LabelTurno.Text)
                    command.Parameters.AddWithValue("@Tipo", Label2.Text)
                    Dim resultado = command.ExecuteScalar()
                    'Label15.Text = If(resultado IsNot Nothing, resultado.ToString(), "0")
                End Using
            End Using
        Catch ex As Exception

        End Try
    End Sub
    'Eliminar Defectos
    Private Sub EliminarUltimoRegistroConFiltros()
        Dim queryEliminar As String = "
        ;WITH UltimoRegistro AS (
            SELECT TOP 1 Id
            FROM [ScanSystemDB].[dbo].[RegistrodeDefectos]
            WHERE Fecha = CAST(GETDATE() AS DATE)
              AND NuMesa = @NuMesa
            ORDER BY Hora DESC, Id DESC
        )
        DELETE FROM [ScanSystemDB].[dbo].[RegistrodeDefectos]
        WHERE Id IN (SELECT Id FROM UltimoRegistro);
    "

        Try
            Using connection As New SqlConnection(cadenaConexion)
                connection.Open()

                Using commandEliminar As New SqlCommand(queryEliminar, connection)
                    commandEliminar.Parameters.Add("@NuMesa", SqlDbType.NVarChar).Value = Label1.Text.Trim()

                    Dim rowsAffected As Integer = commandEliminar.ExecuteNonQuery()
                    If rowsAffected > 0 Then
                        ' Eliminación exitosa (puedes agregar logging si lo deseas)
                    Else
                        ' No se encontró registro para eliminar
                    End If
                End Using
            End Using
            GetCountDefectosPorFechaYTurno()
            CargarRegistroDeDefectos()
        Catch ex As Exception

        End Try

        GetCountDefectosPorFechaYTurno()
    End Sub
    'Cerrar Applicacion
    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click
        Application.Exit()
    End Sub
    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click
        MsgBox("")

    End Sub
    Public Sub InsertDownTime(TM As String, DT As String, Mesa As String)
        ' Elimina el carácter "/" al inicio de DT
        Dim cleanDT As String = DT.TrimStart("/"c)

        Dim query As String = "INSERT INTO [dbo].[DownTime] ([TM], [DT], [Mesa]) VALUES (@TM, @DT, @Mesa)"

        Using connection As New SqlConnection(cadenaConexion)
            Using command As New SqlCommand(query, connection)
                command.Parameters.AddWithValue("@TM", TM)
                command.Parameters.AddWithValue("@DT", cleanDT)
                command.Parameters.AddWithValue("@Mesa", Mesa)
                Try
                    connection.Open()
                    command.ExecuteNonQuery()
                Catch ex As Exception
                End Try
            End Using
        End Using
    End Sub
    Private Sub Semaforo_Tick(sender As Object, e As EventArgs) Handles Semaforo.Tick
        Label5.Text = "Introdusca Mandril"
    End Sub


End Class