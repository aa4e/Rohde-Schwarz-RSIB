Imports System.Runtime.InteropServices
Imports System.Text

Namespace RS

    ''' <summary>
    ''' Работа с устройствами Rohde-Schwarz по протоколу RSIB.
    ''' </summary>
    Public Class RSIB
        Implements IDisposable

#Region "CTOR"

        Private Const CLOSED_HANDLE As UShort = &H8000US

        ''' <summary>
        ''' Дескриптор устройства. Присваивается во время открытия устройства.
        ''' </summary>
        Private ReadOnly Property DevHandle As UShort
            Get
                Return _DevHandle
            End Get
        End Property
        Private _DevHandle As UShort = CLOSED_HANDLE

        ''' <summary>
        ''' Ищет устройство RSIB и подключается к нему по IP-адресу <paramref name="ipAddr"/>.
        ''' </summary>
        ''' <param name="ipAddr">IP-адрес устройства.</param>
        Public Sub New(ipAddr As String)
            Dim ibsta As Status 'Флаги статуса.
            Dim iberr As Errors 'Код ошибки
            Dim ibcntl As UInteger 'счётчик показывает число переданных байтов
            _DevHandle = RSDLLibfind(ipAddr, ibsta, iberr, ibcntl)
            CheckResult(ibsta, iberr)
        End Sub

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibfind(udName As String, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

        ''' <summary>
        ''' Отключается от устройства.
        ''' </summary>
        ''' <remarks>
        ''' После закрытия устройства программный доступ к устройству становится невозможен.
        ''' </remarks>
        Public Sub Close()
            If (_DevHandle <> CLOSED_HANDLE) Then
                Try
                    Dim ibsta As Status
                    Dim iberr As Errors
                    Dim ibcntl As UInteger
                    Dim res As Integer = RSDLLibonl(DevHandle, Modes.Local, ibsta, iberr, ibcntl)
                    _DevHandle = CLOSED_HANDLE
                Catch ex As Exception
                    Debug.WriteLine(ex)
                End Try
            End If
        End Sub

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibonl(udName As Integer, mode As Modes, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

#End Region '/CTOR

#Region "ЗАПИСЬ, ЧТЕНИЕ"

        ''' <summary>
        ''' Отправляет команду и читает ответ.
        ''' </summary>
        ''' <param name="command"></param>
        Public Function SendCommand(command As String) As String
            Dim ibsta As Status
            Dim iberr As Errors
            Dim ibcntl As UInteger
            Dim res As Integer = RSDLLibwrt(DevHandle, command & vbNullChar, ibsta, iberr, ibcntl)
            CheckResult(ibsta, iberr)
            If command.Contains("?"c) Then
                Dim sbAns As New StringBuilder(100)
                Dim ans As Integer = RSDLLibrd(DevHandle, sbAns, ibsta, iberr, ibcntl)
                CheckResult(ibsta, iberr)
                Return sbAns.ToString()
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' This function sends zero-terminated string data to the device with the handle ud.
        ''' </summary>
        ''' <param name="udName"></param>
        ''' <param name="data"></param>
        ''' <param name="ibsta"></param>
        ''' <param name="iberr"></param>
        ''' <param name="ibcntl"></param>
        ''' <remarks>
        ''' This function allows setting and query commands to be sent to the measuring instruments. 
        ''' Whether the data Is interpreted as a complete command can be set using the Function RSDLLibeot().
        ''' </remarks>
        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibwrt(udName As Integer, data As String, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As Integer
        End Function

        ''' <summary>
        ''' The function reads data from the device with the handle ud into the string Rd.
        ''' </summary>
        ''' <param name="udName"></param>
        ''' <param name="data"></param>
        ''' <param name="ibsta"></param>
        ''' <param name="iberr"></param>
        ''' <param name="ibcntl"></param>
        ''' <remarks>
        ''' In the case of Visual Basic programming, a string of sufficient length must be generated beforehand. 
        ''' This can be done during the definition Of the String Or Using the command Space$().
        ''' </remarks>
        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibrd(udName As Integer, data As StringBuilder, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

#End Region '/ЗАПИСЬ, ЧТЕНИЕ 

#Region "УПРАВЛЕНИЕ РЕЖИМОМ"

        ''' <summary>
        ''' Переключает устройство в режим <paramref name="mode"/>.
        ''' </summary>
        ''' <remarks>
        ''' После переключения в режим <see cref="Modes.TemporarilyLocal"/> устройство может управляться вручную через фронтальную панель.
        ''' При следующем удалённом доступе устройство снова переключится в режим <see cref="Modes.Remote"/>.
        ''' </remarks>
        Public Sub SetMode(mode As Modes)
            Dim ibsta As Status
            Dim iberr As Errors
            Dim ibcntl As UInteger
            Select Case mode
                Case Modes.Local, Modes.Remote
                    Dim res As Integer = RSDLLibsre(DevHandle, mode, ibsta, iberr, ibcntl)
                Case Modes.TemporarilyLocal
                    Dim res As Integer = RSDLLibloc(DevHandle, ibsta, iberr, ibcntl)
            End Select
            CheckResult(ibsta, iberr)
        End Sub

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibsre(udName As Integer, mode As Modes, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibloc(udName As Integer, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

        ''' <summary>
        ''' Включает или выключает отправку конца сообщения после операции записи.
        ''' </summary>
        ''' <param name="sendEnd">Если False, то команды могут отправляться несколькими последовательными вызовами <see cref="SendCommand(String)"/>.
        ''' True должно быть снова установлено перед отправкой последнего блока данных.
        ''' </param>
        ''' <remarks>
        ''' If the END message is disabled, the data of a command can be sent with several successive calls of write functions. 
        ''' The END message must be enabled again before sending the last data block.
        ''' </remarks>
        Public Sub SetEndOfMessage(sendEnd As Boolean)
            Dim ibsta As Status
            Dim iberr As Errors
            Dim ibcntl As UInteger
            Dim res As Integer = RSDLLibeot(DevHandle, sendEnd, ibsta, iberr, ibcntl)
            CheckResult(ibsta, iberr)
        End Sub

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibeot(udName As Integer, <MarshalAs(UnmanagedType.Bool)> sendEnd As Boolean, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

#End Region '/УПРАВЛЕНИЕ РЕЖИМОМ

#Region "NATIVE"

        Private Const LibPath As String = "c:\Temp\rsib32.dll"

#Region "ERRORS, STATUS"

        ''' <summary>
        ''' Опрашивает устройство и возвращает его регистр статуса.
        ''' </summary>
        Public Function GetDeviceStatusRegister() As Byte
            Dim ibsta As Status
            Dim iberr As Errors
            Dim ibcntl As UInteger
            Dim spr As Byte
            Dim res As Integer = RSDLLibrsp(DevHandle, spr, ibsta, iberr, ibcntl)
            CheckResult(ibsta, iberr)
            Return spr
        End Function

        <DllImport(LibPath, CharSet:=CharSet.Ansi)>
        Private Shared Function RSDLLibrsp(udName As Integer, ByRef spr As Byte, ByRef ibsta As Status, ByRef iberr As Errors, ByRef ibcntl As UInteger) As UShort
        End Function

        ''' <summary>
        ''' Проверяет статус и ошибки.
        ''' </summary>
        ''' <param name="err">Р</param>
        ''' <param name="stat"></param>
        Private Sub CheckResult(stat As Status, err As Errors)
            If IsErrorOccured(stat) Then
                Throw New Exception(GetErrorText(err))
            ElseIf IsTimeout(stat) Then
                Throw New Exception("Превышено время ожидания.")
            End If
        End Sub

        ''' <summary>
        ''' Возвращает текст ошибки.
        ''' </summary>
        Private Function GetErrorText(err As Errors) As String
            Select Case err
                Case Errors.IBERR_TIMEOUT
                    Return "Превышено время ожидания."
                Case Errors.IBERR_BUSY
                    Return "Протокол RSIB заблокирован функцией, которая ещё выполняется."
                Case Errors.IBERR_CONNECT
                    Return "Подключение к инструменту не удалось."
                Case Errors.IBERR_NO_DEVICE
                    Return "Функция была вызвана с некорректным идентификатором устройства."
                Case Errors.IBERR_MEM
                    Return "Нет свободной памяти."
                Case Errors.IBERR_FILE
                    Return "Ошибка во время чтения или записи в файл."
                Case Errors.IBERR_SEMA
                    Return "Ошибка создания или присваивания семафора (только для UNIX)."
                Case Else
                    Return "Ошибка во время работы с устройством."
            End Select
        End Function

        Private Function IsAnswerReady(stat As Status) As Boolean
            Return ((stat And Status.CMPL) = Status.CMPL)
        End Function

        Private Function IsTimeout(stat As Status) As Boolean
            Return ((stat And Status.TIMO) = Status.TIMO)
        End Function

        Private Function IsErrorOccured(stat As Status) As Boolean
            Return ((stat And Status.ERR) = Status.ERR)
        End Function

#End Region '/ERRORS, STATUS

#Region "ENUMS"

        ''' <summary>
        ''' Status word - ibsta.
        ''' </summary>
        ''' <remarks>
        ''' The status word ibsta provides information On the status Of the RSIB Interface.
        ''' </remarks>
        <Flags()>
        Public Enum Status As UShort
            ''' <summary>
            ''' Флаг установлен, когда произошла ошибка во время вызова функции.
            ''' </summary>
            ''' <remarks>
            ''' При наличии этого флага в статусе необходимо проверить код ошибки для выяснения деталей.
            ''' </remarks>
            ERR = &H8000
            ''' <summary>
            ''' Флаг установлен, когда превышено время ожидания во время вызова функции.
            ''' </summary>
            TIMO = &H4000
            ''' <summary>
            ''' Флаг установлен, когда устройство завершило обработку принятой команды.
            ''' Флаг сброшен, когда ответ устройства прочитан или длина буфера недостаточна для ответа.
            ''' </summary>
            CMPL = &H100
        End Enum

        ''' <summary>
        ''' Error variable - iberr.
        ''' </summary>
        ''' <remarks>
        ''' If the ERR bit (8000h) Is Set In the status word, iberr contains an Error code which
        ''' allows the Error To be specified In greater detail. Extra Error codes are defined For the
        ''' RSIB protocol, independent Of the National Instruments Interface.
        ''' </remarks>
        Public Enum Errors As UShort
            None = 0
            ''' <summary>
            ''' Подключение к инструменту не удалось.
            ''' </summary>
            IBERR_CONNECT = 2
            ''' <summary>
            ''' Функция была вызвана с некорректным идентификатором устройства.
            ''' </summary>
            IBERR_NO_DEVICE = 3
            ''' <summary>
            ''' Нет свободной памяти.
            ''' </summary>
            IBERR_MEM = 4
            ''' <summary>
            ''' Превышено время ожидания.
            ''' </summary>
            IBERR_TIMEOUT = 5
            ''' <summary>
            ''' Протокол RSIB заблокирован функцией, которая ещё выполняется.
            ''' </summary>
            IBERR_BUSY = 6
            ''' <summary>
            ''' Ошибка во время чтения или записи в файл.
            ''' </summary>
            IBERR_FILE = 7
            ''' <summary>
            ''' Ошибка создания или присваивания семафора (только для UNIX).
            ''' </summary>
            IBERR_SEMA = 8
        End Enum

        Public Enum Modes As UShort
            Local = 0
            Remote = 1
            TemporarilyLocal = 2
        End Enum

#End Region '/ENUMS

#End Region '/NATIVE

#Region "IDISPOSABLE"

        Private disposedValue As Boolean

        Protected Overridable Sub Dispose(disposing As Boolean)
            If (Not disposedValue) Then
                If disposing Then
                    'dispose managed state (managed objects)
                End If
                'free unmanaged resources (unmanaged objects) and override finalizer;  set large fields to null
                Close()
                disposedValue = True
            End If
        End Sub

        ' ' TODO: override finalizer only if 'Dispose(disposing As Boolean)' has code to free unmanaged resources
        ' Protected Overrides Sub Finalize()
        '     ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
        '     Dispose(disposing:=False)
        '     MyBase.Finalize()
        ' End Sub

        Public Sub Dispose() Implements IDisposable.Dispose
            ' Do not change this code. Put cleanup code in 'Dispose(disposing As Boolean)' method
            Dispose(disposing:=True)
            GC.SuppressFinalize(Me)
        End Sub

#End Region '/IDISPOSABLE

    End Class

End Namespace