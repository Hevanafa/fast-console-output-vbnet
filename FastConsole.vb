Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Win32.SafeHandles

''' <summary>
''' Reference:
''' https://stackoverflow.com/questions/2754518/
''' </summary>
Public Class FastConsole
    <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Shared Function CreateFile(
        fileName As String,
        <MarshalAs(UnmanagedType.U4)> fileAccess As UInteger,
        <MarshalAs(UnmanagedType.U4)> fileShare As UInteger,
        securityAttributes As IntPtr,
        <MarshalAs(UnmanagedType.U4)> creationDisposition As FileMode,
        <MarshalAs(UnmanagedType.U4)> flags As Integer,
        template As IntPtr) As SafeFileHandle
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Shared Function WriteConsoleOutputW(
        hConsoleOutput As SafeFileHandle,
        lpBuffer As CharInfo(),
        dwBufferSize As Coord,
        dwBufferCoord As Coord,
        ByRef lpWriteRegion As SmallRect) As Boolean
    End Function

    <StructLayout(LayoutKind.Sequential)>
    Public Structure Coord
        Public X As Short
        Public Y As Short

        Public Sub New(X As Short, Y As Short)
            Me.X = X
            Me.Y = Y
        End Sub
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure CharUnion
        <FieldOffset(0)> Public UnicodeChar As UShort
        <FieldOffset(0)> Public AsciiChar As Byte
    End Structure

    <StructLayout(LayoutKind.Explicit)>
    Public Structure CharInfo
        ' https://stackoverflow.com/questions/5868777/
        <FieldOffset(0)> Public [Char] As CharUnion
        <FieldOffset(2)> Public Attributes As Short
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure SmallRect
        Public Left As Short
        Public Top As Short
        Public Right As Short
        Public Bottom As Short
    End Structure

    Property CursorLeft As Integer
        Get
            Return Console.CursorLeft
        End Get
        Set
            Console.CursorLeft = Value
        End Set
    End Property

    Property CursorTop As Integer
        Get
            Return Console.CursorTop
        End Get
        Set
            Console.CursorTop = Value
        End Set
    End Property

    Property Foreground As Short
        Get
            Return Console.ForegroundColor
        End Get
        Set
            Console.ForegroundColor = Value
            ComputedAttributes = Background * 16 + Foreground
        End Set
    End Property

    Property Background As Short
        Get
            Return Console.BackgroundColor
        End Get
        Set
            Console.BackgroundColor = Value
            ComputedAttributes = Background * 16 + Foreground
        End Set
    End Property

    Dim ComputedAttributes As Short

    ' Candidate output
    'Dim sfh As SafeFileHandle = getstdhandle

    Dim fh As SafeFileHandle = CreateFile(
        "CONOUT$",
        &H40000000,
        2,
        IntPtr.Zero,
        FileMode.Open,
        0,
        IntPtr.Zero)

    ReadOnly BufferWidth = 80
    ReadOnly BufferHeight = 25

    Dim char_info_buffer(BufferWidth * BufferHeight) As CharInfo
    Dim dest_rect As New SmallRect With {
        .Left = 0,
        .Top = 0,
        .Right = BufferWidth,
        .Bottom = BufferHeight
    }

    Sub Cls()
        CursorLeft = 0
        CursorTop = 0

        For a = 0 To UBound(char_info_buffer)
            With char_info_buffer(a)
                .Attributes = ComputedAttributes
                .Char.AsciiChar = 0
            End With
        Next
    End Sub

    Sub Print(text$)
        For Each c In text
            Dim idx = CursorTop * BufferWidth + CursorLeft

            With char_info_buffer(idx)
                .Attributes = ComputedAttributes
                .Char.AsciiChar = AscW(c)
            End With

            CursorLeft += 1

            If CursorLeft >= BufferWidth Then
                CursorLeft = 0
                CursorTop += 1
            End If
        Next
    End Sub

    Function Flush() As Boolean
        Flush = WriteConsoleOutputW(
            fh,
            char_info_buffer,
            New Coord() With {.X = BufferWidth, .Y = BufferHeight},
            New Coord() With {.X = 0, .Y = 0},
            dest_rect)
    End Function

    Public Sub New()
        ' initialise ComputedAttributes
        Foreground = Console.ForegroundColor
        Background = Console.BackgroundColor


        If fh.IsInvalid Then Exit Sub

        For ch As Byte = 65 To 65 + 25
            For attribute As Short = 0 To 15
                Cls()

                Background = attribute

                Print("Hello " + Chr(ch))

                'For i = 0 To UBound(char_info_buffer)
                '    char_info_buffer(i).Attributes = attribute
                '    char_info_buffer(i).Char.AsciiChar = ch
                'Next

                Flush()
            Next
        Next

        Console.ReadKey()

    End Sub

End Class
