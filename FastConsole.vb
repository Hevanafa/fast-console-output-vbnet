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

    Public Sub New()
        Dim h As SafeFileHandle = CreateFile(
            "CONOUT$",
            &H40000000,
            2,
            IntPtr.Zero,
            FileMode.Open,
            0,
            IntPtr.Zero)

        If h.IsInvalid Then Exit Sub

        Dim buf(80 * 25) As CharInfo
        Dim rect As New SmallRect With {
            .Left = 0,
            .Top = 0,
            .Right = 80,
            .Bottom = 25
        }

        For ch As Byte = 65 To 65 + 25
            For attribute As Short = 0 To 15
                For i = 0 To UBound(buf)
                    buf(i).Attributes = attribute
                    buf(i).Char.AsciiChar = ch
                Next


                Dim b As Boolean = WriteConsoleOutputW(
                    h,
                    buf,
                    New Coord() With {.X = 80, .Y = 25},
                    New Coord() With {.X = 0, .Y = 0},
                    rect)
            Next
        Next

        Console.ReadKey()

    End Sub

End Class
