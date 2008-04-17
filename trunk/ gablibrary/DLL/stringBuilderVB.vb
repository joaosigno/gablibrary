'taken from http://msdn2.microsoft.com/en-us/library/ms972323.aspx

Option Explicit

' default initial size of buffer and growth factor
Private Const DEF_INITIALSIZE As Long = 1000
Private Const DEF_GROWTH As Long = 1000

' buffer size and growth
Private m_nInitialSize As Long
Private m_nGrowth As Long

' buffer and buffer counters
Private m_sText As String
Private m_nSize As Long
Private m_nPos As Long

Private Sub Class_Initialize()
   ' set defaults for size and growth
   m_nInitialSize = DEF_INITIALSIZE
   m_nGrowth = DEF_GROWTH
   ' initialize buffer
   InitBuffer
End Sub

' set the initial size and growth amount
Public Sub Init(ByVal InitialSize As Long, ByVal Growth As Long)
   If InitialSize > 0 Then m_nInitialSize = InitialSize
   If Growth > 0 Then m_nGrowth = Growth
End Sub

' initialize the buffer
Private Sub InitBuffer()
   m_nSize = -1
   m_nPos = 1
End Sub

' grow the buffer
Private Sub Grow(Optional MinimimGrowth As Long)
   ' initialize buffer if necessary
   If m_nSize = -1 Then
      m_nSize = m_nInitialSize
      m_sText = Space$(m_nInitialSize)
   Else
      ' just grow
      Dim nGrowth As Long
      nGrowth = IIf(m_nGrowth > MinimimGrowth, 
            m_nGrowth, MinimimGrowth)
      m_nSize = m_nSize + nGrowth
      m_sText = m_sText & Space$(nGrowth)
   End If
End Sub

' trim the buffer to the currently used size
Private Sub Shrink()
   If m_nSize > m_nPos Then
      m_nSize = m_nPos - 1
      m_sText = RTrim$(m_sText)
   End If
End Sub

' add a single text string
Private Sub AppendInternal(ByVal Text As String)
   If (m_nPos + Len(Text)) > m_nSize Then Grow Len(Text)
   Mid$(m_sText, m_nPos, Len(Text)) = Text
   m_nPos = m_nPos + Len(Text)
End Sub

' add a number of text strings
Public Sub Append(ParamArray Text())
   Dim nArg As Long
   For nArg = 0 To UBound(Text)
      AppendInternal CStr(Text(nArg))
   Next nArg
End Sub
 
' return the current string data and trim the buffer
Public Function ToString() As String
   If m_nPos > 0 Then
      Shrink
      ToString = m_sText
   Else
      ToString = ""
   End If
End Function

' clear the buffer and reinit
Public Sub Clear()
   InitBuffer
End Sub
