VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsToastQueueManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************
' Class: clsToastQueueManager
' Purpose: Intelligent toast queue with priority, stacking, and collision detection
' Version: 6.0
' Features: Multiple simultaneous toasts, smart positioning, priority queue
'***************************************************************
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
#End If

Private Const SM_CXSCREEN As Long = 0
Private Const SM_CYSCREEN As Long = 1
Private Const MAX_VISIBLE_TOASTS As Long = 5
Private Const TOAST_STACK_SPACING As Long = 10
Private Const TOAST_HEIGHT As Long = 160

' === PRIORITY LEVELS ===
Public Enum ToastPriority
    PriorityLow = 1
    PriorityNormal = 2
    PriorityHigh = 3
    PriorityCritical = 4
End Enum

' === QUEUE ITEM ===
Private Type QueuedToast
    toast As clsToastNotification
    Priority As ToastPriority
    QueueTime As Double
    IsShowing As Boolean
    StackPosition As Long
End Type

' === STATE ===
Private m_Queue() As QueuedToast
Private m_QueueCount As Long
Private m_MaxVisible As Long
Private m_DoNotDisturb As Boolean
Private m_ProcessingEnabled As Boolean
Private m_DefaultPosition As String
Private m_AutoStackEnabled As Boolean

' === INITIALIZATION ===
Private Sub Class_Initialize()
    ReDim m_Queue(0 To 99)
    m_QueueCount = 0
    m_MaxVisible = MAX_VISIBLE_TOASTS
    m_DoNotDisturb = False
    m_ProcessingEnabled = True
    m_DefaultPosition = "BR"
    m_AutoStackEnabled = True
End Sub

Private Sub Class_Terminate()
    ClearAll
End Sub

' === PROPERTIES ===
Public Property Let MaxVisibleToasts(ByVal value As Long)
    If value > 0 And value <= 20 Then m_MaxVisible = value
End Property
Public Property Get MaxVisibleToasts() As Long
    MaxVisibleToasts = m_MaxVisible
End Property

Public Property Let DoNotDisturb(ByVal value As Boolean)
    m_DoNotDisturb = value
End Property
Public Property Get DoNotDisturb() As Boolean
    DoNotDisturb = m_DoNotDisturb
End Property

Public Property Let ProcessingEnabled(ByVal value As Boolean)
    m_ProcessingEnabled = value
    If value Then ProcessQueue
End Property
Public Property Get ProcessingEnabled() As Boolean
    ProcessingEnabled = m_ProcessingEnabled
End Property

Public Property Let DefaultPosition(ByVal value As String)
    m_DefaultPosition = UCase$(value)
End Property
Public Property Get DefaultPosition() As String
    DefaultPosition = m_DefaultPosition
End Property

Public Property Let AutoStack(ByVal value As Boolean)
    m_AutoStackEnabled = value
End Property
Public Property Get AutoStack() As Boolean
    AutoStack = m_AutoStackEnabled
End Property

Public Property Get QueueLength() As Long
    QueueLength = m_QueueCount
End Property

Public Property Get visibleCount() As Long
    Dim i As Long, count As Long
    For i = 0 To m_QueueCount - 1
        If m_Queue(i).IsShowing Then count = count + 1
    Next
    visibleCount = count
End Property

' === PUBLIC METHODS ===

' Quick notification methods
Public Sub NotifyInfo(ByVal Title As String, ByVal Message As String, _
                     Optional ByVal Priority As ToastPriority = PriorityNormal)
    QuickNotify Title, Message, "INFO", Priority
End Sub

Public Sub NotifyWarning(ByVal Title As String, ByVal Message As String, _
                        Optional ByVal Priority As ToastPriority = PriorityHigh)
    QuickNotify Title, Message, "WARNING", Priority
End Sub

Public Sub NotifyError(ByVal Title As String, ByVal Message As String, _
                      Optional ByVal Priority As ToastPriority = PriorityCritical)
    QuickNotify Title, Message, "ERROR", Priority
End Sub

Public Sub NotifySuccess(ByVal Title As String, ByVal Message As String, _
                        Optional ByVal Priority As ToastPriority = PriorityNormal)
    QuickNotify Title, Message, "SUCCESS", Priority
End Sub

Private Sub QuickNotify(ByVal Title As String, ByVal Message As String, _
                       ByVal Level As String, ByVal Priority As ToastPriority)
    Dim toast As New clsToastNotification
    toast.Title = Title
    toast.Message = Message
    toast.Level = Level
    toast.Position = m_DefaultPosition
    toast.Duration = 5
    
    Enqueue toast, Priority
End Sub

' Add toast to queue
Public Sub Enqueue(ByVal toast As clsToastNotification, _
                  Optional ByVal Priority As ToastPriority = PriorityNormal)
    On Error GoTo ErrorHandler
    
    ' Check if we need to expand array
    If m_QueueCount >= UBound(m_Queue) Then
        ReDim Preserve m_Queue(0 To UBound(m_Queue) + 100)
    End If
    
    ' Add to queue
    Set m_Queue(m_QueueCount).toast = toast
    m_Queue(m_QueueCount).Priority = Priority
    m_Queue(m_QueueCount).QueueTime = Timer
    m_Queue(m_QueueCount).IsShowing = False
    m_Queue(m_QueueCount).StackPosition = -1
    
    m_QueueCount = m_QueueCount + 1
    
    ' Sort by priority
    SortQueue
    
    ' Process immediately if enabled
    If m_ProcessingEnabled And Not m_DoNotDisturb Then
        ProcessQueue
    End If
    
    Exit Sub
ErrorHandler:
    Debug.Print "[QueueManager] Enqueue error: " & Err.Description
End Sub

' Process queue and show toasts
Public Sub ProcessQueue()
    On Error Resume Next
    
    If m_DoNotDisturb Then Exit Sub
    If Not m_ProcessingEnabled Then Exit Sub
    
    Dim visibleCount As Long
    visibleCount = Me.visibleCount
    
    ' Show toasts up to max visible
    Dim i As Long
    For i = 0 To m_QueueCount - 1
        If Not m_Queue(i).IsShowing Then
            If visibleCount < m_MaxVisible Then
                ' Calculate stack position
                If m_AutoStackEnabled Then
                    m_Queue(i).StackPosition = visibleCount
                    AdjustToastPosition m_Queue(i).toast, visibleCount
                End If
                
                ' Show toast
                If m_Queue(i).toast.Show() Then
                    m_Queue(i).IsShowing = True
                    visibleCount = visibleCount + 1
                End If
            Else
                Exit For
            End If
        End If
    Next
End Sub

' Remove completed toasts from queue
Public Sub CleanupCompleted()
    On Error Resume Next
    
    Dim i As Long
    For i = m_QueueCount - 1 To 0 Step -1
        If m_Queue(i).IsShowing And Not m_Queue(i).toast.IsShowing Then
            ' Toast has closed, remove from queue
            RemoveAt i
        End If
    Next
    
    ' Restack remaining toasts
    If m_AutoStackEnabled Then RestackToasts
End Sub

' Clear all toasts
Public Sub ClearAll()
    On Error Resume Next
    
    Dim i As Long
    For i = 0 To m_QueueCount - 1
        If m_Queue(i).IsShowing Then
            m_Queue(i).toast.Close
        End If
        Set m_Queue(i).toast = Nothing
    Next
    
    m_QueueCount = 0
End Sub

' Pause/Resume processing
Public Sub Pause()
    m_ProcessingEnabled = False
End Sub

Public Sub Resume()
    m_ProcessingEnabled = True
    ProcessQueue
End Sub

' Enable Do Not Disturb for specified duration (minutes)
Public Sub EnableDND(Optional ByVal DurationMinutes As Long = 60)
    m_DoNotDisturb = True
    
    If DurationMinutes > 0 Then
        ' Schedule auto-disable
        Dim endTime As Double
        endTime = Timer + (DurationMinutes * 60)
        
        ' Note: In production, you'd use Application.OnTime for this
        ' For now, caller should manually disable DND
    End If
End Sub

Public Sub DisableDND()
    m_DoNotDisturb = False
    ProcessQueue
End Sub

' Get queue statistics
Public Function GetStats() As String
    Dim stats As String
    stats = "=== Toast Queue Statistics ===" & vbCrLf
    stats = stats & "Queued: " & m_QueueCount & vbCrLf
    stats = stats & "Visible: " & Me.visibleCount & vbCrLf
    stats = stats & "Max Visible: " & m_MaxVisible & vbCrLf
    stats = stats & "DND Mode: " & m_DoNotDisturb & vbCrLf
    stats = stats & "Processing: " & m_ProcessingEnabled & vbCrLf
    stats = stats & "Auto Stack: " & m_AutoStackEnabled & vbCrLf
    
    ' Priority breakdown
    Dim pLow As Long, pNormal As Long, pHigh As Long, pCritical As Long
    Dim i As Long
    For i = 0 To m_QueueCount - 1
        Select Case m_Queue(i).Priority
            Case PriorityLow: pLow = pLow + 1
            Case PriorityNormal: pNormal = pNormal + 1
            Case PriorityHigh: pHigh = pHigh + 1
            Case PriorityCritical: pCritical = pCritical + 1
        End Select
    Next
    
    stats = stats & vbCrLf & "Priority Breakdown:" & vbCrLf
    stats = stats & "  Critical: " & pCritical & vbCrLf
    stats = stats & "  High: " & pHigh & vbCrLf
    stats = stats & "  Normal: " & pNormal & vbCrLf
    stats = stats & "  Low: " & pLow
    
    GetStats = stats
End Function

' === PRIVATE METHODS ===

Private Sub SortQueue()
    ' Bubble sort by priority (Critical first)
    Dim i As Long, j As Long
    Dim temp As QueuedToast
    
    For i = 0 To m_QueueCount - 2
        For j = i + 1 To m_QueueCount - 1
            If m_Queue(j).Priority > m_Queue(i).Priority Then
                ' Swap
                temp = m_Queue(i)
                m_Queue(i) = m_Queue(j)
                m_Queue(j) = temp
            End If
        Next j
    Next i
End Sub

Private Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index >= m_QueueCount Then Exit Sub
    
    Set m_Queue(Index).toast = Nothing
    
    ' Shift remaining items down
    Dim i As Long
    For i = Index To m_QueueCount - 2
        m_Queue(i) = m_Queue(i + 1)
    Next
    
    m_QueueCount = m_QueueCount - 1
End Sub

Private Sub AdjustToastPosition(ByRef toast As clsToastNotification, ByVal stackPos As Long)
    ' Modify toast position to stack vertically
    Dim basePos As String
    basePos = m_DefaultPosition
    
    ' Calculate offset based on stack position
    Dim offsetY As Long
    offsetY = stackPos * (TOAST_HEIGHT + TOAST_STACK_SPACING)
    
    ' For now, we can't dynamically adjust HTA position after creation
    ' This would require a more complex system with a positioning coordinator
    ' Left as TODO for advanced implementation
End Sub

Private Sub RestackToasts()
    ' Recalculate positions for all visible toasts
    Dim stackPos As Long
    Dim i As Long
    
    For i = 0 To m_QueueCount - 1
        If m_Queue(i).IsShowing Then
            m_Queue(i).StackPosition = stackPos
            stackPos = stackPos + 1
        End If
    Next
End Sub

'***************************************************************
' USAGE EXAMPLE:
'
' Sub TestQueueManager()
'     Dim qm As New clsToastQueueManager
'     qm.MaxVisibleToasts = 3
'     qm.AutoStack = True
'
'     ' Add multiple toasts
'     qm.NotifyInfo "Download", "File 1 downloading...", PriorityNormal
'     qm.NotifyInfo "Download", "File 2 downloading...", PriorityNormal
'     qm.NotifyWarning "Warning", "Disk space low", PriorityHigh
'     qm.NotifyError "Error", "Network disconnected", PriorityCritical
'
'     ' The critical error will show first
'     ' Only 3 will be visible at once
'
'     ' Show stats
'     Debug.Print qm.GetStats
'
'     ' Enable DND mode
'     qm.EnableDND 30  ' 30 minutes
'
'     ' These will queue but not show
'     qm.NotifyInfo "Test", "This is queued during DND"
'
'     ' Disable DND
'     qm.DisableDND  ' Queued toasts will now show
' End Sub
'***************************************************************

