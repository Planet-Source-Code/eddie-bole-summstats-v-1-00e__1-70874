VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHeapSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Sub heapSort(heap() As Long)
    Dim tree()  As Long
    Dim i As Long
    
    ReDim Preserve tree(1 To 1)
    For i = 1 To UBound(heap)      ' construct heap
        Call insHeap(tree, i - 1, heap(i))
        ReDim Preserve tree(1 To i + 1)
    Next i
    ReDim Preserve tree(1 To UBound(tree) - 1)
    
    For i = 1 To UBound(tree)
        Call delHeap(tree, UBound(tree), heap(i))
        If UBound(tree) > 1 Then
            ReDim Preserve tree(1 To UBound(tree) - 1)
        End If
    Next i
End Sub

Private Sub insHeap(tree() As Long, totalElements As Long, item As Long)
    Dim ptr As Long
    Dim par As Long
    
    totalElements = totalElements + 1
    ptr = totalElements
    While ptr > 1
        par = Fix(ptr / 2)
        If item >= tree(par) Then
            tree(ptr) = item
            Exit Sub
        End If
        tree(ptr) = tree(par)
        ptr = par
    Wend
    tree(1) = item
End Sub

Private Sub delHeap(tree() As Long, totalElements As Long, item As Long)
    Dim ptr As Long, last As Long, left As Long, right As Long
    
    item = tree(1)
    last = tree(totalElements): totalElements = totalElements - 1
    ptr = 1: left = 2: right = 3
    
    While right <= totalElements
        If last <= tree(left) And last <= tree(right) Then
            tree(ptr) = last
            Exit Sub
        End If
        If tree(right) >= tree(left) Then
            tree(ptr) = tree(left)
            ptr = left
        Else
            tree(ptr) = tree(right)
            ptr = right
        End If
        left = 2 * ptr
        right = left + 1
    Wend
    If left = totalElements Then
        If last > tree(left) Then
            tree(ptr) = tree(left)
            ptr = left
        End If
    End If
    tree(ptr) = last
End Sub
