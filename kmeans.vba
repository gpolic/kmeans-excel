Option Base 1
Option Explicit


Public Sub kmeans()
    Dim InitCentrRange As String: InitCentrRange = ActiveSheet.Range("C7").Value
    Dim InitialCentroids As Variant: InitialCentroids = ActiveSheet.Range(InitCentrRange)
    
    Application.StatusBar = "   [ Initialize ]"
    Dim MaxIt As Integer: MaxIt = ActiveSheet.Range("MaxIt").Value
    Dim DataSht As String: DataSht = ActiveSheet.Range("C3").Value
    Dim DataRange As String: DataRange = ActiveSheet.Range("C4").Value
    Dim OutputRange As String: OutputRange = ActiveSheet.Range("C8").Value
    Dim X As Variant: X = Worksheets(DataSht).Range(DataRange)
    
    ' first pass
    Dim my_idx As Variant: my_idx = FindClosestCentroid(X, InitialCentroids)
    
    Dim K As Integer: K = UBound(InitialCentroids, 1)
    Dim J As Integer: J = UBound(X, 2)
    Dim M As Integer: M = UBound(X, 1)
    
    Dim centroids As Variant
    Dim ii As Integer: ii = 1
    
    'Application.ScreenUpdating = False
    
    For ii = 1 To MaxIt
        Application.StatusBar = "   [ Pass:" + CStr(ii) + " ]"
        centroids = ComputeCentroids(X, my_idx, K)
        my_idx = FindClosestCentroid(X, centroids)
    Next ii
    
    Range(OutputRange).Resize(K, J).Value = centroids
    
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = ActiveSheet.Range("C5").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = ActiveSheet.Range("C6").Value
    
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(M, 1).Value = WorksheetFunction.Transpose(my_idx)
    
End Sub


Public Function EuclideanDistance(X As Variant, Y As Variant, num_obs As Integer) As Double
    Dim ii As Integer: ii = 1
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For ii = 1 To num_obs
        RunningSumSqr = RunningSumSqr + ((X(ii) - Y(ii)) ^ 2)
    Next ii
    
    EuclideanDistance = Sqr(RunningSumSqr)
End Function


Public Function FindClosestCentroid(ByRef X As Variant, ByRef centroids As Variant) As Variant
    Dim K As Integer: K = UBound(centroids, 1)
    Dim J As Integer: J = UBound(centroids, 2)
    Dim M As Integer: M = UBound(X, 1)
    Dim idx() As Variant: ReDim idx(M) As Variant
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1

    For ii = 1 To M
        Dim Dist_min As Double: Dist_min = 10000000
        Dim Dist As Double: Dist = 0
        For cc = 1 To K
            Dist = EuclideanDistance(Application.Index(X, ii, 0), Application.Index(centroids, cc, 0), J)
            If Dist < Dist_min Then
                idx(ii) = cc
                Dist_min = Dist
            End If
            
        Next cc
    Next ii
    FindClosestCentroid = idx()
End Function


Public Function ComputeCentroids(X As Variant, idx As Variant, K As Variant) As Variant
    Dim M As Integer: M = UBound(X, 1)
    Dim J As Integer: J = UBound(X, 2)
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1
    Dim bb As Integer: bb = 1
    Dim counter As Integer: counter = 0
    Dim tempSum() As Variant: ReDim tempSum(K, J) As Variant
    Dim centroids() As Variant: ReDim centroids(K, J) As Variant
    
    For ii = 1 To K
        For bb = 1 To J
            counter = 0
            For cc = 1 To M
                If idx(cc) = ii Then
                centroids(ii, bb) = centroids(ii, bb) + X(cc, bb)
                counter = counter + 1
                End If
            Next cc
            If counter > 0 Then
                centroids(ii, bb) = centroids(ii, bb) / counter
            Else
                centroids(ii, bb) = 0
            End If
        Next bb
    Next ii
    ComputeCentroids = centroids
End Function



