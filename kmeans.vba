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
    Dim DataRecords As Variant: DataRecords = Worksheets(DataSht).Range(DataRange)      ' this holds all the data records
    Dim Cluster_Indexes As Variant
    
    Dim K As Integer: K = UBound(InitialCentroids, 1)                     ' K is the number of clusters
    Dim J As Integer: J = UBound(DataRecords, 2)
    Dim M As Integer: M = UBound(DataRecords, 1)
    Dim Centroids As Variant
    Dim ii As Integer
    
    'Application.ScreenUpdating = False
    
    ' first pass
    Cluster_Indexes = FindClosestCentroid(DataRecords, InitialCentroids)  ' assign each record(observation) in a initial cluster
    
    For ii = 1 To MaxIt
        Application.StatusBar = "   [ Pass:" + CStr(ii) + " ]"
        Centroids = ComputeCentroids(DataRecords, Cluster_Indexes, K)         ' calculate new centroids for each cluster
        Cluster_Indexes = FindClosestCentroid(DataRecords, Centroids)         ' assign each record in a cluster based on the new centroids
    Next ii
    
    Range(OutputRange).Resize(K, J).Value = Centroids
    
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = ActiveSheet.Range("C5").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = ActiveSheet.Range("C6").Value
    
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(M, 1).Value = WorksheetFunction.Transpose(Cluster_Indexes)
    
End Sub


Public Function EuclideanDistance(X As Variant, Y As Variant, num_obs As Integer) As Double
    Dim ii As Integer: ii = 1
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For ii = 1 To num_obs
        RunningSumSqr = RunningSumSqr + ((X(ii) - Y(ii)) ^ 2)
    Next ii
    
    EuclideanDistance = Sqr(RunningSumSqr)
End Function


Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant) As Variant
    Dim K As Integer: K = UBound(Centroids, 1)      ' number of clusters
    Dim J As Integer: J = UBound(Centroids, 2)       ' number of columns (features)
    Dim M As Integer: M = UBound(DataRecords, 1)     ' number of data records
    Dim idx() As Variant: ReDim idx(M) As Variant
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1

    For ii = 1 To M      ' for all records
    
        Dim Dist_min As Double: Dist_min = 10000000
        Dim Dist As Double: Dist = 0
        For cc = 1 To K                         ' for all clusters
            Dist = EuclideanDistance(Application.Index(DataRecords, ii, 0), Application.Index(Centroids, cc, 0), J)
            If Dist < Dist_min Then
                idx(ii) = cc              ' this record is assigned to cluster cc as this is the minimumm distance
                Dist_min = Dist
            End If
        ' pass the Cluster_index here and see if it is equal the idx.  if not then indicate a change in the clusters (global boolean var ??)
        Next cc
    Next ii
    FindClosestCentroid = idx()
End Function



Public Function ComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, NoOfClusters As Variant) As Variant
    Dim M As Integer: M = UBound(DataRecords, 1)
    Dim RecordSize As Integer: RecordSize = UBound(DataRecords, 2)
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1
    Dim bb As Integer: bb = 1
    Dim counter As Integer: counter = 0
    Dim tempSum() As Variant: ReDim tempSum(NoOfClusters, RecordSize) As Variant
    Dim Centroids() As Variant: ReDim Centroids(NoOfClusters, RecordSize) As Variant
    
    For ii = 1 To NoOfClusters       ' for all clusters
        For bb = 1 To RecordSize     ' for all features(columns)
            counter = 0
            For cc = 1 To M                   ' for every observation (data record)
                If ClusterIdx(cc) = ii Then   ' is this record part of cluster ii ??
                    Centroids(ii, bb) = Centroids(ii, bb) + DataRecords(cc, bb)
                    counter = counter + 1
                End If
            Next cc
            If counter > 0 Then
                Centroids(ii, bb) = Centroids(ii, bb) / counter    ' compute the new centroid averaging all records in the cluster
            Else
                Centroids(ii, bb) = 0
            End If
        Next bb
    Next ii
    ComputeCentroids = Centroids
    
End Function



