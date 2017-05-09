Option Base 1
Option Explicit


Public Sub kmeans()
    Dim InitCentrRange As String: InitCentrRange = ActiveSheet.Range("C7").Value
    Dim InitialCentroids As Variant: InitialCentroids = ActiveSheet.Range(InitCentrRange)

    Dim MaxIt As Integer: MaxIt = ActiveSheet.Range("MaxIt").Value
    Dim DataSht As String: DataSht = ActiveSheet.Range("C3").Value
    Dim DataRange As String: DataRange = ActiveSheet.Range("C4").Value
    Dim OutputRange As String: OutputRange = ActiveSheet.Range("C8").Value
    Dim DataRecords As Variant: DataRecords = Worksheets(DataSht).Range(DataRange)      ' this holds all the data records
    Dim ClusterIndexes As Variant
    
    Dim K As Integer: K = UBound(InitialCentroids, 1)                     ' K is the number of clusters
    Dim J As Integer: J = UBound(DataRecords, 2)
    Dim M As Integer: M = UBound(DataRecords, 1)
    Dim Centroids As Variant
    Dim ii As Integer
    Dim ClustersUpdated As Integer: ClustersUpdated = 10000
    
    
    Application.StatusBar = "   [ Initialize ]"
    'Application.ScreenUpdating = False
    
    ' First pass. We do not need the result here, so put it in ii
    ii = FindClosestCentroid(DataRecords, InitialCentroids, ClusterIndexes)  ' assign each record(observation) in a initial cluster
    
    ii = 1
    While ii <= MaxIt And ClustersUpdated > 0  ' We will process k-means until it is normalized or MaxIterations reached
        Application.StatusBar = "   [ Pass:" + CStr(ii) + " ]"
        Centroids = ComputeCentroids(DataRecords, ClusterIndexes, K)         ' calculate new centroids for each cluster
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)    ' assign each record in a cluster based on the new centroids
        ii = ii + 1
    Wend
    
    Range(OutputRange).Resize(K, J).Value = Centroids
    
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = ActiveSheet.Range("C5").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = ActiveSheet.Range("C6").Value
    
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(M, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
End Sub


Public Function EuclideanDistance(X As Variant, Y As Variant, num_obs As Integer) As Double
    Dim ii As Integer: ii = 1
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For ii = 1 To num_obs
        RunningSumSqr = RunningSumSqr + ((X(ii) - Y(ii)) ^ 2)
    Next ii
    
    EuclideanDistance = Sqr(RunningSumSqr)
End Function


'
' For each record DataRecords(ii) the result is calculated and placed in Cluster_Indexes(ii)
' This number is the cluster were we placed the record

Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Integer
    Dim K As Integer: K = UBound(Centroids, 1)      ' number of clusters
    Dim J As Integer: J = UBound(Centroids, 2)       ' number of columns (features)
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)     ' number of data records
    Dim idx() As Variant: ReDim idx(NumRecords) As Variant
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1

    Dim changeCounter As Integer: changeCounter = 0

    For ii = 1 To NumRecords      ' for all records
    
        Dim Dist_min As Double: Dist_min = 10000000
        Dim Dist As Double: Dist = 0
        For cc = 1 To K                         ' for all clusters / centroids
            Dist = EuclideanDistance(Application.Index(DataRecords, ii, 0), Application.Index(Centroids, cc, 0), J)
            If Dist < Dist_min Then
                idx(ii) = cc            ' this record is assigned to cluster cc when we find the min distance
                Dist_min = Dist
            End If
        Next cc                         ' check with next centroid / cluster
    
        If Not (IsEmpty(Cluster_Indexes)) Then
            If Not (Cluster_Indexes(ii) = idx(ii)) Then
                changeCounter = changeCounter + 1
            End If
        End If
    Next ii                ' next record
    FindClosestCentroid = changeCounter
    MsgBox changeCounter, vbOKOnly
    
    Cluster_Indexes = idx()                     ' update the clusters
End Function


'
' After we have assigned each data record in a cluster (based on the distance) we will calculate new centroid for eacj cluster
' For each data record that is in the cluster, we add and then average the features (columns)
'
Public Function ComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, NoOfClusters As Variant) As Variant
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)
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
            For cc = 1 To NumRecords          ' for every observation (data record)
                If ClusterIdx(cc) = ii Then   ' is this record part of (assigned to) cluster ii ??
                    Centroids(ii, bb) = Centroids(ii, bb) + DataRecords(cc, bb)             ' find the average for this observation
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
    ComputeCentroids = Centroids                           ' we have new centroids
    
End Function



