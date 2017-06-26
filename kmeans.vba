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
        Centroids = NewComputeCentroids(DataRecords, ClusterIndexes, K)         ' calculate new centroids for each cluster
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)    ' assign each record in a cluster based on the new centroids
        ii = ii + 1
    Wend
    
    Range(OutputRange).Resize(K, J).Value = Centroids
    
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = ActiveSheet.Range("C5").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = ActiveSheet.Range("C6").Value
    
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(M, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
End Sub


Public Function EuclideanDistance(X As Variant, Y As Variant, NumberOfObservations As Integer) As Double
    Dim ii As Integer: ii = 1
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For ii = 1 To NumberOfObservations
        RunningSumSqr = RunningSumSqr + ((X(ii) - Y(ii)) ^ 2)
    Next ii
    
    EuclideanDistance = Sqr(RunningSumSqr)
End Function


'
' For each record DataRecords(ii) the result is calculated and placed in Cluster_Indexes(ii)
' This number is the cluster were we placed the record
'
Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Integer
    Dim K As Integer: K = UBound(Centroids, 1)      ' number of clusters
    Dim J As Integer: J = UBound(Centroids, 2)       ' number of columns (features)
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)     ' number of data records
    Dim idx() As Variant: ReDim idx(NumRecords) As Variant
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1

    Dim changeCounter As Integer: changeCounter = 0

    For ii = 1 To NumRecords      ' for all records
    
        Dim MinimumDistance As Double: MinimumDistance = 10000000
        Dim MinCluster As Variant
        Dim Dist As Double: Dist = 0
        For cc = 1 To K                    ' meassure distance to all centroids and assign to the minimum distance cluster
            Dist = EuclideanDistance(Application.Index(DataRecords, ii, 0), Application.Index(Centroids, cc, 0), J)
            If Dist < MinimumDistance Then
                MinCluster = cc            ' this record will be assigned to cluster MinCluster when we find the min distance
                MinimumDistance = Dist
            End If
        Next cc                         ' check with next centroid / cluster
        idx(ii) = MinCluster
        
        If Not (IsEmpty(Cluster_Indexes)) Then          ' check what is the difference with our  previous centroids - if any
            If Not (Cluster_Indexes(ii) = idx(ii)) Then ' the old cluster index is not the same as the new one
                changeCounter = changeCounter + 1       ' that means this data record has moved to another cluster
            End If
        End If
    Next ii                ' next record
    FindClosestCentroid = changeCounter
    'MsgBox changeCounter, vbOKOnly
    
    Cluster_Indexes = idx()                     ' update the clusters
End Function


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

'
' After we have assigned each data record in a cluster (based on the distance) we will calculate new centroid for eacj cluster
' For each data record that is in the cluster, we add and then average the features (columns)
'
Public Function NewComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, NoOfClusters As Variant) As Variant
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)
    Dim RecordSize As Integer: RecordSize = UBound(DataRecords, 2)
    Dim ii As Integer: ii = 1
    Dim cc As Integer: cc = 1
    Dim bb As Integer: bb = 1
    Dim counter As Integer: counter = 0
    Dim tempSum() As Variant: ReDim tempSum(NoOfClusters, RecordSize) As Variant
    Dim Centroids() As Variant: ReDim Centroids(NoOfClusters, RecordSize) As Variant
    
'    For ii = 1 To NoOfClusters       ' for all clusters
'        For bb = 1 To RecordSize     ' for all features(columns)
'            counter = 0
'            For cc = 1 To NumRecords          ' for every observation (data record)
'                If ClusterIdx(cc) = ii Then   ' is this record part of (assigned to) cluster ii ??
'                    Centroids(ii, bb) = Centroids(ii, bb) + DataRecords(cc, bb)             ' find the average for this observation
'                    counter = counter + 1
'                End If
'            Next cc
'            If counter > 0 Then
'                Centroids(ii, bb) = Centroids(ii, bb) / counter    ' compute the new centroid averaging all records in the cluster
'            Else
 '               Centroids(ii, bb) = 0
 '           End If
'        Next bb
'    Next ii
    
    Dim Counters() As Integer: ReDim Counters(NoOfClusters, RecordSize) As Integer  ' count the data records in a cluster
    Dim ClusterNumber As Integer
    
    For cc = 1 To NumRecords          ' for every observation (data record)
        ClusterNumber = ClusterIdx(cc)
        
        For bb = 1 To RecordSize       ' for every feature (column)
            Centroids(ClusterNumber, bb) = Centroids(ClusterNumber, bb) + DataRecords(cc, bb)     ' find the sum of all observations
            Counters(ClusterNumber, bb) = Counters(ClusterNumber, bb) + 1
        Next bb
    Next cc
    
    For ii = 1 To NoOfClusters
        For bb = 1 To RecordSize
            If (Counters(ii, bb) > 0) Then
                Centroids(ii, bb) = Centroids(ii, bb) / Counters(ii, bb)
            End If
        Next bb
    Next ii
    
    NewComputeCentroids = Centroids                           ' we have new centroids
    
End Function





