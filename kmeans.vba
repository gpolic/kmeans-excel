Option Base 1
Option Explicit

Public Sub kmeans()
    Dim wkSheet As Worksheet
    Set wkSheet = ActiveWorkbook.Worksheets("Start")

    Dim MaxIt As Integer: MaxIt = wkSheet.Range("MaxIt").Value
    Dim DataSht As String: DataSht = wkSheet.Range("InputSheet").Value
    Dim DataRange As String: DataRange = wkSheet.Range("InputRange").Value
    Dim DataRecords As Variant: DataRecords = Worksheets(DataSht).Range(DataRange)      ' this holds all the data records
    Dim ClusterIndexes As Variant
    Dim J As Integer: J = UBound(DataRecords, 2)                        ' J features in each record
    Dim M As Integer: M = UBound(DataRecords, 1)                        ' M number of records
    Dim K As Integer: K = wkSheet.Range("Clusters").Value
    Dim Centroids As Variant
    Dim counter As Integer
    Dim ClustersUpdated As Integer: ClustersUpdated = 1                 ' initially we have this >0, otherwise the loop wont execute
    
    Application.StatusBar = "   [ Initialize. ]"

Dim StartTime, SecondsElapsed As Double
StartTime = Timer
    
    Dim InitialCentroidsCalc As Variant
    InitialCentroidsCalc = ComputeInitialCentroidsCalc(DataRecords, K)
    
    Application.StatusBar = "   [ Start..     ]"
    'Application.ScreenUpdating = False
    
    ' First pass. The result returned from FindClosestCentroid is not really needed now
    counter = FindClosestCentroid(DataRecords, InitialCentroidsCalc, ClusterIndexes)         ' assign each record(observation) in a initial cluster
    counter = 1
    
    While counter <= MaxIt And ClustersUpdated > 0                  ' We will process k-means until it is normalized or MaxIterations reached
        Application.StatusBar = "   [ Pass: " + CStr(counter) + "     ]"
        Centroids = ComputeCentroids(DataRecords, ClusterIndexes, K)                    ' calculate new centroids for each cluster
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)   ' assign each record in a cluster based on the new centroids
        counter = counter + 1
    Wend
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = wkSheet.Range("OutputSheet").Value        ' show the clusters assigned, on the output sheet/range
    Dim ClusterOutputRange As String: ClusterOutputRange = wkSheet.Range("OutputRange").Value
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(M, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
    Call ShowResult(DataRecords, ClusterIndexes, Centroids, K)
    
    Dim Distance As Double
    Distance = CalculateDistances(DataRecords, Centroids, ClusterIndexes)
    wkSheet.Range("C16").Value = Distance
    
SecondsElapsed = Round(Timer - StartTime, 2)
'MsgBox "Time elapsed " & SecondsElapsed & " seconds", vbInformation
End Sub


Function CalculateDistances(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Variant
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)          ' number of data records
    Dim NumColumns As Integer: NumColumns = UBound(DataRecords, 2)          ' number of columns in each record
    Dim NumClusters As Integer: NumClusters = UBound(Centroids, 1)          ' number of clusters
    Dim Distance, DistanceSum As Double
    
    Dim DistanceInCluster() As Variant:   ReDim DistanceInCluster(NumClusters)
    Dim clusterCounter, recordCounter, recordsInCluster As Integer
    
    For clusterCounter = 1 To NumClusters
        recordsInCluster = 0
        For recordCounter = 1 To NumRecords
            If Cluster_Indexes(recordCounter) = clusterCounter Then
                DistanceInCluster(clusterCounter) = DistanceInCluster(clusterCounter) + _
                    EuclideanDistance(Application.Index(Centroids, clusterCounter, 0), Application.Index(DataRecords, recordCounter, 0), NumColumns)
                recordsInCluster = recordsInCluster + 1
            End If
        Next recordCounter
        'DistanceSum = DistanceSum + Sqr(DistanceInCluster(clusterCounter) / recordsInCluster)
        DistanceSum = DistanceSum + DistanceInCluster(clusterCounter)
    Next clusterCounter
    CalculateDistances = DistanceSum

End Function


' Inspired from org.apache.commons.math3.ml.clustering.KMeansPlusPlusClusterer
'
Function ComputeInitialCentroidsCalc(ByRef DataRecords As Variant, NumClusters As Integer) As Variant

    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)          ' number of data records
    Dim NumColumns As Integer: NumColumns = UBound(DataRecords, 2)          ' number of columns in each record
    Dim Taken() As Variant: ReDim Taken(NumRecords)
    
    Dim InitialCentroidsCalc As Variant: ReDim InitialCentroidsCalc(NumClusters, NumColumns) As Variant
    Dim minDistSquared As Variant: ReDim minDistSquared(NumRecords)
    Dim counter As Integer
    Dim CentroidsFound As Integer
    Dim dist As Double
    Dim preventLoop As Boolean: preventLoop = True

    Dim FirstCentroid As Variant: ReDim FirstCentroid(NumColumns)
    Dim firstnum As Integer
    
    'firstnum = Int(Rnd * NumRecords) + 1                ' select first centroid by random from our data records
    
' new First Centroid - Choose the record that is closer to the mean
' Change the kmeans++ standard algorithm. We take the first centroid by the means, not by random selection
' ------------------------------------------------------------------
    Dim colCounter As Integer
    For colCounter = 1 To NumColumns
        For counter = 1 To NumRecords
            FirstCentroid(colCounter) = FirstCentroid(colCounter) + DataRecords(counter, colCounter)
        Next counter
        FirstCentroid(colCounter) = FirstCentroid(colCounter) / NumRecords  ' find the average values
    Next colCounter
    
    Dim MinimumDistance As Double: MinimumDistance = 99999999
    Dim MinRecord As Variant
    Dim cc As Integer
    For cc = 1 To NumRecords          ' calculate distance to all records and select the record closer to the mean
        dist = EuclideanDistance(Application.Index(DataRecords, cc, 0), FirstCentroid, NumColumns)
        If dist < MinimumDistance Then
            firstnum = cc            ' the record with lowest distance to the means will be 1st centroid
            MinimumDistance = dist
        End If
    Next cc                            ' check with next data record
' ------------------------------------------------------------------
    
    For counter = 1 To NumColumns
        FirstCentroid(counter) = DataRecords(firstnum, counter)     ' put this data record in FirstCentroid
        InitialCentroidsCalc(1, counter) = FirstCentroid(counter)   ' and put it also in the array to be returned
    Next counter
    
    Taken(firstnum) = 1         ' mark point as Taken
    CentroidsFound = 1          ' we have one cluster center
    
    For counter = 1 To NumRecords
        If Not counter = firstnum Then
            dist = EuclideanDistance(FirstCentroid, Application.Index(DataRecords, counter, 0), NumColumns)
            minDistSquared(counter) = dist * dist
        End If
    Next counter

    ' main loop
Do While CentroidsFound < NumClusters And preventLoop = True
    Dim distSqSum As Variant: distSqSum = 0
    
    For counter = 1 To NumRecords   ' sum all the squared distances of the points not already taken
        If Not Taken(counter) = 1 Then
        distSqSum = distSqSum + minDistSquared(counter)
        End If
    Next counter

    Dim R As Variant                ' add one new point. each point is chosen with probability proportional to D(x)2
    R = Rnd * distSqSum

    Dim nextpoint As Integer        ' the index of the next point to be added as cluster center
    nextpoint = -1
    
    Dim sum As Variant
    
    For counter = 1 To NumRecords   ' scan through the dist squared distances until sum > R
        If Not Taken(counter) = 1 Then
            sum = sum + minDistSquared(counter)
            If sum > R Then
                nextpoint = counter
                Exit For
            End If
        End If
    Next counter
    
    If nextpoint = -1 Then          ' if a new point was not found yet. just pick the last available data record
        For counter = NumRecords To 1
            If Not Taken(counter) = 1 Then
                nextpoint = counter
            End If
        Next counter
    End If
    
    If nextpoint >= 0 Then                      ' we found the next cluster center !
        CentroidsFound = CentroidsFound + 1
        Taken(nextpoint) = 1                    ' mark the data record as Taken
        For counter = 1 To NumColumns           ' copy the data in the array to be returned
            InitialCentroidsCalc(CentroidsFound, counter) = DataRecords(nextpoint, counter)
        Next counter
            
        If CentroidsFound < NumClusters Then    ' need to find more centroids. we will adjust the minSqDistance
            For counter = 1 To NumRecords
                If Not Taken(counter) = 1 Then
                    Dim dist2 As Variant
                    dist2 = EuclideanDistance(Application.Index(InitialCentroidsCalc, CentroidsFound, 0), Application.Index(DataRecords, counter, 0), NumColumns)
                                                ' find the distance to the new centroid
                    Dim d2 As Variant
                    d2 = dist2 * dist2
                    
                    If d2 < minDistSquared(counter) Then        ' if the distance to the new centroid is lower than the previous then use it
                        minDistSquared(counter) = d2
                    End If
                End If
            Next counter
        End If
    Else                        ' there is no cluster center found
        preventLoop = False     ' make sure that the while loop can terminate
    End If
Loop

    ComputeInitialCentroidsCalc = InitialCentroidsCalc

End Function
    


Public Function EuclideanDistance(X As Variant, Y As Variant, NumberOfObservations As Integer) As Double
    Dim counter As Integer
    Dim RunningSumSqr As Double: RunningSumSqr = 0
    
    For counter = 1 To NumberOfObservations
        RunningSumSqr = RunningSumSqr + ((X(counter) - Y(counter)) ^ 2)
    Next counter
    
    EuclideanDistance = Sqr(RunningSumSqr)
End Function


'
' For each record in Data Records, find the closest Centroid (cluster)
' The result is calculated and placed in Cluster_Indexes()
' This number is the cluster were we placed the record. This is more effective than creating new Arrays with Clusters
'
Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Integer
    Dim NumClusters As Integer: NumClusters = UBound(Centroids, 1)     ' number of clusters
    Dim NumColumns As Integer: NumColumns = UBound(Centroids, 2)       ' number of columns (features)
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)     ' number of data records
    Dim idx() As Variant: ReDim idx(NumRecords) As Variant
    Dim counterR As Integer
    Dim cc As Integer

    Dim changeCounter As Integer: changeCounter = 0

    For counterR = 1 To NumRecords      ' for all records
    
        Dim MinimumDistance As Double: MinimumDistance = 99999999
        Dim MinCluster As Variant
        Dim dist As Double: dist = 0
        For cc = 1 To NumClusters          ' calculate distance to all centroids and assign to the minimum distance cluster
            dist = EuclideanDistance(Application.Index(DataRecords, counterR, 0), Application.Index(Centroids, cc, 0), NumColumns)
            If dist < MinimumDistance Then
                MinCluster = cc            ' this record will be assigned to cluster MinCluster when we find the min distance
                MinimumDistance = dist
            End If
        Next cc                            ' check with next centroid / cluster
        idx(counterR) = MinCluster
        
        If Not (IsEmpty(Cluster_Indexes)) Then          ' check what is the difference with our  previous centroids - if any
            If Not (Cluster_Indexes(counterR) = idx(counterR)) Then ' the old cluster index is not the same as the new one
                changeCounter = changeCounter + 1       ' indicate that a data record has moved to another cluster
            End If
        End If
    Next counterR                ' next record
    FindClosestCentroid = changeCounter
    'MsgBox changeCounter, vbOKOnly
    
    Cluster_Indexes = idx()                     ' update the clusters
End Function



' Show the results in the Result sheet
'
Public Sub ShowResult(ByRef DataRecords As Variant, ByRef Cluster_Indexes As Variant, ByRef Centroids, NumClusters As Integer)
    Dim resultSheet As Worksheet
    Dim lRowLast, lColLast As Integer
    Dim Rng As Range
    Dim ClusterObjects() As Variant: ReDim ClusterObjects(NumClusters) As Variant
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Dim NumRecords As Integer: NumRecords = UBound(DataRecords, 1)     ' number of data records
    
    ' clear the old data in Result sheet
    With resultSheet
        lRowLast = .UsedRange.Row + .UsedRange.Rows.Count - 1
        lColLast = .UsedRange.Column + .UsedRange.Columns.Count - 1
        Set Rng = .Range(.Range("B4"), .Cells(lRowLast, lColLast))
    End With
    Rng.ClearContents  ' delete contents without deleting the format
    
    Dim cluster As Integer
    For cluster = 1 To NumClusters
        ClusterObjects(cluster) = 0  ' initialize Cluster object count
        resultSheet.Cells(4, 1 + cluster).Value = cluster
    Next cluster

    Dim counter As Integer
    For counter = 1 To NumRecords
        ClusterObjects(Cluster_Indexes(counter)) = ClusterObjects(Cluster_Indexes(counter)) + 1  ' for every record in this cluster, increase the counter
    Next counter

    resultSheet.Range("B5").Resize(1, NumClusters).Value = ClusterObjects
    resultSheet.Range("B9").Resize(UBound(Centroids, 1), UBound(Centroids, 2)).Value = Centroids         ' Show the final centroids in the results
    
End Sub


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


