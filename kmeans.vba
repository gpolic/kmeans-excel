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
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)                        ' number of data records
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = wkSheet.Range("Clusters").Value
    Dim Centroids As Variant
    Dim counter As Integer
    Dim ClustersUpdated As Integer: ClustersUpdated = 1                 ' initially we have this >0, otherwise the loop wont execute
    
    Application.StatusBar = "   [ Initialize. ]"

Dim StartTime, SecondsElapsed As Double
StartTime = Timer
    
    Dim InitialCentroidsCalc As Variant
    InitialCentroidsCalc = ComputeInitialCentroidsCalc(DataRecords, NUMCLUSTERS)

    Application.StatusBar = "   [ Start..     ]"
    'Application.ScreenUpdating = False
    
    ' First pass. Assign each record(observation) in a initial cluster
    counter = FindClosestCentroid(DataRecords, InitialCentroidsCalc, ClusterIndexes)
    counter = 1   '  The result returned from FindClosestCentroid is not relevant right now
    
    While counter <= MaxIt And ClustersUpdated > 0                  ' We will process k-means until it is normalized or MaxIterations reached
        Application.StatusBar = "   [ Pass: " + CStr(counter) + "     ]"
        Centroids = ComputeCentroids(DataRecords, ClusterIndexes, NUMCLUSTERS)              ' calculate new centroids for each cluster
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)       ' assign each record in a cluster based on the new centroids
        counter = counter + 1
    Wend
    'Application.ScreenUpdating = True
    
    Dim ClusterOutputSht As String: ClusterOutputSht = wkSheet.Range("OutputSheet").Value   ' show the clusters assigned, on the output sheet/range
    Dim ClusterOutputRange As String: ClusterOutputRange = wkSheet.Range("OutputRange").Value
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(NUMRECORDS, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
    Call ShowResult(DataRecords, ClusterIndexes, Centroids, NUMCLUSTERS)
    
    Dim Distance As Double
    Distance = CalculateDistances(DataRecords, Centroids, ClusterIndexes)
    wkSheet.Range("C16").Value = Distance
    

Dim ExpO As Double
ExpO = CalculateExpectation(DataRecords, NUMCLUSTERS)

Dim Wk As Double
Wk = (1 / (2 * NUMRECORDS)) * Distance
'wkSheet.Range("C18").Value = ExpO
'wkSheet.Range("C19").Value = Wk
wkSheet.Range("C17").Value = ExpO - Log(Wk)
    
SecondsElapsed = Round(Timer - StartTime, 2)
'MsgBox "Time elapsed " & SecondsElapsed & " seconds", vbInformation
End Sub


Function CalculateDistances(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Variant
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)          ' number of data records
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(DataRecords, 2)          ' number of columns in each record
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = UBound(Centroids, 1)          ' number of clusters
    Dim DistanceSum As Double: DistanceSum = 0
    
    Dim DistanceInCluster() As Variant:   ReDim DistanceInCluster(NUMCLUSTERS)
    Dim clusterCounter, recordCounter, recordsInCluster As Integer
    
    For clusterCounter = 1 To NUMCLUSTERS
        recordsInCluster = 0
        For recordCounter = 1 To NUMRECORDS
            If Cluster_Indexes(recordCounter) = clusterCounter Then
                DistanceInCluster(clusterCounter) = DistanceInCluster(clusterCounter) + _
                    EuclideanDistance(Application.Index(Centroids, clusterCounter, 0), Application.Index(DataRecords, recordCounter, 0), NUMCOLUMNS)
                recordsInCluster = recordsInCluster + 1
            End If
        Next recordCounter
        'DistanceSum = DistanceSum + Sqr(DistanceInCluster(clusterCounter) / recordsInCluster)
        DistanceSum = DistanceSum + DistanceInCluster(clusterCounter)
    Next clusterCounter
    CalculateDistances = DistanceSum

End Function


Function CalculateExpectation(ByRef DataRecords As Variant, NUMCLUSTERS As Integer) As Double
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)          ' number of data records
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(DataRecords, 2)          ' number of columns in each record
    Dim Exp As Double
    
    Exp = Log((NUMRECORDS * NUMCOLUMNS) / 12) - ((2 / NUMCOLUMNS) * Log(NUMCLUSTERS))
    
    CalculateExpectation = Exp
End Function


' Choose initial centroids with KMeans++
'
Function ComputeInitialCentroidsCalc(ByRef DataRecords As Variant, NUMCLUSTERS As Integer) As Variant

    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)          ' number of data records
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(DataRecords, 2)          ' number of columns in each record
    Dim Taken() As Variant: ReDim Taken(NUMRECORDS)
    
    Dim InitialCentroidsCalc As Variant: ReDim InitialCentroidsCalc(NUMCLUSTERS, NUMCOLUMNS) As Variant
    Dim minDistSquared As Variant: ReDim minDistSquared(NUMRECORDS)
    Dim counter As Integer
    Dim CentroidsFound As Integer
    Dim dist As Double
    Dim preventLoop As Boolean: preventLoop = True

    Dim FirstCentroid As Variant: ReDim FirstCentroid(NUMCOLUMNS)
    Dim FirstCentroidIndex As Integer
    
    FirstCentroidIndex = Int(Rnd * NUMRECORDS) + 1         ' select first centroid by random from our data records
    
' Change the kmeans++ standard algorithm. We choose the first centroid with the mean values, not by random selection
' First Centroid - Choose the record that is closer to the mean
' ------------------------------------------------------------------
'    Dim colCounter As Integer
'    For colCounter = 1 To NUMCOLUMNS
'        For counter = 1 To NUMRECORDS
'            FirstCentroid(colCounter) = FirstCentroid(colCounter) + DataRecords(counter, colCounter)
'        Next counter
'        FirstCentroid(colCounter) = FirstCentroid(colCounter) / NUMRECORDS  ' find the mean
'    Next colCounter
'
'    Dim MinimumDistance As Double: MinimumDistance = 99999999
'    Dim MinRecord As Variant
'    Dim cc As Integer
'    For cc = 1 To NUMRECORDS          ' calculate distance to all records and select the record closer to the mean
'        dist = EuclideanDistance(Application.Index(DataRecords, cc, 0), FirstCentroid, NUMCOLUMNS)
'        If dist < MinimumDistance Then
'            FirstCentroidIndex = cc            ' the record with lowest distance to the means will be 1st centroid
'            MinimumDistance = dist
'        End If
'    Next cc                            ' check with next data record
' ------------------------------------------------------------------
    
    For counter = 1 To NUMCOLUMNS
        FirstCentroid(counter) = DataRecords(FirstCentroidIndex, counter)       ' put this data record in FirstCentroid
        InitialCentroidsCalc(1, counter) = FirstCentroid(counter)               ' and put it also in the array of results
    Next counter
    
    Taken(FirstCentroidIndex) = 1         ' mark point as Taken
    CentroidsFound = 1          ' we have one cluster center
    
    For counter = 1 To NUMRECORDS
        If Not counter = FirstCentroidIndex Then
            dist = EuclideanDistance(FirstCentroid, Application.Index(DataRecords, counter, 0), NUMCOLUMNS)
            minDistSquared(counter) = dist * dist
        End If
    Next counter

    ' main loop
Do While CentroidsFound < NUMCLUSTERS And preventLoop = True
    Dim distSqSum As Variant: distSqSum = 0
    
    For counter = 1 To NUMRECORDS   ' sum all the squared distances of the points not already taken
        If Not Taken(counter) = 1 Then
        distSqSum = distSqSum + minDistSquared(counter)
        End If
    Next counter

    Dim R As Variant                ' add one new point. each point is chosen with probability proportional to D(x)2
    R = Rnd * distSqSum

    Dim nextpoint As Integer        ' the index of the next point to be added as cluster center
    nextpoint = -1
    
    Dim sum As Variant
    
    For counter = 1 To NUMRECORDS   ' scan through the dist squared distances until sum > R
        If Not Taken(counter) = 1 Then
            sum = sum + minDistSquared(counter)
            If sum > R Then
                nextpoint = counter
                Exit For
            End If
        End If
    Next counter
    
    If nextpoint = -1 Then          ' if a new point was not found yet. just pick the last available data record
        For counter = NUMRECORDS To 1
            If Not Taken(counter) = 1 Then
                nextpoint = counter
            End If
        Next counter
    End If
    
    If nextpoint >= 0 Then                      ' we found the next cluster center !
        CentroidsFound = CentroidsFound + 1
        Taken(nextpoint) = 1                    ' mark the data record as Taken
        For counter = 1 To NUMCOLUMNS           ' copy the data in the array to be returned
            InitialCentroidsCalc(CentroidsFound, counter) = DataRecords(nextpoint, counter)
        Next counter
            
        If CentroidsFound < NUMCLUSTERS Then    ' need to find more centroids. we will adjust the minSqDistance
            For counter = 1 To NUMRECORDS
                If Not Taken(counter) = 1 Then
                    Dim dist2 As Variant
                    dist2 = EuclideanDistance(Application.Index(InitialCentroidsCalc, CentroidsFound, 0), Application.Index(DataRecords, counter, 0), NUMCOLUMNS)
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
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = UBound(Centroids, 1)     ' number of clusters
    Dim NUMCOLUMNS As Integer: NUMCOLUMNS = UBound(Centroids, 2)       ' number of columns (features)
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)     ' number of data records
    Dim idx() As Variant: ReDim idx(NUMRECORDS) As Variant
    Dim recordsCounter, clusterCounter As Integer
    Dim changeCounter As Integer: changeCounter = 0

    For recordsCounter = 1 To NUMRECORDS      ' for all records
    
        Dim MinimumDistance As Double: MinimumDistance = 99999999
        Dim MinCluster As Variant
        Dim dist As Double: dist = 0
        For clusterCounter = 1 To NUMCLUSTERS          ' calculate distance to all centroids and assign to the minimum distance cluster
            dist = EuclideanDistance(Application.Index(DataRecords, recordsCounter, 0), Application.Index(Centroids, clusterCounter, 0), NUMCOLUMNS)
            If dist < MinimumDistance Then
                MinCluster = clusterCounter            ' this record will be assigned to cluster MinCluster when we find the min distance
                MinimumDistance = dist
            End If
        Next clusterCounter                            ' check with next centroid / cluster
        idx(recordsCounter) = MinCluster               ' change the cluster index to the closest cluster
        
        If Not (IsEmpty(Cluster_Indexes)) Then         ' check what is the difference with our  previous centroids - if any
            If Not (Cluster_Indexes(recordsCounter) = idx(recordsCounter)) Then         ' the old cluster index is not the same as the new one
                changeCounter = changeCounter + 1      ' indicate that a record has moved to another cluster
            End If
        End If
    Next recordsCounter                ' next record
    FindClosestCentroid = changeCounter
    'MsgBox changeCounter, vbOKOnly
    
    Cluster_Indexes = idx()                     ' update the clusters
End Function



' Show the results in the Result sheet
'
Public Sub ShowResult(ByRef DataRecords As Variant, ByRef Cluster_Indexes As Variant, ByRef Centroids, NUMCLUSTERS As Integer)
    Dim resultSheet As Worksheet
    Dim lRowLast, lColLast As Integer
    Dim Rng As Range
    Dim ClusterObjects() As Variant: ReDim ClusterObjects(NUMCLUSTERS) As Variant
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)     ' number of data records
    
    ' clear the old data in Result sheet
    With resultSheet
        lRowLast = .UsedRange.Row + .UsedRange.Rows.Count - 1
        lColLast = .UsedRange.Column + .UsedRange.Columns.Count - 1
        Set Rng = .Range(.Range("B4"), .Cells(lRowLast, lColLast))
    End With
    Rng.ClearContents  ' delete contents without deleting the format
    
    Dim cluster As Integer
    For cluster = 1 To NUMCLUSTERS
        ClusterObjects(cluster) = 0  ' initialize Cluster object count
        resultSheet.Cells(4, 1 + cluster).Value = cluster
    Next cluster

    Dim counter As Integer
    For counter = 1 To NUMRECORDS
        ClusterObjects(Cluster_Indexes(counter)) = ClusterObjects(Cluster_Indexes(counter)) + 1  ' for every record in this cluster, increase the counter
    Next counter

    resultSheet.Range("B5").Resize(1, NUMCLUSTERS).Value = ClusterObjects
    resultSheet.Range("B9").Resize(UBound(Centroids, 1), UBound(Centroids, 2)).Value = Centroids         ' Show the final centroids in the results
    
End Sub


Public Function ComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, NoOfClusters As Variant) As Variant
    Dim NUMRECORDS As Integer: NUMRECORDS = UBound(DataRecords, 1)
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
            For cc = 1 To NUMRECORDS          ' for every observation (data record)
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

