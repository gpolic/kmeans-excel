Option Base 1
Option Explicit

Public Sub kmeans()
    Dim wkSheet As Worksheet
    Set wkSheet = ActiveWorkbook.Worksheets("Start")

    Dim MaximumIterations As Integer: MaximumIterations = wkSheet.Range("MaximumIterations").Value
    Dim DataSht As String: DataSht = wkSheet.Range("InputSheet").Value
    Dim DataRange As String: DataRange = wkSheet.Range("InputRange").Value
    Dim DataRecords As Variant: DataRecords = Worksheets(DataSht).Range(DataRange)
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = wkSheet.Range("Clusters").Value
    Dim ClusterIndexes As Variant, Centroids As Variant, InitialCentroidsCalc As Variant
    Dim ClustersUpdated As Integer, counter As Integer: counter = 1
    Dim StartTime As Double
    
    StartTime = Timer
    Application.StatusBar = "   [ Initialize. ]"
    
    ' initialize centroids with kmeans++ method
    InitialCentroidsCalc = ComputeInitialCentroidsCalc(DataRecords, NUMCLUSTERS)

    Application.StatusBar = "   [ Start..     ]"
    'Application.ScreenUpdating = False
    
    ' First pass. Assign each record(observation) in a initial cluster. ClusterIndexes is updated
    ClustersUpdated = FindClosestCentroid(DataRecords, InitialCentroidsCalc, ClusterIndexes)
    
    '  The result returned from FindClosestCentroid is not relevant right now
    ClustersUpdated = 1
    
    ' We will process k-means until it is normalized or MaximumIterationserations reached
    While counter <= MaximumIterations And ClustersUpdated > 0
        Application.StatusBar = "   [ Pass: " + CStr(counter) + "     ]"
        
        ' calculate new centroids for each cluster
        Centroids = ComputeCentroids(DataRecords, ClusterIndexes, NUMCLUSTERS)
        
        ' assign each record in a cluster based on the new centroids
        ClustersUpdated = FindClosestCentroid(DataRecords, Centroids, ClusterIndexes)
        counter = counter + 1
    Wend
    
    Application.StatusBar = "   Completed after " + CStr(counter - 1) + " iterations"
    'Application.ScreenUpdating = True
    
    ' show the clusters assigned in the output sheet/range
    Dim ClusterOutputSht As String: ClusterOutputSht = wkSheet.Range("OutputSheet").Value
    Dim ClusterOutputRange As String: ClusterOutputRange = wkSheet.Range("OutputRange").Value
    Worksheets(ClusterOutputSht).Range(ClusterOutputRange).Resize(NUMBER_OF_RECORDS, 1).Value = WorksheetFunction.Transpose(ClusterIndexes)
    
    Call ShowResult(DataRecords, ClusterIndexes, Centroids, NUMCLUSTERS)
    
    ' show more results
    Dim Distance As Double, ExpO As Double, Wk As Double
    
    Distance = CalculateDistances(DataRecords, Centroids, ClusterIndexes)
    ExpO = CalculateExpectation(DataRecords, NUMCLUSTERS)
    Wk = (1 / (2 * NUMBER_OF_RECORDS)) * Distance
    
    wkSheet.Range("C16").Value = Distance
    wkSheet.Range("C17").Value = ExpO - Log(Wk)
    'wkSheet.Range("C18").Value = ExpO
    'wkSheet.Range("C19").Value = Wk
    
    'MsgBox "Time elapsed " & Round(Timer - StartTime, 2) & " seconds", vbInformation
End Sub


Function CalculateDistances(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Variant
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim NUMBER_OF_COLUMNS As Integer: NUMBER_OF_COLUMNS = UBound(DataRecords, 2)
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = UBound(Centroids, 1)
    Dim DistanceInCluster() As Variant:   ReDim DistanceInCluster(NUMCLUSTERS)
    Dim clusterCounter, recordCounter, recordsInCluster As Integer
    Dim DistanceSum As Double: DistanceSum = 0
    
    For clusterCounter = 1 To NUMCLUSTERS
            
            recordsInCluster = 0
            For recordCounter = 1 To NUMBER_OF_RECORDS
            
                If Cluster_Indexes(recordCounter) = clusterCounter Then
                    DistanceInCluster(clusterCounter) = DistanceInCluster(clusterCounter) + _
                        EuclideanDistance(Application.Index(Centroids, clusterCounter, 0), Application.Index(DataRecords, recordCounter, 0), NUMBER_OF_COLUMNS)
                    recordsInCluster = recordsInCluster + 1
                End If
                
            Next recordCounter
            
            'DistanceSum = DistanceSum + Sqr(DistanceInCluster(clusterCounter) / recordsInCluster)
            DistanceSum = DistanceSum + DistanceInCluster(clusterCounter)
    Next clusterCounter
    
    CalculateDistances = DistanceSum
End Function


Function CalculateExpectation(ByRef DataRecords As Variant, NUMCLUSTERS As Integer) As Double
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim NUMBER_OF_COLUMNS As Integer: NUMBER_OF_COLUMNS = UBound(DataRecords, 2)
    
    CalculateExpectation = Log((NUMBER_OF_RECORDS * NUMBER_OF_COLUMNS) / 12) - ((2 / NUMBER_OF_COLUMNS) * Log(NUMCLUSTERS))
End Function


' Select initial centroids
'
Function ComputeInitialCentroidsCalc(ByRef DataRecords As Variant, NUMCLUSTERS As Integer) As Variant

    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim NUMBER_OF_COLUMNS As Integer: NUMBER_OF_COLUMNS = UBound(DataRecords, 2)
    Dim Taken() As Variant: ReDim Taken(NUMBER_OF_RECORDS)
    
    Dim InitialCentroidsCalc As Variant: ReDim InitialCentroidsCalc(NUMCLUSTERS, NUMBER_OF_COLUMNS) As Variant
    Dim minDistSquared As Variant: ReDim minDistSquared(NUMBER_OF_RECORDS)
    Dim counter As Integer, CentroidsFound As Integer, FirstCentroidIndex As Integer
    Dim dist As Double
    Dim preventLoop As Boolean: preventLoop = True
    Dim FirstCentroid As Variant: ReDim FirstCentroid(NUMBER_OF_COLUMNS)
   
    
    FirstCentroidIndex = Int(Rnd * NUMBER_OF_RECORDS) + 1         ' The first centroid is random !
    
' Change the kmeans++ standard algorithm. We choose the first centroid with the mean values, not by random selection
' First Centroid - Choose the record that is closer to the mean
' ------------------------------------------------------------------
'    Dim colCounter As Integer
'    For colCounter = 1 To NUMBER_OF_COLUMNS
'        For counter = 1 To NUMBER_OF_RECORDS
'            FirstCentroid(colCounter) = FirstCentroid(colCounter) + DataRecords(counter, colCounter)
'        Next counter
'        FirstCentroid(colCounter) = FirstCentroid(colCounter) / NUMBER_OF_RECORDS  ' find the mean
'    Next colCounter
'
'    Dim MinimumDistance As Double: MinimumDistance = 99999999
'    Dim MinRecord As Variant
'    Dim recordNumber As Integer
'    For recordNumber = 1 To NUMBER_OF_RECORDS          ' calculate distance to all records and select the record closer to the mean
'        dist = EuclideanDistance(Application.Index(DataRecords, recordNumber, 0), FirstCentroid, NUMBER_OF_COLUMNS)
'        If dist < MinimumDistance Then
'            FirstCentroidIndex = recordNumber            ' the record with lowest distance to the means will be 1st centroid
'            MinimumDistance = dist
'        End If
'    Next recordNumber                            ' check with next data record
' ------------------------------------------------------------------
    
    For counter = 1 To NUMBER_OF_COLUMNS
        ' put this data record in FirstCentroid
        FirstCentroid(counter) = DataRecords(FirstCentroidIndex, counter)
        
        ' and put it also in the array of results
        InitialCentroidsCalc(1, counter) = FirstCentroid(counter)
    Next counter
    
    ' mark point as Taken. We have one cluster center
    Taken(FirstCentroidIndex) = 1
    CentroidsFound = 1
    
    For counter = 1 To NUMBER_OF_RECORDS
    
        If Not counter = FirstCentroidIndex Then
            dist = EuclideanDistance(FirstCentroid, Application.Index(DataRecords, counter, 0), NUMBER_OF_COLUMNS)
            minDistSquared(counter) = dist * dist
        End If
        
    Next counter

    ' main loop
    Do While CentroidsFound < NUMCLUSTERS And preventLoop = True
        
            ' sum all the squared distances of the points not already taken
            Dim distSqSum As Double: distSqSum = 0
            For counter = 1 To NUMBER_OF_RECORDS
            
                If Not Taken(counter) = 1 Then
                distSqSum = distSqSum + minDistSquared(counter)
                End If
                
            Next counter
        
            ' add one new point. each point is chosen with probability proportional to D(x)2
            Dim R As Double
            R = Rnd * distSqSum
        
            ' the index of the next point to be added as cluster center
            Dim nextpoint As Integer
            nextpoint = -1
            
    
             ' scan through the dist squared distances until sum > R
            Dim sum As Double: sum = 0
            For counter = 1 To NUMBER_OF_RECORDS
            
                If Not Taken(counter) = 1 Then
                    sum = sum + minDistSquared(counter)
                    
                    If sum > R Then
                        nextpoint = counter
                        Exit For
                    End If
                    
                End If
                
            Next counter
            
            ' if a new point was not found yet, just pick the last available data record
            If nextpoint = -1 Then
                For counter = NUMBER_OF_RECORDS To 1 Step -1
                
                    If Not Taken(counter) = 1 Then
                        nextpoint = counter
                    End If
                    
                Next counter
            End If
            
            If nextpoint >= 0 Then
            
                ' we found the next cluster center! Mark the data record as Taken
                CentroidsFound = CentroidsFound + 1
                Taken(nextpoint) = 1
                
                ' copy the data in the array to our result
                For counter = 1 To NUMBER_OF_COLUMNS
                    InitialCentroidsCalc(CentroidsFound, counter) = DataRecords(nextpoint, counter)
                Next counter
                
                ' need to find more centroids. we will adjust the minSqDistance
                If CentroidsFound < NUMCLUSTERS Then
                
                    For counter = 1 To NUMBER_OF_RECORDS
                    
                        If Not Taken(counter) = 1 Then
                        
                            ' find the distance to the new centroid
                            Dim dista As Double, distSquared As Double
                            
                            dista = EuclideanDistance(Application.Index(InitialCentroidsCalc, CentroidsFound, 0), Application.Index(DataRecords, counter, 0), NUMBER_OF_COLUMNS)
                            distSquared = dista * dista
                            
                            ' if the distance to the new centroid is lower than the previous, then use it
                            If distSquared < minDistSquared(counter) Then
                                minDistSquared(counter) = distSquared
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



' For each record in Data Records, find the closest Centroid (cluster)
' The result is calculated and placed in Cluster_Indexes()
' This number is the cluster were we placed the record. This is more effective than creating new Arrays with Clusters
'
Public Function FindClosestCentroid(ByRef DataRecords As Variant, ByRef Centroids As Variant, ByRef Cluster_Indexes As Variant) As Integer
    Dim NUMCLUSTERS As Integer: NUMCLUSTERS = UBound(Centroids, 1)
    Dim NUMBER_OF_COLUMNS As Integer: NUMBER_OF_COLUMNS = UBound(Centroids, 2)
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim idx() As Variant: ReDim idx(NUMBER_OF_RECORDS) As Variant
    Dim recordsCounter As Integer, clusterCounter As Integer
    Dim changeCounter As Integer: changeCounter = 0

    For recordsCounter = 1 To NUMBER_OF_RECORDS
    
            Dim MinimumDistance As Double: MinimumDistance = 99999999
            Dim MinCluster As Integer
            Dim dist As Double: dist = 0
            
            ' calculate distance to all centroids and assign to the minimum distance cluster
            For clusterCounter = 1 To NUMCLUSTERS
                dist = EuclideanDistance(Application.Index(DataRecords, recordsCounter, 0), Application.Index(Centroids, clusterCounter, 0), NUMBER_OF_COLUMNS)
                If dist < MinimumDistance Then
                
                     ' this record will be assigned to cluster MinCluster when we find the min distance
                    MinCluster = clusterCounter
                    MinimumDistance = dist
                End If
            Next clusterCounter
            
            ' change the cluster index to the closest cluster
            idx(recordsCounter) = MinCluster
            
            ' During the first run Cluster Indexes is Empty
            If Not (IsEmpty(Cluster_Indexes)) Then
                
                ' If the old cluster index is not the same as the new one
                If Not (Cluster_Indexes(recordsCounter) = idx(recordsCounter)) Then
                
                    ' indicate that a change occured
                    changeCounter = changeCounter + 1
                End If
                
            End If
        
    Next recordsCounter                ' next record
    
    FindClosestCentroid = changeCounter
    
    ' update the clusters
    Cluster_Indexes = idx()
End Function



' Show the results in the Result sheet
'
Public Sub ShowResult(ByRef DataRecords As Variant, ByRef Cluster_Indexes As Variant, ByRef Centroids, NUMCLUSTERS As Integer)
    Dim resultSheet As Worksheet
    Dim lRowLast As Integer, lColLast As Integer, counter As Integer
    Dim Rng As Range
    Dim ClusterObjects() As Variant: ReDim ClusterObjects(NUMCLUSTERS) As Variant
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    
    Set resultSheet = ActiveWorkbook.Worksheets("Result")

    
    ' clear the old data in Result sheet
    With resultSheet
        lRowLast = .UsedRange.Row + .UsedRange.Rows.Count - 1
        lColLast = .UsedRange.Column + .UsedRange.Columns.Count - 1
        Set Rng = .Range(.Range("B4"), .Cells(lRowLast, lColLast))
    End With
    Rng.ClearContents
    
    ' initialize Cluster object count
    For counter = 1 To NUMCLUSTERS
        ClusterObjects(counter) = 0
        resultSheet.Cells(4, 1 + counter).Value = counter
    Next counter

    ' for every record in this cluster, increase the counter
    For counter = 1 To NUMBER_OF_RECORDS
        ClusterObjects(Cluster_Indexes(counter)) = ClusterObjects(Cluster_Indexes(counter)) + 1
    Next counter

    ' Show the final centroids in the results
    resultSheet.Range("B5").Resize(1, NUMCLUSTERS).Value = ClusterObjects
    resultSheet.Range("B9").Resize(UBound(Centroids, 1), UBound(Centroids, 2)).Value = Centroids
    
End Sub


' This will sum all the records in a cluster, and average the values. The calculated averages will form the new Centroids
'
Public Function ComputeCentroids(DataRecords As Variant, ClusterIdx As Variant, Number_Of_Clusters As Integer) As Variant
    Dim NUMBER_OF_RECORDS As Integer: NUMBER_OF_RECORDS = UBound(DataRecords, 1)
    Dim NUMBER_OF_FEATURES As Integer: NUMBER_OF_FEATURES = UBound(DataRecords, 2)
    Dim clusterNumber As Integer, columnNumber As Integer, recordNumber As Integer, counter As Integer
    Dim tempSum() As Variant: ReDim tempSum(Number_Of_Clusters, NUMBER_OF_FEATURES) As Variant
    Dim Centroids() As Variant: ReDim Centroids(Number_Of_Clusters, NUMBER_OF_FEATURES) As Variant
    
    For clusterNumber = 1 To Number_Of_Clusters
    
            For columnNumber = 1 To NUMBER_OF_FEATURES
            
                    counter = 0
                    For recordNumber = 1 To NUMBER_OF_RECORDS
                        If ClusterIdx(recordNumber) = clusterNumber Then
                            
                            ' if this record is part of the cluster then add
                            Centroids(clusterNumber, columnNumber) = Centroids(clusterNumber, columnNumber) + DataRecords(recordNumber, columnNumber)
                            counter = counter + 1
                        End If
                    Next recordNumber
                    
                    If counter > 0 Then
                        
                        ' compute the new centroid averaging all records in the cluster
                        Centroids(clusterNumber, columnNumber) = Centroids(clusterNumber, columnNumber) / counter
                    Else
                        Centroids(clusterNumber, columnNumber) = 0
                    End If
                    
            Next columnNumber
            
    Next clusterNumber
    
    ComputeCentroids = Centroids
End Function




