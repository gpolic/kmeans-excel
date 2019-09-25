# k-means

K-means is an algorithm for cluster analysis (clustering). It is the process of partitioning a set of data into related groups / clusters.
K-means clustering is useful for Data Mining and Business Intelligence.

Here is k-means in plain English:
https://rayli.net/blog/data/top-10-data-mining-algorithms-in-plain-english/#2_k-means

This script is based on the work of bquanttrading. His blog on market modelling and market analytics:

https://asmquantmacro.com


# What does it do

k-means will classify each record in your data, placing it into a group (cluster). You do not need to specify the properties of each group, k-means will decide for the groups. However, usually we need to provide the number of groups that we want in the output.

The records in the same cluster are similar to each other. Records in different clusters are dissimilar.

Each row of your Excel data, should be a record/observation with one or more features. Each column is a feature in the observation.

As an example, here is a data set with the height and weight of 25,000 children in Hong Kong : http://socr.ucla.edu/docs/resources/SOCR_Data/SOCR_Data_Dinov_020108_HeightsWeights.html

Each row in the data represents a person. Each column is a feature of the person.

Currently the script works _only_ with numerical data.



# How does it work

* Enter your data in a new Excel worksheet
* Enter the name of the worksheet in cell C4, and the range of the data at C5
* Enter the worksheet for the output to be placed, at C6 (you can use the one where your data is)
* Enter the cell where the output will be updated at C7
* Number of groups in your data at C8
* Click the button to start
* Check the Result

If you do not know the number of clusters/groups contained in your data, try different values for example 1 up to 10. 
Execute the script several times and observe the GAP figure. 
At the point where GAP reaches its maximum value, it indicates that the number of clusters is efficient for this data set.

As an example, changing the number of clusters and calculating with the IRIS data set, GAP will maximize when we have 3 clusters.
	
The original paper that describes the GAP calculation: https://web.stanford.edu/~hastie/Papers/gap.pdf
	
# The results

The result is a number assigned on each record, that indicates the group/cluster the record belongs to.

The Result sheet contains information on the clusters, along with the cluster centers. 


# Performance

When the "Distance" value is minimized, it indicates the output accuracy is higher. 

Execute the algorithm several times to find the best results.

The script will stop execution when the clusters are normalized or when the maximum iterations are reached (whichever comes first).  You can increase the number of iterations for better results.

Unfortunately Excel VBA runs on a single thread, therefore it does not take full advantage of your current CPU's

# Why is this different?

The script calculates the initial centroids using _k-means++_ algorithm. You do not have to provide the initial centroids.
It also provides an indication of the number of groups contained in the data, using the GAP calculation.

# More info

This is implementing David Arthur and Sergei Vassilvitski k-means++ algorithm, which chooses the initial centroids.
https://theory.stanford.edu/~sergei/papers/kMeansPP-soda.pdf

The example dataset provided in kmeans.xlsx is _IRIS_ from UC Irvine Machine Learning Repository: https://archive.ics.uci.edu/ml/datasets.html


