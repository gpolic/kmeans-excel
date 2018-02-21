# k-means

K-means is an algorithm for cluster analysis (clustering). It is the process of partitioning a set of data into related groups / clusters.
K-means clustering is useful for Data Mining and Business Intelligence.

Here is k-means in plain English:
https://rayli.net/blog/data/top-10-data-mining-algorithms-in-plain-english/#2_k-means

# What does it do
This is a k-means clustering implementation in VBA. The script is based on David Arthur and Sergei Vassilvitski k-means++, which chooses the initial centroids.

Currently the script works _only_ with numerical data.

Each row of your excel data should be a record/observation with one or more features. Each column is a feature in the observation.

As an example, this is a data set with the height and weight of 25,000 children in Hong Kong : http://socr.ucla.edu/docs/resources/SOCR_Data/SOCR_Data_Dinov_020108_HeightsWeights.html

k-means will assign a number on each record indicating which groups it belongs to.

You only need to provide the number of clusters/groups you want.


# How does it work

	Enter your data in a new Excel worksheet. 
	Enter the name of the worksheet in cell C4, and the range of the data at C5
	Enter the worksheet for the results to be placed at C6 (you can use the one where your data is)
	Enter the cell were the result will be updated at C7
	Number of groups in your data at C8
	Click the button to start.

# The results

The result is a number assigned on each record, that indicates the group/cluster the record belongs to.

The Result sheet contains information on the clusters, along with the cluster centers. 


# Performance

When the "Distance" value is minimized, it indicates the resulting accuracy is higher. 

Execute the algorithm several times to find the best results.

The script will stop execution when the clusters are normalized or when the maximum iterations are reached (whichever comes first).  You can increase the number of iterations for better results.


# More info

The example dataset is IRIS from UC Irvine Machine Learning Repository: https://archive.ics.uci.edu/ml/datasets.html

The original VBA script is the work of bquanttrading, https://quantmacro.wordpress.com/


