# k-means-excel

Clustering is the process of partitioning a set of data into related groups.
K-means clustering is usefull for Data Mining and Business Intelligence.

Here is k-means in plain English:
https://rayli.net/blog/data/top-10-data-mining-algorithms-in-plain-english/#2_k-means

This is a k-means clustering implementation in VBA. The script is based on David Arthur and Sergei Vassilvitski k-means++, which chooses the initial centroids.

Please enter your data in Excel worksheet "Data". Follow the instructions and click the button to start.

The Result sheet contains information on the clusters, along with the centroids. When the "Distance" is minimised, the result accuracy is higher. Execute it several times for the best result.

The script  will stop execution when the clusters are normalized or when the maximum iterations are reached (whichever comes first).

The example dataset is IRIS from UC Irvine Machine Learning Repository: https://archive.ics.uci.edu/ml/datasets.html

The original VBA script is presented in the blog: 
https://quantmacro.wordpress.com/


