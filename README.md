# k-means-excel

Clustering is the process of partitioning a set of data into related groups.
K-means clustering is usefull for Data Mining and Business Intelligence.

Here is k-means in plain English:
https://rayli.net/blog/data/top-10-data-mining-algorithms-in-plain-english/#2_k-means

K-means does not need to know the qualities of the data, just the number of clusters (groups) that need to be created.
The output is a number indicating which group (cluster) each record belongs to.

This example is k-means clustering implemented in EXCEL VBA.

Please enter your data in worksheet "Data" starting at cell B2
Provide the initial cluster centroids, for as many clusters as you want, in worksheet "Start", and click the button to go.

The script  will stop execution when the clusters are normalized or when the maximum iterations are reached (whichever comes first).

The original VBA script is presented in the blog: 
https://quantmacro.wordpress.com/


