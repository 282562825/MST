import pandas as pd
import numpy as np
import networkx as nx
import matplotlib.pyplot as plt

# Specify Excel file path
excel_file_path = r'C:\Users\qingyichen\Desktop\stockrate_data - 23.xlsx'

# Use pandas read_excel function to read the Excel file
data = pd.read_excel(excel_file_path, index_col=0)

# Ensure data is loaded correctly
print(data.head())

# Calculate the correlation matrix
correlation_matrix = data.corr()

# Build the graph
G = nx.Graph()

# Add nodes
G.add_nodes_from(data.columns)

# Add edges with weights
for i in range(len(data.columns)):
    for j in range(i+1, len(data.columns)):
        stock1 = data.columns[i]
        stock2 = data.columns[j]
        weight = correlation_matrix.loc[stock1, stock2]
        G.add_edge(stock1, stock2, weight=weight)

# Use Prim's algorithm to find the minimum spanning tree
mst = nx.minimum_spanning_tree(G)

# Calculate node degrees
node_degrees = dict(mst.degree())

# Get the top 15 nodes based on degree
top_15_degrees = dict(sorted(node_degrees.items(), key=lambda x: x[1], reverse=True)[:15])

# Convert to DataFrame
degrees_df = pd.DataFrame(list(top_15_degrees.items()), columns=['Node', 'Degree'])

# Print and save the DataFrame
print(degrees_df)
degrees_df.to_excel(r'D:\python\top_15_degrees.xlsx', index=False)

# Draw the degree distribution plot
plt.figure(figsize=(12, 8))
degree_values = list(node_degrees.values())
plt.hist(degree_values, bins=20, alpha=0.75, color='skyblue', edgecolor='black')
plt.title('Degree Distribution of Minimum Spanning Tree')
plt.xlabel('Degree')
plt.ylabel('Frequency')
plt.grid(True)
plt.savefig(r'D:\python\degree_distribution.png')
plt.show()
