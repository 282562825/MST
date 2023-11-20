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

# Add edges with weights, filtering out edges with weights between -0.05 and 0.05
for i in range(len(data.columns)):
    for j in range(i + 1, len(data.columns)):
        stock1 = data.columns[i]
        stock2 = data.columns[j]
        weight = correlation_matrix.loc[stock1, stock2]
        if -0.5 <= weight <= 0.5:
            continue  # Skip edges with weights between -0.05 and 0.05
        G.add_edge(stock1, stock2, weight=weight)

# Use Kruskal's algorithm to generate the minimum spanning tree
mst = nx.minimum_spanning_tree(G)

# Calculate node degrees
node_degrees = dict(G.degree())

# Select nodes to keep (nodes with degree greater than or equal to 50)
nodes_to_keep = [node for node, degree in node_degrees.items() if degree >= 50]
# Build a subgraph containing the selected nodes
G_subgraph = G.subgraph(nodes_to_keep)

# Set node sizes based on node degrees
node_sizes = [node_degrees[node] * 50000 for node in G.nodes]

# Draw the minimum spanning tree network
pos = nx.spring_layout(mst)
labels = {stock: stock for stock in data.columns}
edge_labels = {(stock1, stock2): f'{weight:.2f}' for (stock1, stock2, weight) in mst.edges(data='weight')}

# Visualize the minimum spanning tree
plt.figure(figsize=(50, 30))

nx.draw_networkx_nodes(mst, pos, node_color='skyblue', alpha=0.8)
nx.draw_networkx_labels(mst, pos, labels=labels)
nx.draw_networkx_edges(mst, pos, width=1, edge_color='gray', alpha=0.5)
nx.draw_networkx_edge_labels(mst, pos, edge_labels=edge_labels, font_color='red')

plt.title('Minimum Spanning Tree of Stock Correlations')

plt.savefig(r'D:\python\minimum_spanning_tree.png')
