import pandas as pd
import networkx as nx

# Specify Excel file path
excel_file_path = r'C:\Users\qingyichen\Desktop\stockrate_data.xlsx'

# Use pandas read_excel function to read the Excel file
data = pd.read_excel(excel_file_path, index_col=0)

# Calculate the correlation matrix
correlation_matrix = data.corr()

# Build the graph
G = nx.Graph()

# Add nodes
G.add_nodes_from(data.columns)

# Add edges with weights
for i in range(len(data.columns)):
    for j in range(i + 1, len(data.columns)):
        stock1 = data.columns[i]
        stock2 = data.columns[j]
        weight = correlation_matrix.loc[stock1, stock2]
        G.add_edge(stock1, stock2, weight=weight)

# Use Prim's algorithm to find the minimum spanning tree
mst = nx.minimum_spanning_tree(G)

# Calculate closeness centrality
closeness_centrality = nx.closeness_centrality(mst)

# Get the top 15 nodes
top_nodes_closeness = sorted(closeness_centrality, key=closeness_centrality.get, reverse=True)[:15]

# Save data for only the top 15 nodes to a file
result_df = pd.DataFrame([(node, closeness_centrality[node]) for node in top_nodes_closeness],
                         columns=['Node', 'Closeness Centrality'])
result_df.to_excel(r'D:\python\closeness_centrality_top15.xlsx', index=False)


