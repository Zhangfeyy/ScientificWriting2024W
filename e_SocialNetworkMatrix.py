import pandas as pd
import numpy as np
from Country import network_country_list

def create_network_matrix(network_country_list):
    # Get unique countries
    unique_countries = sorted(set(country for group in network_country_list for country in group))

    # Create a matrix of zeros
    matrix = np.zeros((len(unique_countries), len(unique_countries)), dtype=int)

    # Create a mapping of countries to indices
    country_to_index = {country: index for index, country in enumerate(unique_countries)}

    # Fill the matrix
    for group in network_country_list:
        # Create edges between all pairs of countries in the group
        for i in range(len(group)):
            for j in range(i + 1, len(group)):
                country1, country2 = group[i], group[j]
                idx1 = country_to_index[country1]
                idx2 = country_to_index[country2]
                # Increment both symmetric positions
                matrix[idx1, idx2] += 1
                matrix[idx2, idx1] += 1

    # Create DataFrame
    df = pd.DataFrame(matrix, index=unique_countries, columns=unique_countries)

    return df

# Create and save the adjacency matrix
matrix_df = create_network_matrix(network_country_list)

# Save to Excel
matrix_df.to_excel('SocialNetworkWeighted.xlsx')
