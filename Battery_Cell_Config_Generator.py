import pandas as pd

def group_batteries(filename, num_cells_per_pack, num_packs):
    # Read in the data from Excel
    df = pd.read_excel(filename, engine='openpyxl')

    # Convert the 'Capacity (mAh)' column to float
    df['Capacity (mAh)'] = df['Capacity (mAh)'].astype(float)

    # Sort the battery capacities
    sorted_batteries = df.sort_values(by=['Capacity (mAh)'], ascending=False).reset_index()

    # Initialize empty packs
    packs = [[] for _ in range(num_packs)]
    pack_ocvs = [[] for _ in range(num_packs)]
    pack_capacity = [[] for _ in range(num_packs)]

    pack_sums = [0] * num_packs

    # Distribute batteries among packs to equalize capacities
    for i, row in sorted_batteries.iterrows():
        # Find the pack with the lowest current Capacity (mAh) sum
        min_pack_index = pack_sums.index(min(pack_sums))
        
        # If the pack isn't full, add the battery
        if len(packs[min_pack_index]) < num_cells_per_pack:
            packs[min_pack_index].append(row['Location'])
            pack_ocvs[min_pack_index].append(row['Final OCV'])
            pack_capacity[min_pack_index].append(row['Capacity (mAh)'])
            pack_sums[min_pack_index] += row['Capacity (mAh)']

    return packs, pack_sums, pack_ocvs, pack_capacity

def save_to_excel(packs, pack_sums, pack_ocvs, pack_capacity, output_filename):
    # Create a DataFrame from the packs
    df_output = pd.DataFrame()
    for i, pack in enumerate(packs):
        df_output[f'Pack {i+1} Location'] = pd.Series(pack)
        df_output[f'Pack {i+1} OCV (V)'] = pd.Series(pack_ocvs[i])
        df_output[f'Pack {i+1} Capacity (mAh)'] = pd.Series(pack_capacity[i])
        
        #df_output.loc['Total Capacity'] = pack_sums
        # Save the DataFrame to an Excel file
    df_output.to_excel(output_filename, index=False)

# Example
filename = 'Battery Test Logger.xlsx'
num_cells_per_pack = 14
num_packs = 29

packs, pack_sums, pack_ocvs, pack_capacity = group_batteries(filename, num_cells_per_pack, num_packs)
save_to_excel(packs, pack_sums, pack_ocvs, pack_capacity,  'battery_packs_output.xlsx')


