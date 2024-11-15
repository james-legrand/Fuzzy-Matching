import pandas as pd
import numpy as np
from fuzzywuzzy import fuzz as fw_fuzz
from rapidfuzz import fuzz as rf_fuzz
from tqdm import tqdm

vivid_data_path = r"Z:\y.utilities\fuzzy_matching\data\vividseatsPricesTest.csv"
stubhub_data_path = r"Z:\y.utilities\fuzzy_matching\data\stubhubPricesTest.csv"

output_path = r"Z:\y.utilities\fuzzy_matching\Experiments\Replication\replicationOutput.xlsx"

matching_library = 'rapidfuzz'  # 'rapidfuzz' or 'fuzzywuzzy'

vivid_df = pd.read_csv(vivid_data_path)
stubhub_df = pd.read_csv(stubhub_data_path)

column1_id = 'vividseatsTitle'  # set this to your unique identifier column for vivid_df
column2_id = 'stubhubTitle'  # set this to your unique identifier column for stubhub_df
column1_match = 'vividseatsTitle'  # set the column to match on vivid_df
column2_match = 'stubhubTitle'  # set the column to match on stubhub_df
output_type = 2  # 1 for all possible combinations, 2 for highest matches only
matching_type = 1  # 1 for set ratio, 2 for sort ratio, 3 for max of both

# Select the correct fuzzing functions based on the library
if matching_library == 'fuzzywuzzy':
    token_set_ratio = fw_fuzz.token_set_ratio
    token_sort_ratio = fw_fuzz.token_sort_ratio
elif matching_library == 'rapidfuzz':
    token_set_ratio = rf_fuzz.token_set_ratio
    token_sort_ratio = rf_fuzz.token_sort_ratio

if output_type == 1:
    data = []
    for i in tqdm(range(len(vivid_df)), desc="Processing vivid_df"):
        for j in range(len(stubhub_df)):
            if pd.isna(vivid_df[column1_match].iloc[i]) or pd.isna(stubhub_df[column2_match].iloc[j]):
                set_ratio, sort_ratio, ratio = 'N/A', 'N/A', 'N/A'
            else:
                set_ratio = token_set_ratio(vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[j]) / 100.0
                sort_ratio = token_sort_ratio(vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[j]) / 100.0
                ratio = max(set_ratio, sort_ratio)

            if matching_type == 1:
                score = set_ratio
            elif matching_type == 2:
                score = sort_ratio
            else:
                score = ratio

            data.append([vivid_df[column1_id].iloc[i], stubhub_df[column2_id].iloc[j], vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[j], score])

    column_list = [column1_id, column2_id, column1_match, column2_match, 'Match Score']
    result_df = pd.DataFrame(data, columns=column_list)

elif output_type == 2:
    data = []
    for i in tqdm(range(len(vivid_df)), desc="Processing vivid_df"):
        max_score, best_match = 0, None
        for j in range(len(stubhub_df)):
            if pd.isna(vivid_df[column1_match].iloc[i]) or pd.isna(stubhub_df[column2_match].iloc[j]):
                score = 0
            else:
                set_ratio = token_set_ratio(vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[j]) / 100.0
                sort_ratio = token_sort_ratio(vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[j]) / 100.0
                ratio = max(set_ratio, sort_ratio)

                if matching_type == 1:
                    score = set_ratio
                elif matching_type == 2:
                    score = sort_ratio
                else:
                    score = ratio

            if score > max_score:
                max_score = score
                best_match = j

        if best_match is not None:
            data.append([vivid_df[column1_id].iloc[i], stubhub_df[column2_id].iloc[best_match], vivid_df[column1_match].iloc[i], stubhub_df[column2_match].iloc[best_match], max_score])

    column_list = [column1_id, column2_id, column1_match, column2_match, 'Best Match Score']
    result_df = pd.DataFrame(data, columns=column_list)

result_df.to_excel(output_path, index=False)
