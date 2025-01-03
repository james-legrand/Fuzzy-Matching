import pandas as pd
from rapidfuzz import fuzz as rf
from tqdm import tqdm

vivid_data_path = r"Z:\z.training\Python\Fuzzy-Matching\Raw Data\vividseatsPricesTest.csv"
stubhub_data_path = r"Z:\z.training\Python\Fuzzy-Matching\Raw Data\stubhubPricesTest.csv"

output_path = r"Z:\z.training\Python\Fuzzy-Matching\Output\output_test.xlsx"

vivid_df = pd.read_csv(vivid_data_path)
stubhub_df = pd.read_csv(stubhub_data_path)

vivid_id = 'vividseatsID' 
stub_id = 'stubhubID' 
vivid_match1 = 'vividseatsTitle'  
stub_match1 = 'stubhubTitle'
vivid_match2 = 'vividseatsPrice'  
stub_match2 = 'stubhubPrice' 

vivid_df = vivid_df.astype(str)
stubhub_df = stubhub_df.astype(str)

data = []
for i in tqdm(range(len(vivid_df)), desc="Processing vivid_df"):
    for j in range(len(stubhub_df)):
            set_ratio1 = rf.token_set_ratio(vivid_df[vivid_match1].iloc[i], stubhub_df[stub_match1].iloc[j])
            set_ratio2 = rf.token_set_ratio(vivid_df[vivid_match2].iloc[i], stubhub_df[stub_match2].iloc[j])
            
            ratio = min(set_ratio1, set_ratio2)
            
            data.append([vivid_df[vivid_id].iloc[i], stubhub_df[stub_id].iloc[j], vivid_df[vivid_match1].iloc[i], stubhub_df[stub_match1].iloc[j], vivid_df[vivid_match2].iloc[i],stubhub_df[stub_match2].iloc[j], set_ratio1, set_ratio2, ratio])

column_list = [vivid_id, stub_id, vivid_match1, stub_match1,vivid_match2, stub_match2, 'Match Score 1', 'Match Score 2', 'Match Score']
result_df = pd.DataFrame(data, columns=column_list)

result_df.to_excel(output_path, index=False)