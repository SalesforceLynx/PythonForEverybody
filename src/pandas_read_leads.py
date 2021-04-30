import pandas as pd
import numpy as np

pd.options.display.float_format = '{:.2f}'.format

data_df = pd.read_excel("../reports/2021 Q1 Leads Snap Report-2021-04-11-13-22-33.xlsx")
data_df['CreatedDate'] = pd.to_datetime(data_df['CreatedDate'])

data_df['MaxDate'] = data_df.groupby(['LeadId', 'LeadStatus'])['CreatedDate'].transform('max')
data_df['MinDate'] = data_df.groupby(['LeadId', 'LeadStatus'])['CreatedDate'].transform('min')
data_df['Difference'] = pd.to_datetime(data_df['MaxDate']) - pd.to_datetime(data_df['MinDate'])

agg_df = data_df.groupby(['LeadId', 'LeadStatus', 'Email']).agg(MaxDate=('CreatedDate', 'max'),
                                                                MinDate=('CreatedDate', 'min')).reset_index()
agg_df['Difference'] = pd.to_datetime(agg_df['MaxDate']) - pd.to_datetime(agg_df['MinDate'])
agg_df['Quarter'] = agg_df['MaxDate'].dt.quarter
agg_df_c = agg_df.groupby(['LeadStatus', 'Quarter', 'Difference'])["Quarter"].count().reset_index(name="Count")
agg_df_c['NumericDifference'] = agg_df_c['Difference'].values.astype(np.int64)
agg_df_s = agg_df_c.groupby(['LeadStatus'])["NumericDifference"].sum().reset_index(name="Sum")
agg_df_m = agg_df_c.groupby([agg_df_c.LeadStatus == 'Contacted'])["NumericDifference"].sum()
agg_df_s.replace(0, np.NAN)
agg_df_s = agg_df_c.groupby('LeadStatus')["NumericDifference"].mean().reset_index()


data_df.to_json(orient='records')

with pd.ExcelWriter('../out/Q1Leads-SnapComputedReport-17042021-0909PM.xlsx', engine='xlsxwriter') as writer:
    data_df.to_excel(writer, sheet_name='New Computed Data', index=False)
    agg_df.to_excel(writer, sheet_name='Computed Agg Data', index=False)
    agg_df_c.to_excel(writer, sheet_name='Computed Statistics', index=False)
    agg_df_s.to_excel(writer, sheet_name='Computed Mean', index=False)


print(agg_df_m)
