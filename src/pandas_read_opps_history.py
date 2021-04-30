import pandas as pd
import numpy as np


data_df = pd.read_excel("../reports/Opportunity Field History Report-2021-04-20-14-11-04.xlsx")
data_df['EditDate'] = pd.to_datetime(data_df['EditDate'])

data_df['MaxDate'] = data_df.groupby(['OpportunityID', 'OpportunityName'])['EditDate'].transform('max')
data_df['MinDate'] = data_df.groupby(['OpportunityID', 'OpportunityName'])['EditDate'].transform('min')
data_df['Difference'] = pd.to_datetime(data_df['MaxDate']) - pd.to_datetime(data_df['MinDate'])

agg_df = data_df.groupby(['OpportunityID', 'OpportunityName', 'NewValue', 'StageDuration'])\
    .agg(MaxDate=('EditDate', 'max'),
         MinDate=('EditDate', 'min')).reset_index()
agg_df['Difference'] = pd.to_datetime(agg_df['MaxDate']) - pd.to_datetime(agg_df['MinDate'])
agg_df['NumericDifference'] = agg_df['Difference'].values.astype(np.int64)
agg_df['Quarter'] = agg_df['MaxDate'].dt.quarter
data_df['Quarter'] = agg_df['MaxDate'].dt.quarter
agg_df_by_NewValue = agg_df.groupby(['OpportunityID', 'OpportunityName', 'NewValue', 'Quarter'])["OpportunityID"].count().reset_index(name="Count")
#agg_df_OppId = data_df.groupby(["OpportunityName", "MaxDate", "StageDuration"])["MaxDate"].count().reset_index(name="NewValue Count")
df_unique_values = data_df.groupby(['OpportunityName', data_df.NewValue == 'Qualified Out', 'MaxDate', 'Quarter'])["MaxDate"].max().reset_index(name="Max of Date")
df_filtered = data_df.groupby(['OpportunityID', 'OpportunityName', 'StageDuration', 'MaxDate'])['NewValue'].apply(', '.join).reset_index()
df_filtered['Quarter'] = df_filtered['MaxDate'].dt.quarter
s1 = pd.Series(df_filtered['NewValue'])
df_filtered['Qualified Out'] = s1.str.contains('Qualified Out', regex=True)
df_filtered['CurrentStage'] = df_filtered["NewValue"].str.split(',').str[-1]
df_filtered['CurrentStage'] = df_filtered['CurrentStage'].str.strip()
#print(df_filtered['CurrentStage'])


with pd.ExcelWriter('../out/Q1Opps-FieldHistory-ComputedReport-20042021-0515PM.xlsx', engine='xlsxwriter') as writer:
    data_df.to_excel(writer, sheet_name='New Computed Data', index=False)
    agg_df.to_excel(writer, sheet_name='Computed Agg Data', index=False)
    agg_df_by_NewValue.to_excel(writer, sheet_name='ComputedStats1', index=False)
    #agg_df_OppId.to_excel(writer, sheet_name='ComputedStats2', index=False)
    df_unique_values.to_excel(writer, sheet_name='ComputedStats3', index=False)
    df_filtered.to_excel(writer, sheet_name='ComputedStats4', index=False)


print(df_filtered)


