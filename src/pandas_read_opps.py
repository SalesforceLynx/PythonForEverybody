import pandas as pd
import numpy as np
#import xlsxwriter

data_df = pd.read_excel("../reports/Opportunity Average Stage Duration-2021-04-19-13-11-46.xlsx")
data_df['CloseDate'] = pd.to_datetime(data_df['CloseDate'])

data_df['MaxDate'] = data_df.groupby(['OpportunityID', 'OpportunityName', 'ToStage', 'FromStage'])['CloseDate'].transform('max')
data_df['MinDate'] = data_df.groupby(['OpportunityID', 'OpportunityName', 'ToStage', 'FromStage'])['CloseDate'].transform('min')
data_df['Difference'] = pd.to_datetime(data_df['MaxDate']) - pd.to_datetime(data_df['MinDate'])

agg_df = data_df.groupby(['OpportunityID', 'OpportunityName', 'ToStage', 'FromStage']).agg(MaxDate=('CloseDate', 'max'),
                                                                              MinDate=('CloseDate', 'min')).reset_index()
agg_df['Difference'] = pd.to_datetime(agg_df['MaxDate']) - pd.to_datetime(agg_df['MinDate'])
agg_df['NumericDifference'] = agg_df['Difference'].values.astype(np.int64)
agg_df['Quarter'] = agg_df['MaxDate'].dt.quarter
agg_df_by_ToStage = agg_df.groupby(['ToStage', 'Quarter', 'Difference'])["ToStage"].count().reset_index(name="Count")
agg_df_by_ToStage = agg_df.groupby(['ToStage', 'Quarter'])["NumericDifference"].mean().reset_index(name="Mean")
#print(agg_df.NumericDifference.dtype)

#data_df.to_json(orient='records')

with pd.ExcelWriter('../out/Q1Opps-AvgStageDuration-ComputedReport-19042021-0400PM.xlsx', engine='xlsxwriter') as writer:
    data_df.to_excel(writer, sheet_name='New Computed Data', index=False)
    agg_df.to_excel(writer, sheet_name='Computed Agg Data', index=False)
    agg_df_by_ToStage.to_excel(writer, sheet_name='Computed Statistics', index=False)

print(agg_df_by_ToStage)

#print(data_df)
