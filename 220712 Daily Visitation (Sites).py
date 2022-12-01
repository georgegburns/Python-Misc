import pandas as pd
import numpy as np
visitation_df = pd.read_excel("Visitation.xlsx")
#print(visitation_df.head(5))
#print(visitation_df.dtypes)
visitation_df["Year"] = visitation_df["DATE"].dt.year.astype('Int64')
visitation_daily = visitation_df.groupby(["SITE","Year","DATE"])["VISITATION"].sum()
#print(visitation_daily)
visitation_daily.to_excel("Daily Visitation (Site).xlsx",merge_cells=False,sheet_name="Daily Visitation",index=True)