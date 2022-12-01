# importing libraries
import pandas as pd
import numpy as np
from matplotlib import pyplot as plt 
from datetime import date

# reading excel files
Full = pd.read_excel("MidValueDonor.xlsx", sheet_name="Full")
FullSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="FullSource")
Other = pd.read_excel("MidValueDonor.xlsx", sheet_name="Other")
OtherSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="OtherSource")
Legacy = pd.read_excel("MidValueDonor.xlsx", sheet_name="Legacy")
LegacySource = pd.read_excel("MidValueDonor.xlsx", sheet_name="LegacySource")
AFS = pd.read_excel("MidValueDonor.xlsx", sheet_name="AFS")
AFSSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="AFSSource")
Patron = pd.read_excel("MidValueDonor.xlsx", sheet_name="Patron")
PatronSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="PatronSource")
Complimentary = pd.read_excel("MidValueDonor.xlsx", sheet_name="Complimentary")
ComplimentarySource = pd.read_excel("MidValueDonor.xlsx", sheet_name="ComplimentarySource")
Life = pd.read_excel("MidValueDonor.xlsx", sheet_name="Life")
LifeSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="LifeSource")
AW = pd.read_excel("MidValueDonor.xlsx", sheet_name="Adopt A Wetland")
AWSource = pd.read_excel("MidValueDonor.xlsx", sheet_name="Adopt A Wetland Source")
FullDemo = pd.read_excel("MidValueDonor.xlsx", sheet_name="Full Demo")
print(FullDemo)

# head and dtypes
print("\n")
print(Full.head(10))
print(Full.dtypes)
print("\n")
print(FullSource.head(10))
print(FullSource.dtypes)

# defining totals

PatronTotal = 39
LegacyTotal = 979
HonoraryTotal = 124
LifeTotal = 4339

# cleaning the files
Full = Full.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
Other = Other.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
Other = Other[~Other['Serial Number'].isin(Legacy['Serial Number'])]
Other = Other[~Other['Serial Number'].isin(AFS['Serial Number'])]
Other = Other[~Other['Serial Number'].isin(Patron['Serial Number'])]
Other = Other[~Other['Serial Number'].isin(Complimentary['Serial Number'])]
Other = Other[~Other['Serial Number'].isin(Life['Serial Number'])]
Other = Other[~Other['Serial Number'].isin(AW['Serial Number'])]
Legacy = Legacy.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
AFS = AFS.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
Patron = Patron.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
Complimentary = Complimentary.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
Life = Life.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()
AW = AW.groupby(['Serial Number', 'Year']).agg({'Sum of Donations':'sum'}).reset_index()

# full data 
quartile_x = np.percentile(Full["Sum of Donations"],99)
quartile_y = np.percentile(Full["Sum of Donations"],97)
#quartile_x = 2000
upper_quartile = np.percentile(Full["Sum of Donations"], 75)
lower_quartile = np.percentile(Full["Sum of Donations"], 25)
stdev = np.std(Full["Sum of Donations"])
mean = np.mean(Full["Sum of Donations"])
max = np.amax(Full["Sum of Donations"])
median = np.median(Full["Sum of Donations"])
min = np.amin(Full["Sum of Donations"])
count = len(np.unique(list(Full["Serial Number"])))
FullxSN = Full[Full["Sum of Donations"]> quartile_x]
FullySN = Full[(Full["Sum of Donations"]>=quartile_y) & (Full["Sum of Donations"]<=quartile_x)]
FullySN = FullySN[~FullySN['Serial Number'].isin(FullxSN['Serial Number'])]
FullzSN = Full[Full["Sum of Donations"]< quartile_y]
FullzSN = FullzSN[~FullzSN['Serial Number'].isin(FullxSN['Serial Number'])]
FullzSN = FullzSN[~FullzSN['Serial Number'].isin(FullySN['Serial Number'])]
countx = len(np.unique(list(FullxSN["Serial Number"])))
county = len(np.unique(list(FullySN["Serial Number"])))
countz = len(np.unique(list(FullzSN["Serial Number"])))
print("\n")
print("All Donors (New High Value):")
print("Standard Deviation: " + str(stdev))
print("Mean: " + str(mean))
print("Max: " + str(max))
print("Median: " + str(median))
print("Min: " + str(min))
print("Upper Quartile: " + str(upper_quartile))
print("Lower Quartile: " + str(lower_quartile))
print("X Quartile: " + str(quartile_x))
print("Y Quartile: " + str(quartile_y))
print("Count of donors: " + str(count))
print("High Value Donors: " + str(countx))
print("Mid Value Donors: " + str(county))
print("Low Value Donors: " + str(countz))

#partitioning the df by < 2000 and >= QY
FullHV = Full[Full["Sum of Donations"] >= 2000]
FullMV = Full[(Full["Sum of Donations"] < 2000) & (Full["Sum of Donations"] >= quartile_y)]
FullMV = FullMV[~FullMV['Serial Number'].isin(FullHV['Serial Number'])]
FullLV = Full[Full['Sum of Donations'] < quartile_y]
FullLV = FullLV[~FullLV['Serial Number'].isin(FullHV['Serial Number'])]
FullLV = FullLV[~FullLV['Serial Number'].isin(FullMV['Serial Number'])]

max3 = np.amax(FullMV["Sum of Donations"])
min3 = np.amin(FullMV["Sum of Donations"])
stdev3 = np.std(FullMV["Sum of Donations"])
mean3 = np.mean(FullMV["Sum of Donations"])
median3 = np.median(FullMV["Sum of Donations"])
count3 = len(np.unique(list(FullMV["Serial Number"])))
count3HV = len(np.unique(list(FullHV["Serial Number"])))
count3LV = len(np.unique(list(FullLV["Serial Number"])))
print("\n")
print("All Donors (Current High Value): < 2000, >QX")
print("Standard Deviation: " + str(stdev3))
print("Mean: " + str(mean3))
print("Max: " + str(max3))
print("Median: " + str(median3))
print("Min: " + str(min3))
print("High value Donors: " + str(count3HV))
print("Mid value Donors: " + str(count3))
print("Low value Donors: " + str(count3LV))

# Other analysis
OtherCount = len(np.unique(list(Other["Serial Number"])))
OtherMean = np.mean(Other["Sum of Donations"])
OtherMax = np.amax(Other["Sum of Donations"])
OtherPercentage = np.round((OtherCount/count) * 100)
OtherHVOld = Other[(Other["Sum of Donations"] >= 2000)]
OtherHVx = Other[(Other["Sum of Donations"] > quartile_x)]
OtherOld = Other[(Other["Sum of Donations"] < 2000) & (Other["Sum of Donations"] >= quartile_y)]
OtherOld = OtherOld[~OtherOld['Serial Number'].isin(OtherHVOld['Serial Number'])]
OtherNew = Other[(Other["Sum of Donations"] <= quartile_x) & (Other["Sum of Donations"] >= quartile_y)]
OtherNew = OtherNew[~OtherNew['Serial Number'].isin(OtherHVx['Serial Number'])]
OtherOldMean = np.mean(OtherOld["Sum of Donations"])
OtherNewMean = np.mean(OtherNew["Sum of Donations"])
OtherOldCount = len(np.unique(list(OtherOld["Serial Number"])))
OtherNewCount = len(np.unique(list(OtherNew["Serial Number"])))
OtherOldPercentage = np.round((OtherOldCount/OtherCount) * 100, 1)
OtherMVOldPercentage = np.round((OtherOldCount/count3) * 100, 1)
OtherNewPercentage = np.round((OtherNewCount/OtherCount) * 100, 1)
OtherMVNewPercentage = np.round((OtherNewCount/county) * 100, 1)
print("\n")
print("Other Donors")
print("Total Other Count: " + str(OtherCount))
print("Total Other Mean Donation Value: " + str(OtherMean))
print("Percentage of Total Donors: " + str(OtherPercentage) + "%")
print("\n")
print("Other Mean (Using Current HV): " + str(OtherOldMean))
print("Other Count (Using Current HV): " + str(OtherOldCount))
print("Other Percentage Of Total Other (Using Current HV): " + str(OtherOldPercentage) + "%")
print("Other Percentage Of Mid-Value (Using Current HV): " + str(OtherMVOldPercentage) + "%")
print("\n")
print("Other Mean (Using New): " + str(OtherNewMean))
print("Other Count (Using New): " + str(OtherNewCount))
print("Other Percentage Of Total Other (Using New): " + str(OtherNewPercentage) + "%")
print("Other Percentage Of Mid-Value (Using New): " + str(OtherMVNewPercentage) + "%")

#Legacy analysis
LegacyCount = len(np.unique(list(Legacy["Serial Number"])))
LegacyMean = np.mean(Legacy["Sum of Donations"])
LegacyMax = np.amax(Legacy["Sum of Donations"])
LegacyDonorPercentage = np.round((LegacyCount/LegacyTotal) * 100)
LegacyPercentage = np.round((LegacyCount/count) * 100)
LegacyHVOld = Legacy[(Legacy["Sum of Donations"] >= 2000)]
LegacyHVx = Legacy[(Legacy["Sum of Donations"] > quartile_x)]
LegacyOld = Legacy[(Legacy["Sum of Donations"] < 2000) & (Legacy["Sum of Donations"] >= quartile_y)]
LegacyOld = LegacyOld[~LegacyOld['Serial Number'].isin(LegacyHVOld['Serial Number'])]
LegacyNew = Legacy[(Legacy["Sum of Donations"] <= quartile_x) & (Legacy["Sum of Donations"] >= quartile_y)]
LegacyNew = LegacyNew[~LegacyNew['Serial Number'].isin(LegacyHVx['Serial Number'])]
LegacyOldMean = np.mean(LegacyOld["Sum of Donations"])
LegacyNewMean = np.mean(LegacyNew["Sum of Donations"])
LegacyOldCount = len(np.unique(list(LegacyOld["Serial Number"])))
LegacyNewCount = len(np.unique(list(LegacyNew["Serial Number"])))
LegacyOldPercentage = np.round((LegacyOldCount/LegacyCount) * 100, 1)
LegacyMVOldPercentage = np.round((LegacyOldCount/count3) * 100, 1)
LegacyNewPercentage = np.round((LegacyNewCount/LegacyCount) * 100, 1)
LegacyMVNewPercentage = np.round((LegacyNewCount/county) * 100, 1)
print("\n")
print("Legacy Donors")
print("Total Legacy Count: " + str(LegacyCount))
print("Total Legacy Mean Donation Value: " + str(LegacyMean))
print("Percentage of Total Donors: " + str(LegacyPercentage) + "%")
print("\n")
print("Legacy Mean (Using Current HV): " + str(LegacyOldMean))
print("Legacy Count (Using Current HV): " + str(LegacyOldCount))
print("Legacy Percentage Of Total Legacy (Using Current HV): " + str(LegacyOldPercentage) + "%")
print("Legacy Percentage Of Mid-Value (Using Current HV): " + str(LegacyMVOldPercentage) + "%")
print("\n")
print("Legacy Mean (Using New): " + str(LegacyNewMean))
print("Legacy Count (Using New): " + str(LegacyNewCount))
print("Legacy Percentage Of Total Legacy (Using New): " + str(LegacyNewPercentage) + "%")
print("Legacy Percentage Of Mid-Value (Using New): " + str(LegacyMVNewPercentage) + "%")

# AFS analysis
AFSCount = len(np.unique(list(AFS["Serial Number"])))
AFSMean = np.mean(AFS["Sum of Donations"])
AFSMax = np.amax(AFS["Sum of Donations"])
AFSPercentage = np.round((AFSCount/count) * 100)
AFSHVOld = AFS[(AFS["Sum of Donations"] >= 2000)]
AFSHVx = AFS[(AFS["Sum of Donations"] > quartile_x)]
AFSOld = AFS[(AFS["Sum of Donations"] < 2000) & (AFS["Sum of Donations"] >= quartile_y)]
AFSOld = AFSOld[~AFSOld['Serial Number'].isin(AFSHVOld['Serial Number'])]
AFSNew = AFS[(AFS["Sum of Donations"] <= quartile_x) & (AFS["Sum of Donations"] >= quartile_y)]
AFSNew = AFSNew[~AFSNew['Serial Number'].isin(AFSHVx['Serial Number'])]
AFSOldMean = np.mean(AFSOld["Sum of Donations"])
AFSNewMean = np.mean(AFSNew["Sum of Donations"])
AFSOldCount = len(np.unique(list(AFSOld["Serial Number"])))
AFSNewCount = len(np.unique(list(AFSNew["Serial Number"])))
AFSOldPercentage = np.round((AFSOldCount/AFSCount) * 100, 1)
AFSMVOldPercentage = np.round((AFSOldCount/count3) * 100, 1)
AFSNewPercentage = np.round((AFSNewCount/AFSCount) * 100, 1)
AFSMVNewPercentage = np.round((AFSNewCount/county) * 100, 1)
print("\n")
print("AFS Donors")
print("Total AFS Count: " + str(AFSCount))
print("Total AFS Mean Donation Value: " + str(AFSMean))
print("Percentage of Total Donors: " + str(AFSPercentage) + "%")
print("\n")
print("AFS Mean (Using Current HV): " + str(AFSOldMean))
print("AFS Count (Using Current HV): " + str(AFSOldCount))
print("AFS Percentage Of Total AFS (Using Current HV): " + str(AFSOldPercentage) + "%")
print("AFS Percentage Of Mid-Value (Using Current HV): " + str(AFSMVOldPercentage) + "%")
print("\n")
print("AFS Mean (Using New): " + str(AFSNewMean))
print("AFS Count (Using New): " + str(AFSNewCount))
print("AFS Percentage Of Total AFS (Using New): " + str(AFSNewPercentage) + "%")
print("AFS Percentage Of Mid-Value (Using New): " + str(AFSMVNewPercentage) + "%")

# Patron analysis
PatronCount = len(np.unique(list(Patron["Serial Number"])))
PatronMean = np.mean(Patron["Sum of Donations"])
PatronMax = np.amax(Patron["Sum of Donations"])
PatronDonorPercentage = np.round((PatronCount/PatronTotal) * 100)
PatronPercentage = np.round((PatronCount/count) * 100)
PatronHVOld = Patron[(Patron["Sum of Donations"] >= 2000)]
PatronHVx = Patron[(Patron["Sum of Donations"] > quartile_x)]
PatronOld = Patron[(Patron["Sum of Donations"] < 2000) & (Patron["Sum of Donations"] >= quartile_y)]
PatronOld = PatronOld[~PatronOld['Serial Number'].isin(PatronHVOld['Serial Number'])]
PatronNew = Patron[(Patron["Sum of Donations"] <= quartile_x) & (Patron["Sum of Donations"] >= quartile_y)]
PatronNew = PatronNew[~PatronNew['Serial Number'].isin(PatronHVx['Serial Number'])]
PatronOldMean = np.mean(PatronOld["Sum of Donations"])
PatronNewMean = np.mean(PatronNew["Sum of Donations"])
PatronOldCount = len(np.unique(list(PatronOld["Serial Number"])))
PatronNewCount = len(np.unique(list(PatronNew["Serial Number"])))
PatronOldPercentage = np.round((PatronOldCount/PatronCount) * 100, 1)
PatronMVOldPercentage = np.round((PatronOldCount/count3) * 100, 1)
PatronNewPercentage = np.round((PatronNewCount/PatronCount) * 100, 1)
PatronMVNewPercentage = np.round((PatronNewCount/county) * 100, 1)
print("\n")
print("Patron Donors")
print("Total Patron Count: " + str(PatronCount))
print("Total Patron Mean Donation Value: " + str(PatronMean))
print("Percentage of Total Donors: " + str(PatronPercentage) + "%")
print('Percentage of Total Patrons: ' + str(PatronDonorPercentage) + "%")
print("\n")
print("Patron Mean (Using Current HV): " + str(PatronOldMean))
print("Patron Count (Using Current HV): " + str(PatronOldCount))
print("Patron Percentage Of Donating Patron (Using Current HV): " + str(PatronOldPercentage) + "%")
print("Patron Percentage Of Mid-Value (Using Current HV): " + str(PatronMVOldPercentage) + "%")
print("\n")
print("Patron Mean (Using New): " + str(PatronNewMean))
print("Patron Count (Using New): " + str(PatronNewCount))
print("Patron Percentage Of Donating Patron (Using New): " + str(PatronNewPercentage) + "%")
print("Patron Percentage Of Mid-Value (Using New): " + str(PatronMVNewPercentage) + "%")

# Complimentary analysis
ComplimentaryCount = len(np.unique(list(Complimentary["Serial Number"])))
ComplimentaryMean = np.mean(Complimentary["Sum of Donations"])
ComplimentaryMax = np.amax(Complimentary["Sum of Donations"])
ComplimentaryDonorPercentage = np.round((ComplimentaryCount/HonoraryTotal) * 100)
ComplimentaryPercentage = np.round((ComplimentaryCount/count) * 100)
ComplimentaryHVOld = Complimentary[(Complimentary["Sum of Donations"] >= 2000)]
ComplimentaryHVx = Complimentary[(Complimentary["Sum of Donations"] > quartile_x)]
ComplimentaryOld = Complimentary[(Complimentary["Sum of Donations"] < 2000) & (Complimentary["Sum of Donations"] >= quartile_y)]
ComplimentaryOld = ComplimentaryOld[~ComplimentaryOld['Serial Number'].isin(ComplimentaryHVOld['Serial Number'])]
ComplimentaryNew = Complimentary[(Complimentary["Sum of Donations"] <= quartile_x) & (Complimentary["Sum of Donations"] >= quartile_y)]
ComplimentaryNew = ComplimentaryNew[~ComplimentaryNew['Serial Number'].isin(ComplimentaryHVx['Serial Number'])]
ComplimentaryOldMean = np.mean(ComplimentaryOld["Sum of Donations"])
ComplimentaryNewMean = np.mean(ComplimentaryNew["Sum of Donations"])
ComplimentaryOldCount = len(np.unique(list(ComplimentaryOld["Serial Number"])))
ComplimentaryNewCount = len(np.unique(list(ComplimentaryNew["Serial Number"])))
ComplimentaryOldPercentage = np.round((ComplimentaryOldCount/ComplimentaryCount) * 100, 1)
ComplimentaryMVOldPercentage = np.round((ComplimentaryOldCount/count3) * 100, 1)
ComplimentaryNewPercentage = np.round((ComplimentaryNewCount/ComplimentaryCount) * 100, 1)
ComplimentaryMVNewPercentage = np.round((ComplimentaryNewCount/county) * 100, 1)
print("\n")
print("Complimentary Donors")
print("Total Complimentary Count: " + str(ComplimentaryCount))
print("Total Complimentary Mean Donation Value: " + str(ComplimentaryMean))
print("Percentage of Total Donors: " + str(ComplimentaryPercentage) + "%")
print('Percentage of Total Honorary: ' + str(ComplimentaryDonorPercentage) + "%")
print("\n")
print("Complimentary Mean (Using Current HV): " + str(ComplimentaryOldMean))
print("Complimentary Count (Using Current HV): " + str(ComplimentaryOldCount))
print("Complimentary Percentage Of Donating Honorary (Using Current HV): " + str(ComplimentaryOldPercentage) + "%")
print("Complimentary Percentage Of Mid-Value (Using Current HV): " + str(ComplimentaryMVOldPercentage) + "%")
print("\n")
print("Complimentary Mean (Using New): " + str(ComplimentaryNewMean))
print("Complimentary Count (Using New): " + str(ComplimentaryNewCount))
print("Complimentary Percentage Of Donating Honorary (Using New): " + str(ComplimentaryNewPercentage) + "%")
print("Complimentary Percentage Of Mid-Value (Using New): " + str(ComplimentaryMVNewPercentage) + "%")

# Life analysis
LifeCount = len(np.unique(list(Life["Serial Number"])))
LifeMean = np.mean(Life["Sum of Donations"])
LifeMax = np.amax(Life["Sum of Donations"])
LifeDonorPercentage = np.round((LifeCount/LifeTotal) * 100)
LifePercentage = np.round((LifeCount/count) * 100)
LifeHVOld = Life[(Life["Sum of Donations"] >= 2000)]
LifeHVx = Life[(Life["Sum of Donations"] > quartile_x)]
LifeOld = Life[(Life["Sum of Donations"] < 2000) & (Life["Sum of Donations"] >= quartile_y)]
LifeOld = LifeOld[~LifeOld['Serial Number'].isin(LifeHVOld['Serial Number'])]
LifeNew = Life[(Life["Sum of Donations"] <= quartile_x) & (Life["Sum of Donations"] >= quartile_y)]
LifeNew = LifeNew[~LifeNew['Serial Number'].isin(LifeHVx['Serial Number'])]
LifeOldMean = np.mean(LifeOld["Sum of Donations"])
LifeNewMean = np.mean(LifeNew["Sum of Donations"])
LifeOldCount = len(np.unique(list(LifeOld["Serial Number"])))
LifeNewCount = len(np.unique(list(LifeNew["Serial Number"])))
LifeOldPercentage = np.round((LifeOldCount/LifeCount) * 100, 1)
LifeMVOldPercentage = np.round((LifeOldCount/count3) * 100, 1)
LifeNewPercentage = np.round((LifeNewCount/LifeCount) * 100, 1)
LifeMVNewPercentage = np.round((LifeNewCount/county) * 100, 1)
print("\n")
print("Life Donors")
print("Total Life Count: " + str(LifeCount))
print("Total Life Mean Donation Value: " + str(LifeMean))
print("Percentage of Total Donors: " + str(LifePercentage) + "%")
print('Percentage of Total Life: ' + str(LifeDonorPercentage) + "%")
print("\n")
print("Life Mean (Using Current HV): " + str(LifeOldMean))
print("Life Count (Using Current HV): " + str(LifeOldCount))
print("Life Percentage Of Donating Life (Using Current HV): " + str(LifeOldPercentage) + "%")
print("Life Percentage Of Mid-Value (Using Current HV): " + str(LifeMVOldPercentage) + "%")
print("\n")
print("Life Mean (Using New): " + str(LifeNewMean))
print("Life Count (Using New): " + str(LifeNewCount))
print("Life Percentage Of Donating Life (Using New): " + str(LifeNewPercentage) + "%")
print("Life Percentage Of Mid-Value (Using New): " + str(LifeMVNewPercentage) + "%")

# AW analysis
AWCount = len(np.unique(list(AW["Serial Number"])))
AWMean = np.mean(AW["Sum of Donations"])
AWMax = np.amax(AW["Sum of Donations"])
AWDonorPercentage = np.round((AWCount/count) * 100)
AWPercentage = np.round((AWCount/count) * 100)
AWHVOld = AW[(AW["Sum of Donations"] >= 2000)]
AWHVx = AW[(AW["Sum of Donations"] > quartile_x)]
AWOld = AW[(AW["Sum of Donations"] < 2000) & (AW["Sum of Donations"] >= quartile_y)]
AWOld = AWOld[~AWOld['Serial Number'].isin(AWHVOld['Serial Number'])]
AWNew = AW[(AW["Sum of Donations"] <= quartile_x) & (AW["Sum of Donations"] >= quartile_y)]
AWNew = AWNew[~AWNew['Serial Number'].isin(AWHVx['Serial Number'])]
AWOldMean = np.mean(AWOld["Sum of Donations"])
AWNewMean = np.mean(AWNew["Sum of Donations"])
AWOldCount = len(np.unique(list(AWOld["Serial Number"])))
AWNewCount = len(np.unique(list(AWNew["Serial Number"])))
AWOldPercentage = np.round((AWOldCount/AWCount) * 100, 1)
AWMVOldPercentage = np.round((AWOldCount/count3) * 100, 1)
AWNewPercentage = np.round((AWNewCount/AWCount) * 100, 1)
AWMVNewPercentage = np.round((AWNewCount/county) * 100, 1)
print("\n")
print("AW Donors")
print("Total AW Count: " + str(AWCount))
print("Total AW Mean Donation Value: " + str(AWMean))
print("Percentage of Total Donors: " + str(AWPercentage) + "%")
print('Percentage of Total AW: ' + str(AWDonorPercentage) + "%")
print("\n")
print("AW Mean (Using Current HV): " + str(AWOldMean))
print("AW Count (Using Current HV): " + str(AWOldCount))
print("AW Percentage Of Donating AW (Using Current HV): " + str(AWOldPercentage) + "%")
print("AW Percentage Of Mid-Value (Using Current HV): " + str(AWMVOldPercentage) + "%")
print("\n")
print("AW Mean (Using New): " + str(AWNewMean))
print("AW Count (Using New): " + str(AWNewCount))
print("AW Percentage Of Donating AW (Using New): " + str(AWNewPercentage) + "%")
print("AW Percentage Of Mid-Value (Using New): " + str(AWMVNewPercentage) + "%")

# Creating a dataframe of all the groups
TotalPercentNew = np.round((county/count) * 100)
TotalPercentOld = np.round((count3/count) * 100)
TotalLegacyPercentOld = np.round((LegacyOldCount/LegacyTotal) * 100)
TotalLegacyPercentNew = np.round((LegacyNewCount/LegacyTotal) * 100)
TotalPatronPercentOld = np.round((PatronOldCount/PatronTotal) * 100)
TotalPatronPercentNew = np.round((PatronNewCount/PatronTotal) * 100)
TotalComplimentaryPercentOld = np.round((ComplimentaryOldCount/HonoraryTotal) * 100)
TotalComplimentaryPercentNew = np.round((ComplimentaryNewCount/HonoraryTotal) * 100)
TotalLifePercentOld = np.round((LifeOldCount/LifeTotal) * 100)
TotalLifePercentNew = np.round((LifeNewCount/LifeTotal) * 100)

Names = ("All", 'Other', 'AFS','Legacy', 'Patron', 'Honorary', 'Life', 'Adopt A Wetland')
DonorCount = (count, OtherCount, AFSCount, LegacyCount, PatronCount, ComplimentaryCount, LifeCount, AWCount)
TypeCount = (count, OtherCount, AFSCount, LegacyTotal, PatronTotal, HonoraryTotal, LifeTotal, AWCount)
PercentDonorGroup = (100, 100, 100, LegacyDonorPercentage,  PatronDonorPercentage, ComplimentaryDonorPercentage, LifeDonorPercentage, 100)
MeanValue = (mean, OtherMean, AFSMean, LegacyMean, PatronMean, ComplimentaryMean, LifeMean, AWMean)
max_donations = (max,OtherMax,AFSMax, LegacyMax, PatronMax, ComplimentaryMax, LifeMax, AWMax)
MVCountOld = (count3, OtherOldCount, AFSOldCount, LegacyOldCount,  PatronOldCount, ComplimentaryOldCount, LifeOldCount, AWOldCount)
PercentMVTotalGroupOld = (TotalPercentOld, OtherOldPercentage, AFSOldPercentage, TotalLegacyPercentOld, TotalPatronPercentOld, TotalComplimentaryPercentOld, TotalLifePercentOld, AWOldPercentage)
PercentMVDonorOld = (100, OtherMVOldPercentage, AFSMVOldPercentage, LegacyMVOldPercentage, PatronMVOldPercentage, ComplimentaryMVOldPercentage, LifeMVOldPercentage, AWMVOldPercentage)
MVCountNew = (county, OtherNewCount, AFSNewCount, LegacyNewCount, PatronNewCount, ComplimentaryNewCount, LifeNewCount, AWNewCount)
PercentMVTotalGroupNew = (TotalPercentNew, OtherNewPercentage, AFSNewPercentage, TotalLegacyPercentNew, TotalPatronPercentNew, TotalComplimentaryPercentNew, TotalLifePercentNew, AWNewPercentage)
PercentMVDonorNew = (100, OtherMVNewPercentage, AFSMVNewPercentage, LegacyMVNewPercentage, PatronMVNewPercentage, ComplimentaryMVNewPercentage, LifeMVNewPercentage, AWMVNewPercentage)

MVDAnalysis = pd.DataFrame(list(zip(Names, DonorCount, TypeCount, PercentDonorGroup, MeanValue, max_donations, MVCountOld, PercentMVTotalGroupOld, PercentMVDonorOld, MVCountNew, PercentMVTotalGroupNew, PercentMVDonorNew)),
                           columns =['Type of Donor', 'Count of Donors', 'Total Count of Type Group', 'Percent of Total Type Group', 'Mean Donation Value', 'Max Donation Value', 'Mid Value Count (Using Current HV)', 
                                     'Percent MV  of Total Type Group (Using Current HV)', 'Percent of MV Donors (Using Current HV)', 
                                     'Mid Value Count (Using New HV)', 'Percent MV of Total Type Group (Using New HV)', 
                                     'Percent of MV Donors (Using New HV)'])

MVDAnalysis = MVDAnalysis.reset_index()
MVDAnalysis['Mean Donation Value'] = MVDAnalysis['Mean Donation Value'].round(decimals = 2)
MVDAnalysis['Max Donation Value'] = MVDAnalysis['Max Donation Value'].round(decimals = 2)
MVDAnalysis = MVDAnalysis.sort_values(by='Mean Donation Value')

#print(MVDAnalysis)

# Active / Lapsed Donors
FullySNAc = FullySN.loc[FullySN['Year'] >= 2020].sort_values('Year').drop_duplicates('Serial Number')
FullySNLaW = FullySN.loc[FullySN['Year'] >= 2017].sort_values('Year').drop_duplicates('Serial Number')
FullySNLaW = FullySNLaW[~FullySNLaW['Serial Number'].isin(FullySNAc['Serial Number'])]
FullySNLa = FullySN.loc[FullySN['Year'] < 2017].sort_values('Year')
FullySNLa = FullySNLa[~FullySNLa['Serial Number'].isin(FullySNAc['Serial Number'])].drop_duplicates('Serial Number')
FullySNLa = FullySNLa[~FullySNLa['Serial Number'].isin(FullySNLaW['Serial Number'])].drop_duplicates('Serial Number')

FullMVAc = FullMV.loc[FullMV['Year'] >= 2020].sort_values('Year').drop_duplicates('Serial Number')
FullMVLaW = FullMV.loc[FullMV['Year'] >= 2017].sort_values('Year').drop_duplicates('Serial Number')
FullMVLaW = FullMVLaW[~FullMVLaW['Serial Number'].isin(FullMVAc['Serial Number'])]
FullMVLa = FullMV.loc[FullMV['Year'] < 2017].sort_values('Year')
FullMVLa = FullMVLa[~FullMVLa['Serial Number'].isin(FullMVAc['Serial Number'])].drop_duplicates('Serial Number')
FullMVLa = FullMVLa[~FullMVLa['Serial Number'].isin(FullMVLaW['Serial Number'])].drop_duplicates('Serial Number')

ActiveCountNew = len(list(FullySNAc['Serial Number']))
LapsedWarmCountNew = len(list(FullySNLaW['Serial Number']))
LapsedColdCountNew = len(list(FullySNLa['Serial Number']))
TotalPopNew = ActiveCountNew + LapsedWarmCountNew + LapsedColdCountNew

ActiveCountOld = len(list(FullMVAc['Serial Number']))
LapsedWarmCountOld = len(list(FullMVLaW['Serial Number']))
LapsedColdCountOld = len(list(FullMVLa['Serial Number']))
TotalPopOld = ActiveCountOld + LapsedWarmCountOld + LapsedColdCountOld

# # Analysing Donation Amount
# fig, ax = plt.subplots(3,1)
# labels = range(0,2100,100)
# values, bins, bars = plt.subplot(3,1,1,).hist(Full["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of All Donors')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=6)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,1,2).hist(FullySN["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Mid-Value (Proposed)')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=6)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,1,3).hist(FullMV["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Mid-Value (Current)')
# plt.xlabel('Donation Amount Bin')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars,fontsize=6) 
# plt.xticks(labels)

# fig, ax = plt.subplots(4, 3)
# labels = range(0,2200,200)
# values, bins, bars = plt.subplot(3,3,1).hist(Other["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Other')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,3,2).hist(AFS["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of AFS')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,3,3).hist(Legacy["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Legacy')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,3,4).hist(Patron["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Patrons')
# plt.xlabel('Donation Amount Bin')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,3,5).hist(Complimentary["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Honorary')
# plt.xlabel('Donation Amount Bin')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# values, bins, bars = plt.subplot(3,3,6).hist(Life["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Life')
# plt.xlabel('Donation Amount Bin')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# plt.subplot(3,3,7).set_axis_off()
# values, bins, bars = plt.subplot(3,3,8).hist(AW["Sum of Donations"], bins=range(0,2000,100))
# plt.title('Distribution of Adopt A Wetland')
# plt.xlabel('Donation Amount Bin')
# plt.ylabel('Frequency of Donations')
# plt.bar_label(bars, fontsize=7)
# plt.xticks(labels)
# plt.subplot(3,3,9).set_axis_off()
# plt.subplots_adjust(left=0.1,
#                     bottom=0.1,
#                     right=0.9,
#                     top=0.9,
#                     wspace=0.4,
#                     hspace=0.4)

# fig = plt.figure(figsize=(18,14))
# ax1 = fig.add_subplot(1,2,1, label='1')
# plt.bar(MVDAnalysis['Type of Donor'], MVDAnalysis['Max Donation Value'])
# plt.title('Max Donation Value By Donor Group')
# plt.xlabel('Type of Donor')
# plt.ylabel('Donation Value (£)')
# plt.ylim(0, 2100)
# plt.axhline(y=quartile_y, color='r', linestyle='-')
# plt.axhline(y=quartile_x, color='b', linestyle='-')
# plt.axhline(y=2000, color='b', linestyle='-')
# ax1 = fig.add_subplot(1,2,2, label='2')
# plt.bar(MVDAnalysis['Type of Donor'], MVDAnalysis['Mean Donation Value'])
# plt.title('Mean Donation Value By Donor Group')
# plt.xlabel('Type of Donor')
# plt.ylim(0, 2100)
# #for index in range(len(MVDAnalysis['Type of Donor'])):
#     #ax.text(MVDAnalysis['Type of Donor'][index],MVDAnalysis['Mean Donation Value'][index],MVDAnalysis['Mean Donation Value'][index], size=12)
# plt.axhline(y=quartile_y, color='r', linestyle='-')
# plt.text(3, quartile_y, 'Mid-Value Lower Limit', fontsize=10, va='center', ha='center', backgroundcolor='w')
# plt.axhline(y=quartile_x, color='b', linestyle='-')
# plt.text(3, quartile_x, 'Mid-Value Upper Limit (Proposed)', fontsize=10, va='center', ha='center', backgroundcolor='w')
# plt.axhline(y=2000, color='b', linestyle='-')
# plt.text(3, 2000, 'Mid-Value Upper Limit (Current)', fontsize=10, va='center', ha='center', backgroundcolor='w')

# # Active vs Lapsed

# fig = plt.figure(figsize=(12,8))
# ax1 = fig.add_subplot(1,2,1)
# y = ([ActiveCountNew,LapsedWarmCountNew, LapsedColdCountNew])
# labels = ['Active', 'Warm Lapsed','Cold Lapsed']
# ax1.pie(y, labels = labels, autopct='%.2f')
# plt.title('Mid Value Donors (New) (Active vs Lapsed)')
# ax2 = fig.add_subplot(1,2,2)
# y = ([ActiveCountOld,LapsedWarmCountOld, LapsedColdCountOld])
# ax2.pie(y, labels = labels, autopct='%.2f')
# plt.title('Mid Value Donors (Current) (Active vs Lapsed)')

# # # Source and Destination
FullDestYr = FullSource[FullSource['Year'] >= 2020].groupby(['Serial Number', 'Destination']).agg({'Sum of Donations':'sum'}).reset_index()
FullDestYrAcN = FullDestYr[FullDestYr['Serial Number'].isin(FullySNAc['Serial Number'])]
FullDestYrAcNCount = FullDestYrAcN.value_counts(['Destination'])
FullDestYrAcNLabels = FullDestYrAcN['Destination'].unique()
FullDestYrAcAnonN = FullDestYrAcN.groupby(['Destination']).agg({'Sum of Donations':'sum'}).sort_values('Sum of Donations').reset_index()

FullDestYr = FullSource[FullSource['Year'] >= 2020].groupby(['Serial Number', 'Destination']).agg({'Sum of Donations':'sum'}).reset_index()
FullDestYrAcC = FullDestYr[FullDestYr['Serial Number'].isin(FullMV['Serial Number'])]
FullDestYrAcCCount = FullDestYrAcC.value_counts(['Destination'])
FullDestYrAcCLabels = FullDestYrAcC['Destination'].unique()
FullDestYrAcAnonC = FullDestYrAcC.groupby(['Destination']).agg({'Sum of Donations':'sum'}).sort_values('Sum of Donations').reset_index()

# fig = plt.figure(figsize=(10, 20))
# ax1 = fig.add_subplot(1,2,1)
# plt.barh(FullDestYrAcAnonC['Destination'],FullDestYrAcAnonC['Sum of Donations'])
# plt.title('Donation Amounts By Destinations (Mid-Value Donors - Current) (2020+)')
# plt.ylabel('Destination')
# plt.xlabel('Sum of Donation (£)')
# plt.xticks(fontsize=7)
# plt.yticks(fontsize=7)
# plt.tick_params(axis='y', which='major', pad=15)
# ax2 = fig.add_subplot(1,2,2)
# plt.barh(FullDestYrAcAnonN['Destination'],FullDestYrAcAnonN['Sum of Donations'])
# plt.title('Donation Amounts By Destinations (Mid-Value Donors - New) (2020+)')
# plt.xlabel('Sum of Donation (£)')
# plt.xticks(fontsize=7)
# plt.yticks(fontsize=7)
# plt.tick_params(axis='y', which='major', pad=15)

# fig = plt.figure(figsize=(10, 20))
# ax1 = fig.add_subplot(1,2,1)
# plt.barh(FullDestYrAcCLabels,FullDestYrAcCCount)
# plt.title('No. Donors By Destinations (Mid-Value Donors - Current) (2020+)')
# plt.ylabel('Destination')
# plt.xlabel('No. of Donors')
# plt.xticks(fontsize=7)
# plt.yticks(fontsize=7)
# plt.tick_params(axis='y', which='major', pad=15)
# ax2 = fig.add_subplot(1,2,2)
# plt.barh(FullDestYrAcNLabels,FullDestYrAcNCount)
# plt.title('No. Donors By Destinations (Mid-Value Donors - New) (2020+)')
# plt.xlabel('No. of Donors')
# plt.xticks(fontsize=7)
# plt.yticks(fontsize=7)
# plt.tick_params(axis='y', which='major', pad=15)

#demographics
FullyDemo = FullDemo[FullDemo['Serial Number'].isin(FullySN['Serial Number'])]
FullMVDemo = FullDemo[FullDemo['Serial Number'].isin(FullMV['Serial Number'])]

#gender pie chart
# fig = plt.figure(figsize=(12,8))
# ax1 = fig.add_subplot(1,2,1)
# FullyDemoM = len(FullyDemo[FullyDemo['Gender'] == 'Male'])
# FullyDemoF = len(FullyDemo[FullyDemo['Gender'] == 'Female'])
# FullMVDemoM = len(FullMVDemo[FullMVDemo['Gender'] == 'Female'])
# FullMVDemoF = len(FullMVDemo[FullMVDemo['Gender'] == 'Male'])
# y = (FullyDemoM, FullyDemoF)
# labels = ['Male', 'Female']
# ax1.pie(y, labels = labels, autopct='%.2f')
# plt.title('Mid Value Donors (New) (Gender)')
# ax2 = fig.add_subplot(1,2,2)
# y = (FullMVDemoM, FullMVDemoF)
# ax2.pie(y, labels = labels, autopct='%.2f')
# plt.title('Mid Value Donors (Current) (Gender)')

#nearest centre bar chart
# FullyDemo['NEAREST CENTRE'].value_counts().sort_values().plot(kind='bar', title='Nearest Centre (Mid-Value Donors - New)')
# plt.show()
# FullMVDemo['NEAREST CENTRE'].value_counts().sort_values().plot(kind='bar',title='Nearest Centre (Mid-Value Donors - Current)')
# plt.show()

#age
# FullyDemoAge = FullyDemo.dropna()
# FullyDemoAge['D.O.B'] = pd.DatetimeIndex(FullyDemoAge['D.O.B']).year
# FullMVDemoAge = FullMVDemo.dropna()
# FullMVDemoAge['D.O.B'] = pd.DatetimeIndex(FullMVDemoAge['D.O.B']).year

# FullyDemoAge['D.O.B'].value_counts().sort_index().plot(kind='barh', title='Age (Mid-Value Donors - New)')
# plt.show()
# FullMVDemoAge['D.O.B'].value_counts().sort_index().plot(kind='barh',title='Age (Mid-Value Donors - Current)')
# plt.show()

# FullyDemoFC = FullyDemo.dropna()
# FullyDemoFC['First Contact'] = pd.DatetimeIndex(FullyDemoFC['First Contact']).year
# FullMVDemoFC = FullMVDemo.dropna()
# FullMVDemoFC['First Contact'] = pd.DatetimeIndex(FullMVDemoFC['First Contact']).year

# FullyDemoFC['First Contact'].value_counts().sort_index().plot(kind='barh', title='First Contact (Mid-Value Donors - New)')
# plt.show()
# FullMVDemoFC['First Contact'].value_counts().sort_index().plot(kind='barh',title='First Contact (Mid-Value Donors - Current)')
# plt.show()

# #show plots
# plt.show()

#outputing desired dfs
MVDAnalysis['Mean Donation Value'] = '£' + MVDAnalysis['Mean Donation Value'].astype(str)
MVDAnalysis.iloc[:,[4,8,9,11,12]] = MVDAnalysis.iloc[:,[4,8,9,11,12]].astype(str) + '%'
file_name = 'Mid Value Donors ' + str(date.today()) + '.xlsx'
with pd.ExcelWriter(file_name) as writer:
    MVDAnalysis.to_excel(writer, sheet_name="MVDAnalysis")
    FullySNAc.to_excel(writer, sheet_name="Active (New)")
    FullySNLaW.to_excel(writer, sheet_name="Warm Lapsed (New)")
    FullySNLa.to_excel(writer, sheet_name="Cold Lapsed (New)")
    FullMVAc.to_excel(writer, sheet_name="Active (Current)")
    FullMVLaW.to_excel(writer, sheet_name="Warm Lapsed (Current)")
    FullMVLa.to_excel(writer, sheet_name="Cold Lapsed (Current)")
    FullDestYrAcN.to_excel(writer, sheet_name="Destinations (New)")
    FullDestYrAcC.to_excel(writer, sheet_name="Destinations (Current)")
    FullyDemo.to_excel(writer, sheet_name="Demographics (MV New)")
    FullMVDemo.to_excel(writer, sheet_name="Demographics (MV Current)")
print('Successfully Exported')
    

# Archived
# partitioning the df by =< HV
# Full2 = Full[Full["Sum of Donations"] <= quartile_x]
# max2 = np.amax(Full2["Sum of Donations"])
# min2 = np.amin(Full2["Sum of Donations"])
# stdev2 = np.std(Full2["Sum of Donations"])
# mean2 = np.mean(Full2["Sum of Donations"])
# median2 = np.median(Full2["Sum of Donations"])
# count2 = np.count_nonzero(Full2["Serial Number"])
# print("\n")
# print("Full2: <HV")
# print("Standard Deviation: " + str(stdev2))
# print("Mean: " + str(mean2))
# print("Max: " + str(max2))
# print("Median: " + str(median2))
# print("Min: " + str(min2))
# print("Count of donors: " + str(count2))

#partitioning the df by =< HV and >= LQ
# Full4 = Full[Full["Sum of Donations"] <= quartile_x]
# Full4 = Full4[Full4["Sum of Donations"] >= lower_quartile]
# max4 = np.amax(Full4["Sum of Donations"])
# min4 = np.amin(Full4["Sum of Donations"])
# stdev4 = np.std(Full4["Sum of Donations"])
# mean4 = np.mean(Full4["Sum of Donations"])
# median4 = np.median(Full4["Sum of Donations"])
# count4 = np.count_nonzero(Full4["Serial Number"])
# print("\n")
# print("Full4: <HV, >LQ")
# print("Standard Deviation: " + str(stdev4))
# print("Mean: " + str(mean4))
# print("Max: " + str(max4))
# print("Median: " + str(median4))
# print("Min: " + str(min4))
# print("Count of donors: " + str(count4))

#partitioning the df by =< HV and >= UQ
# Full5 = Full[Full["Sum of Donations"] <= quartile_x]
# Full5 = Full5[Full5["Sum of Donations"] >= upper_quartile]
# max5 = np.amax(Full5["Sum of Donations"])
# min5 = np.amin(Full5["Sum of Donations"])
# stdev5 = np.std(Full5["Sum of Donations"])
# mean5 = np.mean(Full5["Sum of Donations"])
# median5 = np.median(Full5["Sum of Donations"])
# count5 = np.count_nonzero(Full5["Serial Number"])
# print("\n")
# print("Full5: <HV, >UQ")
# print("Standard Deviation: " + str(stdev5))
# print("Mean: " + str(mean5))
# print("Max: " + str(max5))
# print("Median: " + str(median5))
# print("Min: " + str(min5))
# print("Count of donors: " + str(count5))

#partitioning the df by =< HV(2000) and >= LV(120)
# Full6 = Full[Full["Sum of Donations"] <= 2000]
# Full6 = Full6[Full6["Sum of Donations"] >= 130]
# max6 = np.amax(Full6["Sum of Donations"])
# min6 = np.amin(Full6["Sum of Donations"])
# stdev6 = np.std(Full6["Sum of Donations"])
# mean6 = np.mean(Full6["Sum of Donations"])
# median6 = np.median(Full6["Sum of Donations"])
# count6 = np.count_nonzero(Full6["Serial Number"])
# print("\n")
# print("Full6: <2000, >120")
# print("Standard Deviation: " + str(stdev6))
# print("Mean: " + str(mean6))
# print("Max: " + str(max6))
# print("Median: " + str(median6))
# print("Min: " + str(min6))
# print("Count of donors: " + str(count6))