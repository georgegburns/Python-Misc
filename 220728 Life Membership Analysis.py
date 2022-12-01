
from asyncio.windows_events import NULL
from multiprocessing.sharedctypes import Value
from optparse import Values
import pandas as pd
import numpy as np
from matplotlib import pyplot as plt
from matplotlib.axis import XAxis, YAxis
from cProfile import label
from textwrap import wrap
from re import X
from tkinter import HIDDEN


# #read_excel and filter dfs
LM = pd.read_excel('Life Membership Analysis Output (SQL).xlsx', sheet_name='LM')
V = pd.read_excel('Life Membership Analysis Output (SQL).xlsx', sheet_name='Visits')
TM = pd.read_excel('Life Membership Analysis Output (SQL).xlsx', sheet_name='LM vs Total M')
LM_Ac = LM[(LM['Dead?'].fillna('Contactable').str.contains('Deceased') == False)
        & (LM['Status'].fillna('Unknown').str.contains('Active'))]
LM_Ac_2014 = LM_Ac[LM_Ac['Year Became LM'] > 2013]
LM_Ac_PM = LM_Ac_2014[LM_Ac_2014['Member/New'].str.contains('New') == False]
LM_Ac_N = LM_Ac_2014[LM_Ac_2014['Member/New'].str.contains('New') == True]
LM_AC_Con =  LM_Ac_2014[(LM_Ac_2014['Previous Concession'].str.contains('New') == False)
        & (LM_Ac_2014['Previous Concession'].str.contains('Unknown') == False)]
LM_AC_Em = LM_Ac_2014[(LM_Ac_2014['Email Opt In'].str.contains('Emailable') == True) & (LM_Ac_2014['Email Opt In'].str.contains('Not Emailable') == False)]
# LM_TIME = LM.loc[LM['Year Became LM'] < 2017, 'Year Became LM'] = 2016

# adjust these to create custom time period df
start_date = 2000
end_date = 2023
LM_Ac_X = LM_Ac[(LM_Ac['Year Became LM'] > start_date)
            & (LM_Ac['Year Became LM'] < end_date)]

# #head, count & dtypes
print(LM.head(5), LM.dtypes)
# print(LM_Ac.head(5), LM_Ac['Serial Number'].count())
# print(LM_Ac_2014.head(5), LM_Ac_2014['Serial Number'].count())
# print(LM_Ac_PM.head(5), LM_Ac_PM['Serial Number'].count())
# print(LM_Ac_N.head(5), LM_Ac_N['Serial Number'].count())

# # Previous member

# # the membership type of life members who were members prior
# CMT_Seg = LM_AC_Con.groupby(['Concession / Non-Concession','Previous Concession'])['Concession / Non-Concession','Previous Concession'].value_counts().unstack('Previous Concession').fillna(0)
# ax = CMT_Seg.plot(kind = 'bar', stacked = True, figsize=(14,8),
#     title = 'Concession vs Non-Concession (2014 - Now)',
#     rot = 0, fontsize = 7, legend = 'Previous Concession Status')
# ax.set_ylabel('Count')
# ax.set_xlabel('Current Concession Status')
# ax.bar_label(ax.containers[0], fmt = '%1.0f',
#     label_type = 'center', color = 'snow')
# ax.bar_label(ax.containers[1], fmt = '%1.0f',
#     label_type = 'center', color = 'snow')
# ax.bar_label(ax.containers[-1],padding = 1)
# ax.legend(title = 'Previous Concession Status')
# plt.show()

# # the % of life members who were members prior against new
ax = LM_Ac_2014['Member/New'].value_counts().plot(kind = 'pie', y = 'Member/New',
    autopct='%1.0f%%', figsize=(14,8), title = 'Member vs New (2014 - Now)',
    rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True,
    explode=[0,0.05])
ax.set_ylabel(None)
ax.annotate("Total Population: " + str(LM_Ac_2014['Member/New'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
plt.show()

# # the life membership types of  previous members

# PM = LM_Ac_PM['Membership Type'].value_counts()
# N = LM_Ac_N['Membership Type'].value_counts()
# X = LM_Ac_2014['Membership Type'].unique()
# LM_Ac_PMvN_2014 = pd.DataFrame({'Previous Members' : PM, 'New Members' : N},
#  index = X)
# ax = LM_Ac_PMvN_2014.plot.bar(figsize = (14,8), rot = 0,
#     title = 'Membership Types By New and Previous Members (Active 2014 - Now)',
#     fontsize = 7)
# ax.bar_label(ax.containers[0], padding = 1)
# ax.bar_label(ax.containers[-1], padding = 1)
# plt.show()


# # Previous members vs new donations

# ax = LM.groupby('Member/New')['Donations (Last Year)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Previous Members/New (1Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel(None)
# ax.bar_label(ax.containers[-1], fmt = '£%.2f' , padding = 1)
# plt.show()

# ax = LM.groupby('Member/New')['Donations (Last 2 Years)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Previous Members/New (2Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel(None)
# ax.bar_label(ax.containers[-1], fmt = '£%.2f' , padding = 1)
# plt.show()

# ax = LM.groupby('Member/New')['Donations (Last 4 Years)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Previous Members/New (4Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel(None)
# ax.bar_label(ax.containers[-1], fmt = '£%.2f' , padding = 1)
# plt.show()

# # # Segment - IGNORE until new segmentation has been inputted

# # # life members by segments
# # ax = LM['Segment'].value_counts().plot(kind = 'bar',
# #     figsize=(14,8), title = 'Life Member Segments',
# #     rot = 0, fontsize = 8)
# # ax.set_xlabel('Segment')
# # ax.set_ylabel('Count')
# # ax.bar_label(ax.containers[0], padding = 1)
# # plt.show()

# # # life member type by segment
# # MT_Seg = LM.groupby(['Membership Type', 'Segment'])['Membership Type', 'Segment'].value_counts().unstack('Segment').fillna(0)
# # ax = MT_Seg.plot(kind = 'barh', stacked = True,
# #     figsize=(14,8), title = 'Life Membership Types',
# #     rot = 0, fontsize = 7)
# # ax.set_xlabel('Count')
# # ax.set_ylabel('Membership Type')
# # ax.bar_label(ax.containers[-1],padding = 1)
# # plt.show()

# # Gender

# ax = LM['Gender'].value_counts().plot(kind = 'pie',
#     y = 'Gender', figsize=(14,8), autopct='%1.0f%%',
#     title = 'Life Membership Gender',
#     rot = 0, fontsize = 12, textprops={'color':'w'},
#     legend = True, explode=[0,0,0.05])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Gender'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# #  Age
# ax = LM['Age > 65'].value_counts().plot(kind = 'pie',
#     y = 'Age', figsize=(14,8), autopct='%1.0f%%',
#     title = 'Life Membership Age',
#     rot = 0, fontsize = 12, textprops={'color':'w'},
#     legend = True, explode=[0,0.05,0])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Age > 65'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# ax = LM['Age Became LM > 65'].value_counts().plot(kind = 'pie',
#     y = 'Age', figsize=(14,8), autopct='%1.0f%%',
#     title = 'Age At Becoming Life Member',
#     rot = 0, fontsize = 12, textprops={'color':'w'},
#     legend = True, explode=[0,0,0.05])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Age Became LM > 65'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# # Nearest site
# ax = LM_Ac['Nearest Site'].value_counts().plot(kind = 'bar',
#     figsize=(14,8), title = 'Life Membership Centre',
#     rot = 0, fontsize = 7)
# ax.set_xlabel('Centre')
# ax.set_ylabel('Count')
# ax.bar_label(ax.containers[0],padding = 1)
# plt.show()

# # Nearest site donations

# ax = LM.groupby('Nearest Site')['Donations (Last Year)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Nearest Site (1Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Centre')
# ax.bar_label(ax.containers[-1], fmt = '£%.2f', padding = 1)
# plt.show()

# ax = LM.groupby('Nearest Site')['Donations (Last 2 Years)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Nearest Site (2Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Centre')
# ax.bar_label(ax.containers[-1], fmt = '£%.2f', padding = 1)
# plt.show()

# ax = LM.groupby('Nearest Site')['Donations (Last 4 Years)'].sum().sort_values().plot(
#     kind = 'bar', figsize=(14,8),
#     title = 'Donations By Nearest Site (4Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Centre')
# ax.bar_label(ax.containers[-1], fmt = '£%.2f', padding = 1)
# plt.show()

# # Membership type

ax = LM_Ac['Membership Type'].value_counts().plot(kind = 'bar',
    figsize=(14,8), title = 'Life Membership Types',
    rot = 0, fontsize = 7)
ax.set_xlabel('Type')
ax.set_ylabel('Count')
ax.bar_label(ax.containers[0],padding = 1)
plt.show()

# # Membership type donations

# MT_DO_1Y = LM.groupby('Membership Type')['Donations (Last Year)'].sum().sort_values()
# ax = MT_DO_1Y.plot(kind = 'bar', figsize=(14,8),
#     title = 'Donations By Life Membership Types (1Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Membership Type')
# ax.bar_label(ax.containers[0], fmt = '£%.2f', padding = 1)
# plt.show()

# MT_DO_2Y = LM.groupby('Membership Type')['Donations (Last 2 Years)'].sum().sort_values()
# ax = MT_DO_2Y.plot(kind = 'bar', figsize=(14,8),
#     title = 'Donations By Life Membership Types (2Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Membership Type')
# ax.bar_label(ax.containers[0], fmt = '£%.2f', padding = 1)
# plt.show()

# MT_DO_4Y = LM.groupby('Membership Type')['Donations (Last 4 Years)'].sum().sort_values()
# ax = MT_DO_4Y.plot(kind = 'bar', figsize=(14,8),
#     title = 'Donations By Life Membership Types (4Y)',
#     rot = 0, fontsize = 7)
# ax.set_ylabel('Amount (£)')
# ax.set_xlabel('Membership Type')
# ax.bar_label(ax.containers[0], fmt = '£%.2f', padding = 1)
# plt.show()

# # Email Collected
# ax = LM_Ac_2014['Email Opt In'].value_counts().plot(kind = 'pie', y = 'Email Opt In',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Email Collected (2014 - Now)',
#     rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True,
#     explode=[0.05,0,0])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM_Ac_2014['Email Opt In'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# # Email Collected PM vs N

# Em_MN = LM_Ac_2014.groupby(['Member/New','Email Opt In'])['Member/New','Email Opt In'].value_counts().unstack('Email Opt In').fillna(0)
# ax = Em_MN.plot(kind = 'bar', figsize=(14,8), stacked = True,
#         title = 'Email Collected (2014 - Now) (Previous Members vs New)',
#         rot = 0, fontsize = 12, legend = True)
# ax.bar_label(ax.containers[0], fmt = '%1.0f',
#     label_type = 'center', color = 'snow')
# ax.bar_label(ax.containers[1], fmt = '%1.0f',
#     label_type = 'center', color = 'snow')
# ax.bar_label(ax.containers[2], fmt = '%1.0f',
#     label_type = 'center', color = 'snow')
# ax.bar_label(ax.containers[-1],padding = 1)
# ax.set_ylabel('Count')
# plt.show()

# # Email Opt In (Updates/News)
# ax = LM_Ac_2014['LM UN'].value_counts().plot(kind = 'pie', y = 'LM UN',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Updates/News Opt In (2014 - Now)',
#     rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True)
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM_Ac_2014['Email Opt In'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# ax = LM_Ac_2014['Legacies (email)'].value_counts().plot(kind = 'pie', y = 'Legacies (email)',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Updates/News Opt In (2014 - Now)',
#     rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True)
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM_Ac_2014['Email Opt In'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# # Email Collected by Email Opt In
# ax = LM_AC_Em['LM UN'].value_counts().plot(kind = 'pie', y = 'LM UN',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Updates/News Opt In (Emailable) (2014 - Now)',
#     rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True)
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM_AC_Em['LM UN'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# ax = LM_Ac['LM UN'].value_counts().plot(kind = 'pie', y = 'LM UN',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Updates/News Opt In (Emailable)',
#     rot = 0, fontsize = 12, textprops={'color':'w'}, legend = True)
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM_Ac['LM UN'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()


# # Email Opt In Age

Em_age = LM_Ac_2014.groupby(['Age to Nearest Decade','Email Opt In'])['Age to Nearest Decade','Email Opt In'].value_counts().unstack('Email Opt In').fillna(0)
ax = Em_age.plot(kind = 'bar', figsize=(14,8), stacked = True,
        title = 'Email Opt In (2014 - Now) (Age)',
        rot = 0, fontsize = 12, legend = True)
ax.set_ylabel('Count')
plt.show()

# # Timelines

# # overall timeline

# ax = LM_Ac_2014['Year Became LM'].value_counts().sort_index().plot(
#       kind = 'bar', figsize = (25,8),
#       title = 'Timeline Of Life Membership',
#       rot = 0, fontsize = 7, legend = False)
# ax.set_xlabel('Year')
# ax.set_ylabel('Recruited Members')
# ax.bar_label(ax.containers[0],padding = 1)
# plt.show()

# # volunteer, legacy and deceased

# ax = LM['Volunteer'].value_counts().plot(kind = 'pie', y = 'Volunteer',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Volunteers',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.1])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Volunteer'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# ax = LM['Legacy'].value_counts().plot(kind = 'pie', y = 'Legacy',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Legacy',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.1])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Legacy'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# ax = LM['Relation to LM'].value_counts().plot(kind = 'pie', y = 'Relation to LM',
#     autopct='%1.0f%%', figsize=(14,8), title = 'When LM Pledged Legacy',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.1,0])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Relation to LM'].replace({'':np.NAN}).count()), xy = [1,-0.7])
# plt.show()

# ax = LM['Dead?'].value_counts().plot(kind = 'pie', y = 'Deceased',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Deceased',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.05])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Dead?'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# # Visitation

# V_Site_Year = V.groupby(['site','Year'])['total_visits'].sum().unstack('site').fillna(0)
# ax = V_Site_Year.plot(
#         kind = 'bar', figsize = (14,8),
#         rot = 0, stacked = True,
#         title = 'Visits Per Site By Year',
#         fontsize = 9)
# ax.set_xlabel('Year')
# ax.set_ylabel('Visitation Count')
# ax.legend(title = 'Site')
# plt.show()

# ax = LM['Visited Last 5 Years?'].value_counts().plot(kind = 'pie', y = 'Visited',
#     autopct='%1.0f%%', figsize=(14,8), title = 'LM Visited Last 5 Years',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.05])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Visited Last 5 Years?'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()

# # proportion of LM to total M

# TM.plot(x = 'Site', y = ['Life Members', 'Total Members (Excluding Life)'], kind = 'bar',
#         stacked = True, figsize =(14,8),
#         title = 'Proportion of Life Members to Total Membership by Site',
#         rot = 0, fontsize = 10)
# ax = TM['Proportion of Membership'].plot(secondary_y = True, color = 'k')
# plt.show()

# # Donations

# ax = LM['Donates'].value_counts().plot(kind = 'pie', y = 'Donates',
#     autopct='%1.0f%%', figsize=(14,8), title = 'Donated Since 2018',
#     rot = 0, fontsize = 10, textprops={'color':'w'}, legend = True,
#     explode=[0,0.05])
# ax.set_ylabel(None)
# ax.annotate("Total Population: " + str(LM['Donates'].replace({'':np.NAN}).count()), xy = [0.9,-0.7])
# plt.show()