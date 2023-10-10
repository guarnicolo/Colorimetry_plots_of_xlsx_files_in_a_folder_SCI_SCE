#!/usr/bin/env python
# coding: utf-8

# # Colorimetry_plots_of_all_xlsx_files_in_a_folder
# this script plots all colorimetry data from all xlsx files in a folder.
# The folder should be named with the name of the testing region.

# In[2]:


import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.gridspec as gridspec
import numpy as np
import os

#input
##################
region_name='OM'
filename=region_name+' pre e post'
path=r'.'
folder_path=path+'\\'+filename
##################

#to check all files inside the foldes activate following line:
os.listdir(folder_path)

# to create a list name spa_files containing all .spa files present in the folder_path
xlsx_files = [f for f in os.listdir(folder_path) if f.endswith('.xlsx')]

#check of the elemnets of the spa_files list
print(f'Number and list of .xlsx files in the folder_path: {len(xlsx_files)}')
print(xlsx_files)


# In[10]:


#define fixed scale for plts (useful for fast comparison)
a_star_min, a_star_max = 10, 30
b_star_min, b_star_max = 30, 60
L_star_min, L_star_max = 20, 50


# # L* and a* b* plots

# In[4]:


for xlsx_file in xlsx_files:
    
    # open xlsx files
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

    # take data from specified columns from PRE sheet
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

    # take data from specified columns from POST sheet
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]
    
    #creating this dataframes useful later for plotting L*
    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

    #defining list of markers and colors
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

    
    #starting plotting defining fig size
    fig = plt.figure(figsize=(10, 8))

    #defining a grid of 1 row and 2 coulums for a*b* plot and L* plot
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    #first plot for a*b* coordinates, with error bars
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(a_star_pre, b_star_pre, std_a_pre, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(a_star_post, b_star_post, std_a_post, std_b_post, names_post)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    # connecting the corresponding points from PRE and POST with solid lines
    for x_pre, y_pre, x_post, y_post, color in zip(a_star_pre, b_star_pre, a_star_post, b_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)


    # add labels and a legend to the a*b* subplot
    ax1.set_xlabel('a* [CIELAB]', fontsize=15)
    ax1.set_ylabel('b* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    # Second subplot for L* plot
    ax2 = plt.subplot(gs[1])
   
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, L_star_pre, std_a_pre, std_L_pre, names_pre)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, L_star_post, std_a_post, std_L_post, names_post)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    # connect the corresponding points from PRE and POST
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, L_star_pre, ones, L_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    # define axis limits for L*
    x_min, x_max = -0.3, 1.3 
    ax2.set_xlim(x_min, x_max)

    # add labels to the second subplot
    ax2.set_xlabel('')
    ax2.set_ylabel('L* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  # adjust the position of the y-axis label

    # show the plot and save it to a file
    plt.grid(True)
    plt.tight_layout()  # Ensure proper spacing between subplots
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'ab_L.png', dpi=300, bbox_inches='tight')
    plt.show()


# # a* L* and b* plots

# In[5]:


#same script of above but instead plot a* L* coordinates togheter and b* alone

for xlsx_file in xlsx_files:
    
    
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

  
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

  
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]

    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

   
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

   
    
   
    fig = plt.figure(figsize=(10, 8))

    
    
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(a_star_pre, L_star_pre, std_a_pre, std_L_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(a_star_post, L_star_post, std_a_post, std_L_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(a_star_pre, L_star_pre, a_star_post, L_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)



 
    ax1.set_xlabel('a* [CIELAB]', fontsize=15)
    ax1.set_ylabel('L* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    ax2 = plt.subplot(gs[1])
 
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, b_star_pre, zeros, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, b_star_post, zeros, std_b_post, names_post)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

  
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, b_star_pre, ones, b_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

   
    x_min, x_max = -0.3, 1.3  
    ax2.set_xlim(x_min, x_max)

   
    ax2.set_xlabel('')
    ax2.set_ylabel('b* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  

    
    plt.grid(True)
    plt.tight_layout()  
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'aL_b.png', dpi=300, bbox_inches='tight')
    plt.show()


# # b* L* and a* plots

# In[107]:


#same script of above but instead plot b* L* coordinates togheter and a* alone
for xlsx_file in xlsx_files:
    
    
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

   
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

    
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]

    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

    
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

   
    
    
    fig = plt.figure(figsize=(10, 8))

    
    
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(L_star_pre, b_star_pre, std_L_pre, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(L_star_post, b_star_post, std_L_post, std_b_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

   
    for x_pre, y_pre, x_post, y_post, color in zip(L_star_pre, b_star_pre, L_star_post, b_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)


   
    ax1.set_xlabel('L* [CIELAB]', fontsize=15)
    ax1.set_ylabel('b* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    ax2 = plt.subplot(gs[1])
    
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, a_star_pre, zeros, std_a_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, a_star_post, zeros, std_a_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, a_star_pre, ones, a_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    
    x_min, x_max = -0.3, 1.3  
    ax2.set_xlim(x_min, x_max)

   
    ax2.set_xlabel('')
    ax2.set_ylabel('a* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  

    
    plt.grid(True)
    plt.tight_layout()  
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'Lb_a.png', dpi=300, bbox_inches='tight')
    plt.show()


# # scripts with fixed scale

# In[7]:


#same script of above but instead plot L* b* coordinates (with fixed scale) togheter and a* 
for xlsx_file in xlsx_files:
    
   
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

   
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

    
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]

    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

    
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

   
    
    
    fig = plt.figure(figsize=(10, 8))

   
    
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(L_star_pre, b_star_pre, std_L_pre, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(L_star_post, b_star_post, std_L_post, std_b_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(L_star_pre, b_star_pre, L_star_post, b_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    
    
    ax1.set_xlim(L_star_min, L_star_max)
    ax1.set_ylim(b_star_min, b_star_max)
    

   
    ax1.set_xlabel('L* [CIELAB]', fontsize=15)
    ax1.set_ylabel('b* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    ax2 = plt.subplot(gs[1])
    
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, a_star_pre, zeros, std_a_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, a_star_post, zeros, std_a_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

   
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, a_star_pre, ones, a_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    
    x_min, x_max = -0.3, 1.3  
    ax2.set_xlim(x_min, x_max)
    ax2.set_ylim(a_star_min, a_star_max)

    
    ax2.set_xlabel('')
    ax2.set_ylabel('a* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  

    
    plt.grid(True)
    plt.tight_layout()  
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'Lb_a_fixed.png', dpi=300, bbox_inches='tight')
    plt.show()


# In[11]:


#same script of above but instead plot a* b* coordinates (with fixed scale) togheter and L* 
for xlsx_file in xlsx_files:
    
    
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

    
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

    
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]

    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

    
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

    
    
    fig = plt.figure(figsize=(10, 8))

    
    
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(a_star_pre, b_star_pre, std_a_pre, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(a_star_post, b_star_post, std_a_post, std_b_post, names_post)):
        marker = markers[i % len(markers)]  # Cycle through markers
        color = colors[i % len(colors)]  # Cycle through colors
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(a_star_pre, b_star_pre, a_star_post, b_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    
    ax1.set_xlim(a_star_min, a_star_max)
    ax1.set_ylim(b_star_min, b_star_max)

    
    ax1.set_xlabel('a* [CIELAB]', fontsize=15)
    ax1.set_ylabel('b* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    ax2 = plt.subplot(gs[1])
    
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, L_star_pre, std_a_pre, std_L_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, L_star_post, std_a_post, std_L_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, L_star_pre, ones, L_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

   
    x_min, x_max = -0.3, 1.3  
    ax2.set_xlim(x_min, x_max)
    ax2.set_ylim(L_star_min, L_star_max)
    
    
    ax2.set_xlabel('')
    ax2.set_ylabel('L* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  

    
    plt.grid(True)
    plt.tight_layout() 
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'ab_L_fixed.png', dpi=300, bbox_inches='tight')
    plt.show()


# In[12]:


#same script of above but instead plot a* L* coordinates (with fixed scale) togheter and b* 
for xlsx_file in xlsx_files:
    
    
    filepath=folder_path+"/"+xlsx_file
    data_pre = pd.read_excel(filepath, sheet_name=region_name+'_PRE')
    data_post = pd.read_excel(filepath, sheet_name=region_name+'_POST')

    
    a_star_pre = data_pre['a_av']
    b_star_pre = data_pre['b_av']
    L_star_pre = data_pre['L_av']
    std_a_pre = data_pre['a_devst']
    std_b_pre = data_pre['b_devst']
    std_L_pre = data_pre['L_devst']
    names_pre = data_pre.iloc[:, 1]

    
    a_star_post = data_post['a_av']
    b_star_post = data_post['b_av']
    L_star_post = data_post['L_av']
    std_a_post = data_post['a_devst']
    std_b_post = data_post['b_devst']
    std_L_post = data_post['L_devst']
    names_post = data_post.iloc[:, 1]

    zeros = np.zeros(len(L_star_pre))
    ones = np.ones(len(L_star_pre))

    
    markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
    colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

   
    
    
    fig = plt.figure(figsize=(10, 8))

    
    
    gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
    ax1 = plt.subplot(gs[0])
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(a_star_pre, L_star_pre, std_a_pre, std_L_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre, y_pre, xerr=std_x_pre, yerr=std_y_pre, fmt=marker, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(a_star_post, L_star_post, std_a_post, std_L_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post, y_post, xerr=std_x_post, yerr=std_y_post, fmt=marker, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(a_star_pre, L_star_pre, a_star_post, L_star_post, colors):
        ax1.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    ax1.set_xlim(a_star_min, a_star_max)
    ax1.set_ylim(L_star_min, L_star_max)

    
    ax1.set_xlabel('a* [CIELAB]', fontsize=15)
    ax1.set_ylabel('L* [CIELAB]', fontsize=15)
    handles, labels = ax1.get_legend_handles_labels()
    ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.2), ncol=4, fontsize=9,  framealpha=0.8)
    
    ax2 = plt.subplot(gs[1])
    
    for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
            zip(zeros, b_star_pre, zeros, std_b_pre, names_pre)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

    for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
            zip(ones, b_star_post, zeros, std_b_post, names_post)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)

    
    for x_pre, y_pre, x_post, y_post, color in zip(zeros, b_star_pre, ones, b_star_post, colors):
        ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)

    
    x_min, x_max = -0.3, 1.3  
    ax2.set_xlim(x_min, x_max)
    ax2.set_ylim(b_star_min, b_star_max)
    
    ax2.set_xlabel('')
    ax2.set_ylabel('b* [CIELAB]', fontsize=15)
    ax2.set_xticks([])

    ax2.yaxis.set_label_coords(-0.4, 0.5)  

   
    plt.grid(True)
    plt.tight_layout()  
    plt.savefig(folder_path+ '/' + xlsx_file[:-5]+'aL_b_fixed.png', dpi=300, bbox_inches='tight')
    plt.show()


# # Plot SCI and SCE data in the same graph

# In[3]:


# Load data from both SCI and SCE xlsx file
filepathSCE=folder_path+"\\"+filename+'_SCE.xlsx'
filepathSCI=folder_path+"\\"+filename+'_SCI.xlsx'
data_pre_SCE = pd.read_excel(filepathSCE, sheet_name=region_name+'_PRE')
data_post_SCE = pd.read_excel(filepathSCE, sheet_name=region_name+'_POST')
data_pre_SCI = pd.read_excel(filepathSCI, sheet_name=region_name+'_PRE')
data_post_SCI = pd.read_excel(filepathSCI, sheet_name=region_name+'_POST')

# Extract data from columns SCE pre and post
a_star_pre_sce = data_pre_SCE['a_av']
b_star_pre_sce = data_pre_SCE['b_av']
L_star_pre_sce = data_pre_SCE['L_av']
std_a_pre_sce = data_pre_SCE['a_devst']
std_b_pre_sce = data_pre_SCE['b_devst']
std_L_pre_sce = data_pre_SCE['L_devst']
names_pre_sce = data_pre_SCE.iloc[:, 1]


a_star_post_sce = data_post_SCE['a_av']
b_star_post_sce = data_post_SCE['b_av']
L_star_post_sce = data_post_SCE['L_av']
std_a_post_sce = data_post_SCE['a_devst']
std_b_post_sce = data_post_SCE['b_devst']
std_L_post_sce = data_post_SCE['L_devst']
names_post_sce = data_post_SCE.iloc[:, 1]
    
# Extract data from columns SCI pre and post
a_star_pre_sci = data_pre_SCI['a_av']
b_star_pre_sci = data_pre_SCI['b_av']
L_star_pre_sci = data_pre_SCI['L_av']
std_a_pre_sci = data_pre_SCI['a_devst']
std_b_pre_sci = data_pre_SCI['b_devst']
std_L_pre_sci = data_pre_SCI['L_devst']
names_pre_sci = data_pre_SCI.iloc[:, 1]


a_star_post_sci = data_post_SCI['a_av']
b_star_post_sci = data_post_SCI['b_av']
L_star_post_sci = data_post_SCI['L_av']
std_a_post_sci = data_post_SCI['a_devst']
std_b_post_sci = data_post_SCI['b_devst']
std_L_post_sci = data_post_SCI['L_devst']
names_post_sci = data_post_SCI.iloc[:, 1]
    
zeros = np.zeros(len(L_star_pre_sci))
ones = np.ones(len(L_star_pre_sci))


markers = ['o', 's', '^', 'P', 'v', '>', '<', 'X', 'p', '*']
colors = ['b', 'g', 'r', 'c', 'm', 'y', 'k', 'darkred', 'darkgreen', 'darkslategrey', 'navy', 'orangered', 'grey', 'olive']

    

fig = plt.figure(figsize=(10, 8))
    
gs = gridspec.GridSpec(1, 2, width_ratios=[5, 1])
ax1 = plt.subplot(gs[0])
for i, (x_pre_sce, y_pre_sce, std_x_pre_sce, std_y_pre_sce, name_pre_sce) in enumerate(
            zip(a_star_pre_sce, b_star_pre_sce, std_a_pre_sce, std_b_pre_sce, names_pre_sce)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_pre_sce, y_pre_sce, xerr=std_x_pre_sce, yerr=std_y_pre_sce, fmt=marker, markersize=10, color=color, label=name_pre_sce)

for i, (x_post_sce, y_post_sce, std_x_post_sce, std_y_post_sce, name_post_sce) in enumerate(
            zip(a_star_post_sce, b_star_post_sce, std_a_post_sce, std_b_post_sce, names_post_sce)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post_sce, y_post_sce, xerr=std_x_post_sce, yerr=std_y_post_sce, fmt=marker, markersize=10, markerfacecolor='white',     color=color, label=name_post_sce)

#added SCI with trasparency alpha=0.5        
for i, (x_pre_sci, y_pre_sci, std_x_pre_sci, std_y_pre_sci, name_pre_sci) in enumerate(
            zip(a_star_pre_sci, b_star_pre_sci, std_a_pre_sci, std_b_pre_sci, names_pre_sci)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  # 
        ax1.errorbar(x_pre_sci, y_pre_sci, xerr=std_x_pre_sci, yerr=std_y_pre_sci, fmt=marker, markersize=10, color=color, label=name_pre_sci, alpha=0.5)

for i, (x_post_sci, y_post_sci, std_x_post_sci, std_y_post_sci, name_post_sci) in enumerate(
            zip(a_star_post_sci, b_star_post_sci, std_a_post_sci, std_b_post_sci, names_post_sci)):
        marker = markers[i % len(markers)]  
        color = colors[i % len(colors)]  
        ax1.errorbar(x_post_sci, y_post_sci, xerr=std_x_post_sci, yerr=std_y_post_sci, fmt=marker, markersize=10, markerfacecolor='white', color=color, label=name_post_sci, alpha=0.5)


for x_pre_sce, y_pre_sce, x_post_sce, y_post_sce, color in zip(a_star_pre_sce, b_star_pre_sce, a_star_post_sce, b_star_post_sce, colors):
        ax1.plot([x_pre_sce, x_post_sce], [y_pre_sce, y_post_sce], linestyle='-', color=color, linewidth=0.8, alpha=0.5)
for x_pre_sci, y_pre_sci, x_post_sci, y_post_sci, color in zip(a_star_pre_sci, b_star_pre_sci, a_star_post_sci, b_star_post_sci, colors):
        ax1.plot([x_pre_sci, x_post_sci], [y_pre_sci, y_post_sci], linestyle='-', color=color, linewidth=0.8, alpha=0.3)

ax1.set_xlabel('a* [CIELAB]', fontsize=15)
ax1.set_ylabel('b* [CIELAB]', fontsize=15)
ax1.set_title('SCE and SCI(transparent)')

handles, labels = ax1.get_legend_handles_labels()
ax1.legend(loc='upper center', bbox_to_anchor=(0.5, 1.3), ncol=4, fontsize=9,  framealpha=1)

ax2 = plt.subplot(gs[1])

for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
        zip(zeros, L_star_pre_sce, std_a_pre_sce, std_L_pre_sce, names_pre_sce)):
    marker = markers[i % len(markers)] 
    color = colors[i % len(colors)]  
    ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
        zip(ones, L_star_post_sce, std_a_post_sce, std_L_post_sce, names_post_sce)):
    marker = markers[i % len(markers)] 
    color = colors[i % len(colors)] 
    ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)
for i, (x_pre, y_pre, std_x_pre, std_y_pre, name_pre) in enumerate(
        zip(zeros, L_star_pre_sci, std_a_pre_sci, std_L_pre_sci, names_pre_sci)):
    marker = markers[i % len(markers)]  
    color = colors[i % len(colors)]  
    ax2.errorbar(x_pre, y_pre, yerr=std_y_pre, fmt=marker, capsize=5, markersize=10, color=color, label=name_pre)

for i, (x_post, y_post, std_x_post, std_y_post, name_post) in enumerate(
        zip(-ones, L_star_post_sci, std_a_post_sci, std_L_post_sci, names_post_sci)):
    marker = markers[i % len(markers)]  
    color = colors[i % len(colors)]  
    ax2.errorbar(x_post, y_post, yerr=std_y_post, fmt=marker, capsize=5, markersize=10, markerfacecolor='white',
                        color=color, label=name_post)


for x_pre, y_pre, x_post, y_post, color in zip(zeros, L_star_pre_sce, ones, L_star_post_sce, colors):
    ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)
for x_pre, y_pre, x_post, y_post, color in zip(zeros, L_star_pre_sci, -ones, L_star_post_sci, colors):
    ax2.plot([x_pre, x_post], [y_pre, y_post], linestyle='-', color=color, linewidth=0.8, alpha=0.5)
ax2.set_title('SCI(neg) and SCE(pos)')

x_min, x_max = -1.3, 1.3  
ax2.set_xlim(x_min, x_max)

ax2.set_ylabel('L* [CIELAB]', fontsize=15)

plt.grid(True)
plt.tight_layout()  
plt.savefig(folder_path+ '/' + 'SCI_SCE_'+filename[:-5]+'.png', dpi=300, bbox_inches='tight')
plt.show()


# In[ ]:




