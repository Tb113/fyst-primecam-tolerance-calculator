#!/usr/bin/env python
# coding: utf-8

# In[46]:


import numpy as np
import matplotlib.pyplot as plt
from spire.xls import *


# In[47]:


# Load Excel File
wb = Workbook()
wb.LoadFromFile("C:/Users/tanne/Downloads/Parametric_Model_3Lens.xlsx")


# In[48]:


# Define relevant spreadsheets
main=wb.Worksheets[0]
tol=wb.Worksheets[1]

# Calcuate values, in case they haven't been already
wb.CalculateAllValue()

# Define rows in each spreadsheet where data is located
mfilt = (np.where(~np.isnan([x.NumberValue 
                             for x in main.AllocatedRange['J1:J100']]))[0]+1).astype(int)
tfilt = (np.where(~np.isnan([x.NumberValue 
                             for x in tol.AllocatedRange['N1:N10']]))[0]+1).astype(int)


# In[49]:


tfilt


# In[52]:


# Initialize arrays
tol_arr = np.zeros((len(tfilt), 2))
dev_mat = np.zeros((len(mfilt), len(tfilt)))

# Generate array of tolerances from spreadsheet
for i in range(len(tfilt)):
    tol_arr[i][0] = tol.Range[f'J{tfilt[i]}'].NumberValue
    tol_arr[i][1] = tol.Range[f'L{tfilt[i]}'].NumberValue

# From the Excel, generate a matrix that transforms component deltas into lens deltas
# Generating this matrix at the start saves us from having to interface with the excel over every iteration of the MC method
# This drastically reduces computation time
i=0

for m in mfilt: # For each column in the eventual matrix    
    main.Range[f'J{m}'].NumberValue = main.Range[f'K{m}'].NumberValue # Initialize row in case changes have been made
    main.Range[f'J{m}'].NumberValue += 1 # Equivalent to adding an e_j vector to an array to isolate column j
    wb.CalculateAllValue() # Calculate 
    for n in range(len(tfilt)): # For each row
        dev_mat[i][n] = tol.Range[f'K{tfilt[n]}'].FormulaValue # save the column of deltas in the lenses tab to the matrix
    i+=1
    main.Range[f'J{m}'].NumberValue = main.Range[f'K{m}'].NumberValue # Reset row
    
dev_mat = dev_mat.T # Transpose array so we can use it as a matrix operator


# In[54]:


# Function that conducts the monte carlo analysis
def run_carlo(dev, runs, main, tol, rfun):
    # Initialize arrays and dictionaries (including s index)
    frac = np.zeros((len(runs), len(dev)))
    
    res_dict={}
    check_dict={}
    g_dict={}
    
    s = 0
    
    # Iterate over number of runs
    for b in runs:
        
        # Init
        t = 0
        res_dict[b]={}
        check_dict[b]={}
        g_dict[b]={}
        
        # Iterate over size of deviations
        for x in dev:
            
            # Calculate and record each critical dimension delta in the res array
            res=np.matmul(main, rfun(x, size=(main.shape[1], b)))
            
            # Check if each dimension passes the tolerance window for each run
            check=np.array([[(a > tol_arr[c][0]) and (a < tol_arr[c][1]) for a in res[c]] for c in range(len(tfilt))])
            # Check if all dimensions pass their respective tolerance windows for each run
            g = np.prod(check, 0).astype(bool)
            # record what fraction of the runs pass
            frac[s][t] = sum(g)/b
            
            # Record other results in their respective dictionaries
            res_dict[b][round(x, 3)]=res
            check_dict[b][round(x, 3)]=check
            g_dict[b][round(x, 3)]=g
            
            # Increment
            t+=1
        s+=1
        
    return frac, res_dict, check_dict, g_dict


# In[61]:


# TEST 1: Determining convergence for MC Method
dev_T1=np.array([0.25, 0.5, 0.75])
run = 10000

# Run MC method for three deviations and three different RNG functions
fr, rd, cd, gd = run_carlo(dev_T1, np.array([run]), dev_mat, tol_arr, lambda f, **kwargs: np.random.uniform(-f, f, **kwargs))
fr_gauss, rd_gauss, cd_gauss, gd_gauss = run_carlo(dev_T1, np.array([run]), dev_mat, tol_arr, lambda f, **kwargs: np.random.normal(0, f, **kwargs))
fr_binary, rd_binary, cd_binary, gd_binary = run_carlo(dev_T1, np.array([run]), dev_mat, tol_arr, lambda f, **kwargs: np.random.choice((-f, f), **kwargs))


# In[62]:


# For each run of the nine MC method results, calculate the passing fraction for it and all trials prior (this takes some time)
sumarr=np.array([[sum(gd[run][m][0:n])/n for n in range(1, run)] for m in dev_T1])
sumarr_gauss=np.array([[sum(gd_gauss[run][m][0:n])/n for n in range(1, run)] for m in dev_T1])
sumarr_binary=np.array([[sum(gd_binary[run][m][0:n])/n for n in range(1, run)] for m in dev_T1])


# In[63]:


# Plot the convergence of these passing fractions

fig, ax = plt.subplots(1, 3, figsize=(12, 4), clear = True, sharey=True)
dev = np.array([0.25, 0.5, 0.75])
c_arr = ['blue', 'orange', 'green']

# Plot the nine trials
for r in [0, 1, 2]:
    # Opaque (fraction not 0 or 1)
    ind=[np.intersect1d(np.where(sumarr[r]>0), np.where(sumarr[r]<1)), 
        np.intersect1d(np.where(sumarr_gauss[r]>0), np.where(sumarr_gauss[r]<1)),
        np.intersect1d(np.where(sumarr_binary[r]>0), np.where(sumarr_binary[r]<1))]
    ax[0].plot(np.arange(1, run)[ind[0]], 1-sumarr[r][ind[0]], color = f'tab:{c_arr[r]}', label = '$\delta_{max} = \pm$'+f'{.1+.05*r:.2f}')
    ax[1].plot(np.arange(1, run)[ind[1]], 1-sumarr_gauss[r][ind[1]], color = f'tab:{c_arr[r]}', label = '$\sigma$ = '+f'{.1+.05*r:.2f}')
    ax[2].plot(np.arange(1, run)[ind[2]], 1-sumarr_binary[r][ind[2]], color = f'tab:{c_arr[r]}', label = '$\delta = \pm$'+f'{.1+.05*r:.2f}')
    
    # Transluscent (otherwise)
    ax[0].plot(np.arange(1, run), 1-sumarr[r], color = f'tab:{c_arr[r]}', label = '__', alpha = 0.25, zorder = -1)
    ax[1].plot(np.arange(1, run), 1-sumarr_gauss[r], color = f'tab:{c_arr[r]}', label = '__', alpha = 0.25, zorder = -1)
    ax[2].plot(np.arange(1, run), 1-sumarr_binary[r], color = f'tab:{c_arr[r]}', label = '__', alpha = 0.25, zorder = -1)

# Visual formatting
for a in ax:
    a.set_xscale('log')
    a.set_xlim((1, run))
    a.set_ylim((-.05, 1.05))
    a.legend(loc = 'upper right')

ax[0].set_title('Uniform')
ax[1].set_title('Gaussian')
ax[2].set_title('Binary')
ax[0].set_ylabel('Probablility of Tolerance Violation Estimate')
ax[1].set_xlabel('Number of Trials')
fig.suptitle('Testing the Monte Carlo Method', size = 15)


# In[55]:


# TEST 2: Covariance Plot
dev_T2=np.linspace(0, 1, 51)

# Generate a series of trials over a lineraly spaced range
fr4, rd4, cd4, gd4 = run_carlo(dev_T2, np.array([100000]), dev_mat, tol_arr, lambda f, **kwargs: np.random.uniform(-f, f, **kwargs))
fr4g, rd4g, cd4g, gd4g = run_carlo(dev_T2, np.array([100000]), dev_mat, tol_arr, lambda f, **kwargs: np.random.normal(0, f, **kwargs))
fr4b, rd4b, cd4b, gd4b = run_carlo(dev_T2, np.array([100000]), dev_mat, tol_arr, lambda f, **kwargs: np.random.choice((-f, f), **kwargs))


# In[56]:


# Box-plotting function
def plotbox(fig, ax, xmin, xmax, ymin, ymax, **kwargs):
    ax.plot((xmin, xmax), (ymin, ymin), **kwargs)
    ax.plot((xmin, xmax), (ymax, ymax), **kwargs)
    ax.plot((xmin, xmin), (ymin, ymax), **kwargs)
    ax.plot((xmax, xmax), (ymin, ymax), **kwargs)
    
# Label Comprehension formula
lab= lambda row: f"{tol.Range[f'B{row}'].Text} to {tol.Range[f'D{row}'].Text}"


# In[57]:


# plot one set of covariances
def plot_cov(fig, rd, gd, a, b, r, d, ax, tol_arr, xsc, ysc):

    # Scatter plot of the passing and failing runs, by critical dimension deviation
    ax.scatter(rd[r][d][b][gd[r][d]], rd[r][d][a][gd[r][d]], c = 'tab:blue', s = .1, zorder = 1, label = 'Pass')
    ax.scatter(rd[r][d][b][~gd[r][d]], rd[r][d][a][~gd[r][d]], c = 'tab:red', s = .1, zorder = 2, label = 'Fail')
    
    #ax.plot((-5, 5), (5, 5), c='k', ls = '-', lw=2, alpha = 1, label = 'Inv. Increment Tolerances')
    #ax.plot((-5, 5), (5, 5), c='k', ls = '--', lw=2, alpha = 1, label = 'Monte Carlo Tolerances')

    # Plot a box around the two tolerance ranges
    # Inverse Increment
    plotbox(fig, ax, tol_arr[b][0], tol_arr[b][1], tol_arr[a][0], tol_arr[a][1], c='k', ls = '-', zorder = 3, lw=1, alpha = 1, label = '__')
    # Monte Carlo
    #plotbox(fig, ax, -tol_arr[b][2], tol_arr[b][2], -tol_arr[a][2], tol_arr[a][2], c='k', ls = '--', zorder = 3, lw=1, alpha = 1, label = '__')
    
    # Plot the intersection between the two
    #ax.fill_between((max(tol_arr[b][0], -tol_arr[b][2]), min(tol_arr[b][1], tol_arr[b][2])), 
                    #(max(tol_arr[a][0], -tol_arr[a][2]), max(tol_arr[a][0], -tol_arr[a][2])),
                   #(min(tol_arr[a][1], tol_arr[a][2]), min(tol_arr[a][1], tol_arr[a][2])), 
                    #zorder = -2, alpha = 0.5, label = 'Intersection of Tolerances')
    
    #plotbox(fig, ax, max(tol_arr[a][0], -tol_arr[a][2]), min(tol_arr[a][1], tol_arr[a][2]), 
            #max(tol_arr[b][0], -tol_arr[b][2]), min(tol_arr[b][1], tol_arr[a][2]), c='k', ls = '-', zorder = 0,lw=2)

    # Visual formatting
    ax.grid()
    ax.set_xlim((-xsc, xsc))
    ax.set_ylim((-xsc, xsc))


# In[58]:


# Function to concatenate labels for legend
lab = lambda row: f"{tol.Range[f'B{row}'].Text} to {tol.Range[f'D{row}'].Text} (mm)"

# Plot the whole covariance plot
def form_cov(rd, gd, n, r, d, xsc, ysc, save = False):
    fig, ax = plt.subplots(n, n, figsize=(12, 12), clear = True)
    
    for i in np.arange(0, n):
        # All the covariances
        for j in np.arange(i+1, n):
            plot_cov(fig, rd, gd, i, j, r, d, ax[n-1-i][n-1-j], tol_arr, xsc, ysc)
        
        for j in np.arange(0, n):
            ax[i][j].tick_params(labelsize = 15)
        
        # Histograms
        # Pass
        ax[n-1-i][n-1-i].hist(rd[r][d][i], bins = np.arange(-xsc, xsc, 0.01), color='tab:blue', stacked = True, zorder = 0)
        # Fail
        ax[n-1-i][n-1-i].hist(rd[r][d][i][~gd[r][d]], bins = np.arange(-xsc, xsc, 0.01), color='tab:red', stacked = True, zorder = 1)
       
        # Formatting
        ax[i][i].set_xlim((-xsc, xsc))
        ax[i][i].set_ylim((0, ysc))
        ax[n-1][i].set_xlabel(lab(n-1-i+3), fontsize = 25)
        ax[i][0].set_ylabel(lab(n-1-i+3), fontsize = 25)
    
    # Legend setup (in an empty plot)
    ax[0][1].set_xlim(-1, 1)
    ax[0][1].set_ylim(-1, 1)
    ax[0][1].scatter((5), (5), c = 'tab:blue', s = 50, zorder = 3, label = 'Pass')
    ax[0][1].scatter((5), (5), c = 'tab:red', s = 50, zorder = 3, label = 'Fail')
    ax[0][1].plot((-5, 5), (5, 5), c='k', ls = '-', lw=2, alpha = 1, label = 'Lens Tolerances')
    #ax[0][1].plot((-5, 5), (5, 5), c='k', ls = '--', lw=2, alpha = 1, label = 'Monte Carlo Tolerances')
    #ax[0][1].fill_between((4, 5), (5, 5), (4, 4), zorder = -2, alpha = 0.5, label = 'Intersection of Tolerances')
    
    
    ax[0][1].legend(loc = 'lower left', fontsize = 25)
    
    if save == True: # option to save plot
        fig.savefig("[DESIRED FIGURE FILEPATH HERE]")
form_cov(rd4, gd4, 2, 100000, 0.5, 3.5, 800)


# In[59]:


# TEST 3: Pass Fraction as a function of deviation

# Reuses results from Test 2
fig, ax = plt.subplots(1, 2, figsize=(12, 6), clear = True)

# Setting logarithmic Scale for 2nd plot
ax[1].set_yscale('log')
ax[1].set_ylim((.8*10**(-4), 2))

for a in ax:
    # Plot each set of fractions, differentiated by RNG function (across both plots)
    a.plot(dev_T2, 1-fr4[0], color = 'tab:blue', zorder = 2, label = 'Uniform')
    a.plot(dev_T2, 1-fr4g[0], color = 'tab:orange', zorder = 1, label = 'Gaussian')
    a.plot(dev_T2, 1-fr4b[0], color = 'tab:green', zorder = 0, label = 'Worst-Case')
    
    # Formatting
    a.set_xlabel('Characteristic Deviation (mm)', fontsize = 15)
    a.set_ylabel('Probability of Tolerance Violation', fontsize = 15)
    a.tick_params(labelsize = 10)
    #a.set_xlim(-0.025, 0.625)
    a.grid()

ax[0].legend()
ax[0].set_title('Linear', fontsize = 15)
ax[1].set_title('Logarithmic', fontsize = 15)


# In[60]:


# TEST 4: Test 3, but for a smaller window

# Generate the data for the new window, but just with the uniform distribution
fr5, rd5, cd5, gd5 = run_carlo(np.linspace(.05, .2, 51), np.array([1000000]), dev_mat, tol_arr, lambda f, **kwargs: np.random.uniform(-f, f, **kwargs))

fig, ax = plt.subplots(1, 2, figsize=(12, 6), clear = True)

#Repeat everything in Test 3
ax[1].set_yscale('log')
#ax[1].set_ylim((.8*10**(-4), 2))

for a in ax:
    a.plot(np.linspace(.05, .2, 51), 1-fr5[0], marker = 'o', color = 'tab:blue', zorder = 2)
    #a.plot(np.linspace(.05, .1, 21), 1-fr5g[0], marker = 's', color = 'tab:orange', zorder = 1)
    #a.plot(np.linspace(.05, .1, 21), 1-fr5b[0], marker = '^', color = 'tab:green', zorder = 0)
    
    a.set_xlabel('Max Deviation (mm)')
    a.set_ylabel('Probability of Tolerance Violation')
    #a.set_xlim(-0.025, 0.625)
    a.grid()
    
ax[0].set_title('Linear')
ax[1].set_title('Logarithmic')


# In[ ]:




