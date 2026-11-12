"""Plot real calibration snapshot medians used by the scheduler experiment."""
from pathlib import Path
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
ROOT=Path(__file__).resolve().parents[1]
x=pd.read_csv(ROOT/'data'/'ibm_heron_calibration_snapshots.csv')
x['observed_time']=pd.to_datetime(x['observed_time'],format='mixed',utc=True)
colors={'ibm_fez':'#1769aa','ibm_kingston':'#d1495b','ibm_marrakesh':'#2a9d8f'}
fig,axs=plt.subplots(2,1,figsize=(6.5,3.8),sharex=True)
for ax,fam,title in zip(axs,['cz_error','readout_error'],['CZ gate error','Readout error']):
    g=x[x.property_family==fam]
    for be in colors:
        z=g[g.backend==be].sort_values('observed_time')
        ax.plot(z.observed_time,z.median_error,lw=.65,alpha=.85,color=colors[be],label=be.replace('ibm_',''))
    ax.set_yscale('log');ax.set_ylabel('Median error (log)');ax.set_title(title,fontsize=9,loc='left');ax.grid(True,which='both',ls=':',lw=.45,alpha=.6)
axs[-1].set_xlabel('Observation time (UTC)')
axs[-1].xaxis.set_major_locator(mdates.MonthLocator(interval=2))
axs[-1].xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
axs[0].legend(ncol=3,frameon=False,fontsize=8,loc='upper right')
fig.tight_layout(pad=.7)
out=ROOT/'figures'/'calibration_medians.png'
fig.savefig(out,dpi=600,bbox_inches='tight')
fig.savefig(ROOT/'figures'/'calibration_medians.pdf',bbox_inches='tight')
print(out)
