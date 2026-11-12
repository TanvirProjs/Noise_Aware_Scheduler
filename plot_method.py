"""Draw the EB-UCB implementation pipeline for the manuscript."""
from pathlib import Path
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

ROOT=Path(__file__).resolve().parents[1]
fig,ax=plt.subplots(figsize=(6.55,1.72))
ax.set_xlim(0,1);ax.set_ylim(0,1);ax.axis('off')
boxes=[
    (0.025,0.27,0.29,0.56,'Calibration stream','Observed CZ and readout\nerror summaries'),
    (0.355,0.27,0.29,0.56,'Causal risk estimate','Latest value + seven-day\nmedian–MAD and freshness'),
    (0.685,0.27,0.29,0.56,'DAG dispatch + ledger','Budget and deadline priority;\nreserve then replace on finish'),
]
for x,y,w,h,title,body in boxes:
    patch=FancyBboxPatch((x,y),w,h,boxstyle='round,pad=0.012,rounding_size=0.02',
                         linewidth=1.0,edgecolor='#34495e',facecolor='#eef3f8')
    ax.add_patch(patch)
    ax.text(x+w/2,y+h*.72,title,ha='center',va='center',fontsize=8,fontweight='bold',color='#18324b')
    ax.text(x+w/2,y+h*.37,body,ha='center',va='center',fontsize=7.4,color='#202830',linespacing=1.18)
for x0,x1 in [(0.315,0.355),(0.645,0.685)]:
    ax.add_patch(FancyArrowPatch((x0+0.006,0.53),(x1-0.006,0.53),arrowstyle='-|>',mutation_scale=10,
                                 linewidth=1.0,color='#34495e'))
ax.text(.50,.10,'Each completion updates realized workflow cost before the next dispatch',
        ha='center',va='center',fontsize=7,color='#52796f')
fig.savefig(ROOT/'figures'/'ebucb_pipeline.png',dpi=600,bbox_inches='tight',facecolor='white')
fig.savefig(ROOT/'figures'/'ebucb_pipeline.pdf',bbox_inches='tight',facecolor='white')
