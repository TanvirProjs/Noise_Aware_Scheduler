"""Causal next-observation coverage of the rolling median-MAD estimate."""
from pathlib import Path
import numpy as np,pandas as pd
ROOT=Path(__file__).resolve().parents[1]
x=pd.read_csv(ROOT/'data'/'ibm_heron_calibration_snapshots.csv')
x['observed_time']=pd.to_datetime(x['observed_time'],format='mixed',utc=True)
rows=[]
for (be,fam),g in x.groupby(['backend','property_family']):
    g=g.sort_values('observed_time').reset_index(drop=True)
    y=g.median_error.to_numpy(float); t=g.observed_time.astype('int64').to_numpy()/1e9
    source_age=g.p90_age_seconds.fillna(0).to_numpy(float)
    for i in range(1,len(g)):
        hist=np.flatnonzero((t[:i]>=t[i]-7*86400)&np.isfinite(y[:i]))
        if len(hist)<3: continue
        vals=y[hist]; last=y[i-1]
        med=np.median(vals);mad=np.median(np.abs(vals-med))
        # At the next decision time, only the preceding observation's reported
        # calibration age plus elapsed time is causally available.
        age=source_age[i-1]+(t[i]-t[i-1])
        pred=max(last,med+1.645*1.4826*mad)
        pred*=1+min(.10,max(0,age-12*3600)/(60*3600)*.1)
        pred=min(.95,pred)
        rows.append(dict(backend=be,property_family=fam,observed_time=g.observed_time.iloc[i],actual=y[i],point_forecast=last,upper_forecast=pred,absolute_error=abs(y[i]-last),covered=y[i]<=pred,upper_margin=pred-y[i],age_hours=age/3600))
d=pd.DataFrame(rows)
d.to_csv(ROOT/'results'/'forecast_validation_rows.csv',index=False)
out=d.groupby('property_family').agg(n=('actual','size'),coverage=('covered','mean'),point_mae=('absolute_error','mean'),mean_upper_margin=('upper_margin','mean'),median_staleness_h=('age_hours','median')).reset_index()
out.to_csv(ROOT/'results'/'forecast_validation.csv',index=False)
print(out.to_string(index=False))
