"""Optimistic per-workflow budget feasibility diagnostic using real start-time telemetry."""
from pathlib import Path
from collections import defaultdict
import sys,math
import pandas as pd
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'src'))
import simulate
tr=simulate.Trace(simulate.DATA)
rows=[]
for start in simulate.WINDOWS:
    for rate in (4.0,8.0):
        for seed in simulate.SEEDS:
            tasks,wfs,roots=simulate.gen_workloads(start,168,rate,seed)
            bywf=defaultdict(list)
            for task in tasks:
                if task['kind']=='Q':bywf[task['wid']].append(task)
            for w in wfs:
                costs=[]
                for task in bywf[w['wid']]:
                    device=[]
                    for b in simulate.BACKENDS:
                        c=tr.state(b,w['arrival'])
                        if c is not None:device.append(simulate.loss(task,{'cz':c['cz'],'ro':c['ro']}))
                    if not device:break
                    costs.append(min(device))
                if len(costs)!=3:continue
                best=sum(costs)
                rows.append(dict(window=start,rate_per_hour=rate,seed=seed,workflow_id=w['wid'],budget=w['budget'],
                                 optimistic_min_cost=best,optimistic_failure=1-math.exp(-best),
                                 infeasible=best>w['budget']))
d=pd.DataFrame(rows)
d.to_csv(ROOT/'results'/'budget_feasibility_rows.csv',index=False)
out=d.groupby('rate_per_hour').agg(n=('workflow_id','size'),infeasible_rate=('infeasible','mean'),
    median_min_cost=('optimistic_min_cost','median'),median_budget=('budget','median'),
    mean_min_failure=('optimistic_failure','mean')).reset_index()
out.to_csv(ROOT/'results'/'budget_feasibility.csv',index=False)
print(out.to_string(index=False))
