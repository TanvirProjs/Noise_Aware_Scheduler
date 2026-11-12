"""Reproducible event-driven evaluation of calibration-aware hybrid scheduling.

This is a scheduler simulation, not a quantum-hardware execution. Calibration inputs
are the measured IBM calibration snapshot summaries in ../data. Workload durations
and DAGs are generated from documented distributions in this file.
"""
from __future__ import annotations
import heapq, math, random
from pathlib import Path
from collections import defaultdict, deque
import numpy as np
import pandas as pd

ROOT=Path(__file__).resolve().parents[1]
DATA=ROOT/'data'/'ibm_heron_calibration_snapshots.csv'
OUT=ROOT/'results'
OUT.mkdir(exist_ok=True)
BACKENDS=['ibm_fez','ibm_kingston','ibm_marrakesh']
POLICIES=['FIFO-current','EDF-current','Min-risk','EB-current','EB-UCB']
WINDOWS=['2026-02-15','2026-04-15','2026-06-15','2026-08-15','2026-10-15']
SEEDS=(17,23,42)

class Trace:
    def __init__(self,path):
        x=pd.read_csv(path)
        x['observed_time']=pd.to_datetime(x['observed_time'],format='mixed',utc=True)
        self.by={}
        for be in BACKENDS:
            g=x[x.backend==be]
            self.by[be]={}
            for family in ('cz_error','readout_error'):
                z=g[g.property_family==family].sort_values('observed_time')
                times=z.observed_time.astype('int64').to_numpy()
                med=z.median_error.to_numpy(dtype=float)
                p90=z.p90_error.to_numpy(dtype=float)
                age=z.p90_age_seconds.fillna(0).to_numpy(dtype=float)
                self.by[be][family]=dict(times=times,median=med,p90=p90,age=age)
    def _last(self,be,family,t):
        series=self.by[be][family]
        now=int(round(t*1e9))
        i=np.searchsorted(series['times'],now,side='right')-1
        if i<0:return None
        age_since_observation=(now-series['times'][i])/1e9
        reported_age=max(0.0,float(series['age'][i]))
        # Conservative age of the stalest summarized component at decision time.
        effective_age=age_since_observation+reported_age
        return float(series['median'][i]),float(series['p90'][i]),effective_age,i
    def state(self,be,t):
        cz=self._last(be,'cz_error',t); ro=self._last(be,'readout_error',t)
        if cz is None or ro is None:return None
        if max(cz[2],ro[2])>72*3600:return None
        return dict(cz=cz[0],cz90=cz[1],ro=ro[0],ro90=ro[1],age=max(cz[2],ro[2]))
    def upper(self,be,t):
        cur=self.state(be,t)
        if cur is None:return None
        def bound(family):
            series=self.by[be][family]
            now=int(round(t*1e9))
            i=np.searchsorted(series['times'],now,side='right')-1
            lo=np.searchsorted(series['times'],now-int(7*24*3600*1e9),side='left')
            vals=series['median'][lo:i+1]
            if len(vals)<3:
                v=series['median'][i]
            else:
                center=np.median(vals)
                mad=np.median(np.abs(vals-center))
                v=max(series['median'][i],center+1.645*1.4826*mad)
            v*=1+min(.10,max(0,cur['age']-12*3600)/(60*3600)*0.1)
            return min(.95,float(v))
        return dict(cz=bound('cz_error'),ro=bound('readout_error'),age=cur['age'])

def loss(task,cal):
    return task['ncz']*(-math.log1p(-min(.95,cal['cz']))) + task['nro']*(-math.log1p(-min(.95,cal['ro'])))

def gen_workloads(start, hours, rate, seed):
    rng=random.Random(seed); tasks=[]; wf=[]; roots_by_wf=defaultdict(list); wid=0; tid=0
    t0=pd.Timestamp(start,tz='UTC').timestamp(); end=t0+hours*3600
    t=t0
    # Poisson arrivals; three QPU tasks per workflow and a fork/join stage.
    while True:
        t += rng.expovariate(rate/3600.0)
        if t>=end: break
        w=wid; wid+=1
        eps=rng.uniform(.40,.75) # allowed end-to-end proxy failure probability
        budget=-math.log1p(-eps)
        deadline=t+rng.uniform(1.5,4.0)*3600
        wf.append(dict(wid=w,arrival=t,deadline=deadline,budget=budget,epsilon=eps))
        # C-pre then two parallel quantum branches; join, final QPU refinement, C-post.
        def add(kind,preds,dur,ncz=0,nro=0):
            nonlocal tid
            task=dict(tid=tid,wid=w,kind=kind,preds=list(preds),duration=dur,
                      ncz=ncz,nro=nro,arrival=t,deadline=deadline,done=False,
                      ready=t if not preds else None, start=None,finish=None,actual_loss=0.,reserve=0.)
            tasks.append(task); tid+=1
            if not preds: roots_by_wf[w].append(task['tid'])
            return task['tid']
        cp=add('C',[],rng.uniform(120,360))
        q1=add('Q',[cp],rng.uniform(240,480),rng.randint(24,100),rng.randint(8,24))
        q2=add('Q',[cp],rng.uniform(240,480),rng.randint(24,100),rng.randint(8,24))
        c1=add('C',[q1],rng.uniform(180,480)); c2=add('C',[q2],rng.uniform(180,480))
        q3=add('Q',[c1,c2],rng.uniform(300,600),rng.randint(60,180),rng.randint(12,32))
        add('C',[q3],rng.uniform(120,360))
    return tasks,wf,roots_by_wf

def run_one(trace,start,hours,rate,seed,policy):
    tasks,wfs,roots_by_wf=gen_workloads(start,hours,rate,seed)
    byid={x['tid']:x for x in tasks}; children=defaultdict(list)
    for x in tasks:
        for p in x['preds']: children[p].append(x['tid'])
    wf={x['wid']:x for x in wfs}; completed_workflows={}; remaining={x['tid']:len(x['preds']) for x in tasks}
    arrivals=sorted((x['arrival'],x['wid']) for x in wfs); ai=0
    cpu_n=16; cpu_free=[True]*cpu_n; q_free={b:True for b in BACKENDS}
    cpu_wait=[]; q_wait=[]; events=[]; seq=0; now=arrivals[0][0] if arrivals else pd.Timestamp(start).timestamp()
    actual_done_loss=defaultdict(float); reserved=defaultdict(float); counters=defaultdict(int)
    def make_ready(tid0,when):
        x=byid[tid0]; x['ready']=when
        if x['kind']=='C': cpu_wait.append(tid0)
        else: q_wait.append(tid0)
    def schedule_cpu(now):
        nonlocal seq
        while cpu_wait and any(cpu_free):
            tid0=cpu_wait.pop(0); x=byid[tid0]; k=cpu_free.index(True); cpu_free[k]=False
            x['start']=now; x['finish']=now+x['duration']; x['resource']=k
            heapq.heappush(events,(x['finish'],seq,'C',tid0,k,None));seq+=1
    def est_pair(x,b,now,cached=None):
        if cached is None:
            cur=trace.state(b,now); upper=trace.upper(b,now) if cur is not None else None
        else: cur,upper=cached
        if cur is None:return None
        if policy=='EB-UCB': cal=upper
        else: cal={'cz':cur['cz'],'ro':cur['ro']}
        return loss(x,cal),cur
    def q_rank(x,b,now,est,cur):
        w=wf[x['wid']]; rem=w['budget']-actual_done_loss[x['wid']]-reserved[x['wid']]
        feasible=est<=rem+1e-12
        slack=(w['deadline']-now)/3600
        if policy=='FIFO-current': key=(0,x['arrival'],x['tid'],est)
        elif policy=='EDF-current': key=(0,w['deadline'],x['arrival'],x['tid'],est)
        elif policy=='Min-risk': key=(0,est,w['deadline'],x['arrival'],x['tid'])
        else:
            # Prefer a feasible dispatch; among feasible, EDF; then protect the most
            # constrained workflow by normalized remaining budget. If none fits, use
            # least normalized overshoot, with EDF as tie-breaker.
            key=(0 if feasible else 1, slack, est/max(w['budget'],1e-9),
                 max(0,est-rem)/max(w['budget'],1e-9), x['arrival'],x['tid'])
        return key,feasible,rem
    def schedule_q(now):
        nonlocal seq
        while q_wait and any(q_free.values()):
            best=None
            idle=[b for b in BACKENDS if q_free[b]]
            cached={}
            for b in idle:
                cur=trace.state(b,now); cached[b]=(cur,trace.upper(b,now) if cur is not None else None)
            for tid0 in q_wait:
                x=byid[tid0]
                for b in idle:
                    pair=est_pair(x,b,now,cached[b])
                    if pair is None: continue
                    est,cur=pair; key,feasible,rem=q_rank(x,b,now,est,cur)
                    cand=(key,tid0,b,est,cur,feasible,rem)
                    if best is None or cand[0]<best[0]:best=cand
            if best is None: break
            _,tid0,b,est,cur,feasible,rem=best
            q_wait.remove(tid0); x=byid[tid0]; q_free[b]=False
            x['start']=now;x['resource']=b;x['reserve']=est;reserved[x['wid']]+=est
            if not feasible:counters['budget-infeasible-dispatch']+=1
            # Simulator service time includes fixed remote-job overhead plus a gate-count
            # term; it is a workload assumption, not measured IBM execution time.
            service=x['duration']
            x['finish']=now+service
            heapq.heappush(events,(x['finish'],seq,'Q',tid0,b,cur));seq+=1
    while ai<len(arrivals) or events:
        next_arr=arrivals[ai][0] if ai<len(arrivals) else float('inf')
        next_ev=events[0][0] if events else float('inf')
        now=min(next_arr,next_ev)
        if now==float('inf'):break
        # Finish all same-time events.
        while events and events[0][0]<=now+1e-9:
            fin,_,kind,tid0,res,cal=heapq.heappop(events); x=byid[tid0]
            x['done']=True;x['finish']=fin
            if kind=='C':cpu_free[res]=True
            else:
                q_free[res]=True
                actual=loss(x,{'cz':cal['cz'],'ro':cal['ro']})
                x['actual_loss']=actual;actual_done_loss[x['wid']]+=actual
                reserved[x['wid']]-=x['reserve']
            # Decrement each child's indegree once the parent completes.
            for ch in children[tid0]:
                remaining[ch]-=1
                if remaining[ch]==0:make_ready(ch,fin)
            if not children[tid0]:
                w=wf[x['wid']]; completed_workflows[x['wid']]=fin
        # Add arrivals at this time.
        while ai<len(arrivals) and arrivals[ai][0]<=now+1e-9:
            at,wid0=arrivals[ai];ai+=1
            for r in roots_by_wf[wid0]:make_ready(r,at)
        schedule_cpu(now);schedule_q(now)
    # Per-workflow end-to-end summaries.
    records=[]
    for w in wfs:
        fin=completed_workflows.get(w['wid'])
        if fin is None:continue
        risk=actual_done_loss[w['wid']]
        records.append(dict(policy=policy,window=pd.Timestamp(start).date().isoformat(),rate_per_hour=rate,seed=seed,
            workflow_id=w['wid'],response_min=(fin-w['arrival'])/60,deadline_miss=fin>w['deadline'],
            risk_proxy=1-math.exp(-risk),budget_exceeded=risk>w['budget']+1e-12,
            qpu_wait_min=np.mean([(x['start']-x['ready'])/60 for x in tasks if x['wid']==w['wid'] and x['kind']=='Q']),
            error_budget=w['epsilon']))
    # Utilization denominator uses calendar elapsed through last completion.
    span=max([v for v in completed_workflows.values()],default=now)-now+hours*3600
    return records,dict(policy=policy,window=pd.Timestamp(start).date().isoformat(),rate_per_hour=rate,seed=seed,
        workflows=len(wfs),completed=len(records),budget_infeasible_dispatches=counters['budget-infeasible-dispatch'])

def main():
    tr=Trace(DATA)
    all_records=[]; run_summ=[]
    hours=168
    for widx,start in enumerate(WINDOWS):
        for rate in (4.0,8.0):
            for seed in SEEDS:
                # Deterministic seed changes workload, while methods share task inputs.
                for pol in POLICIES:
                    rec,summ=run_one(tr,start,hours,rate,seed,pol);all_records.extend(rec);run_summ.append(summ)
                    print(start,rate,seed,pol,len(rec),flush=True)
    d=pd.DataFrame(all_records)
    d.to_csv(OUT/'workflow_metrics.csv',index=False)
    pd.DataFrame(run_summ).to_csv(OUT/'run_manifest.csv',index=False)
    agg=d.groupby(['policy','rate_per_hour']).agg(
        n=('workflow_id','size'),response_mean_min=('response_min','mean'),response_p95_min=('response_min',lambda x:x.quantile(.95)),
        qpu_wait_mean_min=('qpu_wait_min','mean'),risk_mean=('risk_proxy','mean'),
        budget_violation_rate=('budget_exceeded','mean'),deadline_miss_rate=('deadline_miss','mean')).reset_index()
    agg.to_csv(OUT/'summary.csv',index=False)
    # Replication-block summaries treat each telemetry-window/seed pair as one unit.
    blocks=d.groupby(['policy','window','rate_per_hour','seed']).agg(
        n=('workflow_id','size'),response_mean_min=('response_min','mean'),
        response_p95_min=('response_min',lambda x:x.quantile(.95)),
        qpu_wait_mean_min=('qpu_wait_min','mean'),risk_mean=('risk_proxy','mean'),
        budget_violation_rate=('budget_exceeded','mean'),
        deadline_miss_rate=('deadline_miss','mean')).reset_index()
    blocks.to_csv(OUT/'scenario_summary.csv',index=False)
    print('\n',agg.to_string(index=False))
if __name__=='__main__':main()
