"""Generate the IEEE A4, two-column manuscript from recorded run outputs."""
from pathlib import Path
import pandas as pd, statistics
from docx import Document
from docx.shared import Inches, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parents[1]
DATA=pd.read_csv(ROOT/'data'/'ibm_heron_calibration_snapshots.csv')
SUMMARY=pd.read_csv(ROOT/'results'/'summary.csv')
FVALID=pd.read_csv(ROOT/'results'/'forecast_validation.csv')
FEAS=pd.read_csv(ROOT/'results'/'budget_feasibility.csv')
SCENARIO=pd.read_csv(ROOT/'results'/'scenario_summary.csv')
OUT=ROOT/'paper.docx'

def val(policy,rate,col):
    return float(SUMMARY[(SUMMARY.policy==policy)&(SUMMARY.rate_per_hour==rate)][col].iloc[0])
def block_stat(policy,rate,col):
    x=SCENARIO[(SCENARIO.policy==policy)&(SCENARIO.rate_per_hour==rate)][col]
    digits=3 if col=='risk_mean' else 1
    return f'{x.mean():.{digits}f} ({x.std(ddof=1):.{digits}f})'
def block_pct(policy,rate,col):
    x=SCENARIO[(SCENARIO.policy==policy)&(SCENARIO.rate_per_hour==rate)][col]*100
    return f'{x.mean():.1f} ({x.std(ddof=1):.1f})%'
def fmt(x,d=1): return f'{x:.{d}f}'
def pct(x,d=1): return f'{100*x:.{d}f}%'

def set_cell_shading(cell,fill='E8EEF5'):
    tcPr=cell._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),fill);tcPr.append(shd)
def set_cell_margins(cell,top=40,start=50,bottom=40,end=50):
    tc=cell._tc;tcPr=tc.get_or_add_tcPr();m=OxmlElement('w:tcMar')
    for tag,v in [('top',top),('start',start),('bottom',bottom),('end',end)]:
        el=OxmlElement('w:'+tag);el.set(qn('w:w'),str(v));el.set(qn('w:type'),'dxa');m.append(el)
    tcPr.append(m)
def set_repeat_table_header(row):
    trPr=row._tr.get_or_add_trPr();el=OxmlElement('w:tblHeader');el.set(qn('w:val'),'true');trPr.append(el)
def fmt_table(table, widths=None, font=7.1):
    table.alignment=WD_TABLE_ALIGNMENT.CENTER;table.autofit=False
    total=sum(widths) if widths else 3.2
    tblPr=table._tbl.tblPr;tblW=tblPr.find(qn('w:tblW'))
    if tblW is None:tblW=OxmlElement('w:tblW');tblPr.append(tblW)
    tblW.set(qn('w:w'),str(int(total*1440)));tblW.set(qn('w:type'),'dxa')
    ind=tblPr.find(qn('w:tblInd'))
    if ind is None:ind=OxmlElement('w:tblInd');tblPr.append(ind)
    ind.set(qn('w:w'),'0');ind.set(qn('w:type'),'dxa')
    for ri,row in enumerate(table.rows):
        if ri==0:set_repeat_table_header(row)
        trpr=row._tr.get_or_add_trPr();cs=OxmlElement('w:cantSplit');trpr.append(cs)
        for ci,cell in enumerate(row.cells):
            cell.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            set_cell_margins(cell)
            if widths and ci<len(widths): cell.width=Inches(widths[ci])
            for p in cell.paragraphs:
                p.alignment=WD_ALIGN_PARAGRAPH.CENTER
                if ri < len(table.rows)-1:p.paragraph_format.keep_with_next=True
                p.paragraph_format.space_after=Pt(0);p.paragraph_format.space_before=Pt(0);p.paragraph_format.line_spacing=1.0
                for r in p.runs:
                    r.font.name='Times New Roman';r.font.size=Pt(font)
                    if ri==0:r.bold=True
            if ri==0:set_cell_shading(cell)
    if widths:
        for row in table.rows:
            for i,w in enumerate(widths): row.cells[i].width=Inches(w)

def add_body(doc,text,after=2.1):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after=Pt(after);p.paragraph_format.space_before=Pt(0);p.paragraph_format.line_spacing=1.0
    p.paragraph_format.keep_together=False
    r=p.add_run(text);r.font.name='Times New Roman';r.font.size=Pt(10)
    return p
def add_head(doc,tag,title):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before=Pt(6);p.paragraph_format.space_after=Pt(2);p.paragraph_format.keep_with_next=True
    r=p.add_run((tag+' ' if tag else '')+title.upper());r.bold=True;r.font.name='Times New Roman';r.font.size=Pt(10)
    return p
def add_subhead(doc,letter,title):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.space_before=Pt(4);p.paragraph_format.space_after=Pt(1);p.paragraph_format.keep_with_next=True
    r=p.add_run((f'{letter}. ' if letter else '')+title);r.bold=True;r.italic=True;r.font.name='Times New Roman';r.font.size=Pt(9.2)
    return p
def add_eq(doc,text,num):
    t=doc.add_table(rows=1,cols=2);t.alignment=WD_TABLE_ALIGNMENT.CENTER;t.autofit=False
    tw=OxmlElement('w:tblW');tw.set(qn('w:w'),str(int(3.38*1440)));tw.set(qn('w:type'),'dxa');t._tbl.tblPr.append(tw)
    t.columns[0].width=Inches(3.0);t.columns[1].width=Inches(.38)
    c1,c2=t.rows[0].cells;c1.width=Inches(3.0);c2.width=Inches(.38)
    for c in (c1,c2):
        set_cell_margins(c,0,0,0,0)
        tcPr=c._tc.get_or_add_tcPr();b=OxmlElement('w:tcBorders')
        for side in ('top','left','bottom','right','insideH','insideV'):
            e=OxmlElement('w:'+side);e.set(qn('w:val'),'nil');b.append(e)
        tcPr.append(b)
    p=c1.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(0)
    r=p.add_run(text);r.font.name='Times New Roman';r.font.size=Pt(9.1)
    p=c2.paragraphs[0];p.alignment=WD_ALIGN_PARAGRAPH.RIGHT;p.paragraph_format.space_after=Pt(0)
    r=p.add_run(f'({num})');r.font.name='Times New Roman';r.font.size=Pt(9.1)

def add_table_caption(doc,text):
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(3);p.paragraph_format.space_after=Pt(1);p.paragraph_format.keep_with_next=True
    r=p.add_run(text.upper());r.font.name='Times New Roman';r.font.size=Pt(8);r.bold=True

def add_results_table(doc,rate):
    tabnum='V' if rate==4.0 else 'VI'
    add_table_caption(doc,f'TABLE {tabnum}. SIMULATION RESULTS AT {rate} WORKFLOWS/HOUR')
    headers=['Policy','Mean resp.','P95 resp.','QPU wait','Failure','Budget >','Deadline miss']
    rows=[]
    for pol,label in [('FIFO-current','FIFO'),('EDF-current','EDF'),('Min-risk','Min-risk'),('EB-current','EB-current'),('EB-UCB','EB-UCB')]:
        rows.append([label,block_stat(pol,rate,'response_mean_min'),block_stat(pol,rate,'response_p95_min'),block_stat(pol,rate,'qpu_wait_mean_min'),
                     block_stat(pol,rate,'risk_mean'),block_pct(pol,rate,'budget_violation_rate'),block_pct(pol,rate,'deadline_miss_rate')])
    table=doc.add_table(rows=1,cols=len(headers))
    for i,h in enumerate(headers):table.rows[0].cells[i].text=h
    for row in rows:
        cells=table.add_row().cells
        for i,v in enumerate(row):cells[i].text=v
    fmt_table(table,[.48,.49,.49,.43,.42,.50,.48],font=6.4)
    p=doc.add_paragraph();p.paragraph_format.space_after=Pt(2);p.paragraph_format.space_before=Pt(0)
    r=p.add_run('Cells are mean (SD) across 15 window/seed blocks. Response and QPU wait are minutes; P95 is response. Failure is an independent-error proxy; budget > is target exceedance.');r.italic=True;r.font.name='Times New Roman';r.font.size=Pt(7.0)

# A4 page and typography follow the supplied IEEE conference template.
doc=Document()
for section in doc.sections:
    section.page_width=Cm(21.0);section.page_height=Cm(29.7)
    section.left_margin=Inches(.62);section.right_margin=Inches(.62)
    section.top_margin=Inches(.375);section.bottom_margin=Inches(1.0)
# First section is the full-width title and author block.
first=doc.sections[0]
first.different_first_page_header_footer=True
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(7);p.paragraph_format.keep_with_next=True
r=p.add_run('Noise-Aware, Error-Budgeted Scheduling of Hybrid Quantum–Classical Workflows Using Time-Varying QPU Calibration Data');r.bold=True;r.font.name='Times New Roman';r.font.size=Pt(24)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(1);p.paragraph_format.keep_with_next=True
r=p.add_run('Dr. Tanvir H. Sardar');r.font.name='Times New Roman';r.font.size=Pt(10.5)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(0);p.paragraph_format.keep_with_next=True
r=p.add_run('Department of Computer Science & Engineering, School of Engineering');r.font.name='Times New Roman';r.font.size=Pt(9.3)
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_after=Pt(7)
r=p.add_run('Dayananda Sagar University, Bengaluru, India');r.font.name='Times New Roman';r.font.size=Pt(9.3)
# Start a continuous two-column IEEE body after the full-width author block.
sec=doc.add_section(WD_SECTION.CONTINUOUS)
sec.page_width=Cm(21.0);sec.page_height=Cm(29.7);sec.left_margin=Inches(.62);sec.right_margin=Inches(.62);sec.top_margin=Inches(.375);sec.bottom_margin=Inches(1.0)
cols=sec._sectPr.find(qn('w:cols'))
if cols is None: cols=OxmlElement('w:cols');sec._sectPr.append(cols)
cols.set(qn('w:num'),'2');cols.set(qn('w:space'),'360');cols.set(qn('w:equalWidth'),'true')
# Abstract and index terms.
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY;p.paragraph_format.space_after=Pt(3);p.paragraph_format.line_spacing=1.0
r=p.add_run('Abstract—');r.bold=True;r.font.name='Times New Roman';r.font.size=Pt(9)
r=p.add_run(f"Hybrid quantum–classical applications combine classical preparation and postprocessing with quantum processing unit (QPU) tasks whose execution quality changes with hardware calibration. This paper implements EB-UCB, a feasibility-first dispatch heuristic that schedules ready tasks over three IBM Heron processors using a workflow-level log-error allowance and a causal seven-day median–MAD calibration estimate. We replay 10,020 observed calibration-summary rows and {int(SUMMARY[SUMMARY.policy=='EB-UCB'].n.sum()):,} generated workflows across five telemetry windows, three workload seeds, and two arrival rates. No quantum jobs were submitted; the error metric is an independent-operation proxy, not measured circuit fidelity. At eight workflows/hour, EB-UCB's pooled mean response is {fmt(val('EB-UCB',8.0,'response_mean_min'))} min versus {fmt(val('FIFO-current',8.0,'response_mean_min'))} min for FIFO; budget exceedance is {pct(val('EB-UCB',8.0,'budget_violation_rate'))} versus {pct(val('FIFO-current',8.0,'budget_violation_rate'))}, while deadline misses are {pct(val('EB-UCB',8.0,'deadline_miss_rate'))} versus {pct(val('FIFO-current',8.0,'deadline_miss_rate'))}. These simulator results show a latency–budget trade-off, not a fidelity guarantee or dominance claim.");r.font.name='Times New Roman';r.font.size=Pt(9)
p=doc.add_paragraph();p.paragraph_format.space_after=Pt(4);p.paragraph_format.line_spacing=1.0
r=p.add_run('Index Terms—');r.bold=True;r.font.name='Times New Roman';r.font.size=Pt(9)
r=p.add_run('hybrid quantum–classical workflow, QPU scheduling, calibration drift, error budget, online scheduling.');r.font.name='Times New Roman';r.font.size=Pt(9)

add_head(doc,'I.','Introduction')
add_body(doc,'Hybrid quantum–classical programs execute as iterative workflows: classical code prepares parameters and data, submits one or more circuits, consumes measurement results, and then updates later work. The application is therefore a workflow over heterogeneous resources rather than an isolated circuit submission. A QPU is both scarce and noisy; backend quality varies across processors and over time, while classical tasks and dependency barriers determine when a circuit can be dispatched. Workflow models and quantum cloud scheduling studies have established the importance of representing dependencies, queues, device heterogeneity, and service constraints [1], [2].')
add_body(doc,'A second issue is that calibration-based circuit quality is time-sensitive. IBM publishes backend error properties and supports retrieval of historical snapshots; these data can guide mapping and error-aware compilation, but values may be stale by execution time [3], [4]. Recent quantum operating-system and workflow systems jointly reason about fidelity, resource use, and waiting time [5], [6], while hybrid cloud workflow studies also evaluate execution choreography on real and simulated services [10]. These works motivate an explicitly timestamped, workflow-level error allowance, while setting a demanding baseline for any new scheduling claim.')
add_body(doc,'This paper studies a narrower problem than a full quantum operating system or multi-provider orchestrator. Given a static hybrid workflow DAG, a pool of single-server QPUs, and a chronological calibration stream, how should an online dispatcher choose among ready QPU tasks and available backends while tracking a workflow-level reliability target? We propose EB-UCB, a lightweight feasibility-first heuristic. It computes additive log-failure costs, reserves predicted cost for dispatched QPU work, and uses a robust upper estimate from the recent calibration history. The implementation is deliberately simulator-oriented so the policy can be reproduced without privileged QPU access.')
add_body(doc,'The contribution is a simulation-based proof of concept for a transparent, causal calibration-to-dispatch loop, not a new general quantum operating system or a fidelity guarantee. It provides (1) an explicit DAG scheduling model with a composable log-error allowance; (2) a rolling median–MAD calibration estimator with freshness inflation; (3) executable event-driven code replayed against public IBM Heron history and compared with four baselines over 15 generated scenario blocks per load; and (4) a ledger-accounting argument and a candid latency–budget trade-off analysis. No quantum jobs were submitted, so the evidence is software-in-the-loop only.')

add_head(doc,'II.','Related Work')
add_subhead(doc,'A','Hybrid workflow and quantum resource scheduling')
add_body(doc,'Cranganore et al. formalize hybrid quantum–classical scientific workflows [1], while Ravi et al. study quantum-cloud job and resource management with fidelity and queue-time prediction [2]. Recent systems broaden this scope. QOS combines mitigation, fidelity estimation, multi-programming, and scheduling, and evaluates on real IBM devices [5]. Jha et al. profile quantum–classical FaaS workflow choreography across cloud and quantum services [10]. Qurator models dynamic hybrid DAGs across heterogeneous providers and jointly targets queue time and fidelity [6]. The 2026 qTPU system contributes a hybrid tensor-network programming model, compiler, and runtime that optimize classical cost against quantum error [11].')
add_subhead(doc,'B','Calibration-aware reliability')
add_body(doc,'Nation and Treinish use calibration-derived variability to rank circuit placements [3], and Kurniawan et al. show why historical calibration can matter when execution is delayed [4]. Fulginiti et al. (2026 preprint, accepted to a workshop) use a graph neural network to estimate pre-compilation circuit fidelity across multiple QPUs and tune the fidelity–parallelism trade-off [12]. IBM’s backend interface exposes gate, readout, and qubit properties, including historical properties queried by time [7]; the Qiskit Calibration Drift dataset provides timestamped IBM-derived events for Heron devices [8].')
add_body(doc,'The closest 2025–2026 systems already combine workflow structure, fidelity, and resource scheduling [5], [6], [10]–[12]. This manuscript does not establish superiority to those systems: it does not reproduce their circuit benchmarks, heterogeneous provider queues, compiler, or hardware experiments. Its narrower question is whether a simple causal log-cost estimator and per-workflow reservation ledger can be implemented and replayed reproducibly on a fixed IBM calibration stream. That makes the method a transparent baseline; comparison against those systems on a shared workload and fidelity metric remains future work.')
add_body(doc,'Calibration-aware mapping datasets also exist. Acampora, Schiattarella, and Troiano provide circuit-mapping instances whose features include device calibration and circuit characteristics, with labels derived from mapping solutions [9]. Such data support learning or evaluating placement decisions, whereas the present question concerns ordering dependent tasks over time as calibration observations arrive. The two choices interact in deployment: an accurate task-level cost requires the actual mapped gate locations, which this paper deliberately abstracts into backend-level medians.')

add_head(doc,'III.','System Model and Reliability Budget')
add_subhead(doc,'A','Workflow and resources')
add_body(doc,'A workflow is a DAG G=(V,E). Each task j∈V is either classical or quantum; (i,j)∈E means j cannot start until i completes. A quantum task carries N₂(j), its number of native two-qubit gates, and Nm(j), its number of measurements. At time t, R(t) is the set of ready QPU tasks and B(t) is the set of idle compatible backends. Each backend b has a single execution slot. Classical tasks use a finite CPU-worker pool and become ready using the same dependency rule. Workflows also carry an arrival time a_w, a deadline d_w, and an allowed failure probability ε_w.')
add_subhead(doc,'B','Calibration-implied task cost')
add_body(doc,'For backend b, let p₂,b(t) and pm,b(t) denote the available two-qubit-gate and readout error estimates. Under the independent-operation approximation, task success is estimated as the product of gate and readout success probabilities. For a workflow, multiplying the per-task products gives P̂_success=∏_{j∈Q_w}(1−p̂₂,b_j)^{N₂(j)}(1−p̂m,b_j)^{Nm(j)}=exp(−L̂_w), where L̂_w is the sum of task log-costs. Thus q̂_w=1−exp(−L̂_w), and q̂_w≤ε_w is algebraically equivalent to L̂_w≤−ln(1−ε_w). This derivation is exact for the stated independence model; it does not prove that the proxy equals physical circuit fidelity. Using negative logarithms makes task costs additive:')
add_eq(doc,'ℓ̂(j,b,t) = N₂(j)[−ln(1−p̂₂,b(t))] + Nm(j)[−ln(1−p̂m,b(t))].',1)
add_body(doc,'For a workflow, the nominal failure proxy is q̂_w=1−exp(−Σ_j ℓ̂_j). This is a screening proxy: it ignores coherent and correlated errors, crosstalk, idle-time decoherence, mitigation, shot variance, and circuit-specific physical placement. It must not be interpreted as measured fidelity or a rigorous upper bound on the actual output error. The budget in log space is B_w=−ln(1−ε_w), so a schedule is nominally feasible when its accumulated predicted cost does not exceed B_w.')
add_eq(doc,'B_w = −ln(1−ε_w),     Σ_{j∈Q_w} ℓ̂(j,b_j,t_j) ≤ B_w.',2)
add_subhead(doc,'C','Causal calibration estimate')
add_body(doc,'The input stream is indexed by observation time, not by an unobserved future calibration. For each property family, the dispatcher carries the latest available backend median forward until a newer observation arrives; the CZ and readout streams are aligned independently because IBM reports those properties on different calibration schedules. Every estimate is formed only from records whose observation time is no later than the simulated decision time. For the seven-day window H_b(t), the robust upper estimate is:')
add_eq(doc,'p̃_b(t) = min{0.95, max[p_b(t), median(H_b)+1.645·1.4826·MAD(H_b)]·s_b(t)}.',3)
add_body(doc,'Here p_b(t) is the current snapshot’s backend-level median; H_b contains observations in the preceding seven days up to t; and MAD is median absolute deviation. If fewer than three historical medians are available, the current value is used without a dispersion margin. The freshness age is the source-reported 90th-percentile component calibration age at the latest snapshot plus elapsed time since that observation; the larger age across CZ and readout families is used. The multiplier adds at most 10%, increasing linearly from effective age 12 to 72 hours; the backend is unavailable when either required family exceeds 72 hours. The 1.645 coefficient is a robust dispersion heuristic, not a distribution-free confidence guarantee.')
add_subhead(doc,'D','Scheduling objective')
add_body(doc,'A reference scheduling objective minimizes weighted workflow response, tardiness, and modeled failure cost, subject to DAG precedence, resource capacity, and nominal error-budget constraints. Here α, β, and γ are nonnegative preference weights; the objective is stated to locate the trade-off, not presented as a solved optimization model. EB-UCB instead uses a work-conserving online heuristic with lexicographic feasibility and urgency rules.')
add_eq(doc,'min_schedule Σ_w [α(F_w−a_w) + β max(0,F_w−d_w) + γ q̂_w],  s.t. precedence, capacity, q̂_w≤ε_w.',4)
add_body(doc,'The separation between the formulation and the implementation matters. A mixed-integer or stochastic optimization model could use the same variables and calibration cost terms, but a scheduler operating on live arrivals must act without knowing future tasks or future calibration changes. The evaluated heuristic trades global optimality for a small decision procedure that only requires the current ready queue, current resource state, and a bounded telemetry history.')


add_head(doc,'IV.','EB-UCB Online Dispatch')
add_body(doc,'At each completion or arrival event, the scheduler first releases completed resources, updates task readiness, and starts ready classical tasks on free CPU workers. It then considers every pair (j,b) formed by a ready QPU task and an idle backend. The policy estimates ℓ̂ using p̃, computes the remaining workflow allowance B_w−C_w−R_w, where C_w is realized proxy cost for completed QPU tasks and R_w is the sum of reservations for in-flight QPU tasks, and marks a candidate feasible when its predicted cost fits the remaining allowance. Among feasible pairs it selects the smallest time-to-deadline d_w−t, then the lower normalized predicted cost. This slack is a simple urgency score and does not subtract an estimated remaining critical path. If no pair is feasible, it dispatches the candidate with least deadline slack and uses normalized overshoot as a tie-breaker. This urgency fallback prevents budget accounting from permanently blocking work but means the heuristic does not guarantee a hard budget.')
add_body(doc,'After selecting a pair, the predicted cost is reserved. When the task completes, its reservation is replaced with the proxy cost computed from the latest calibration snapshot available at the task’s simulated start time. A calibration change never retroactively alters an already executed task. The backend and policy decision are recomputed at each event, so the implementation is online and non-preemptive.')
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(1);p.paragraph_format.space_after=Pt(0);p.paragraph_format.keep_with_next=True
p.add_run().add_picture(str(ROOT/'figures'/'ebucb_pipeline.png'),width=Inches(3.35))
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(2);p.paragraph_format.keep_with_next=True
r=p.add_run('Fig. 2. Implemented calibration-to-dispatch loop. Completion replaces a task reservation with its start-time proxy cost.');r.font.name='Times New Roman';r.font.size=Pt(7.5)
add_table_caption(doc,'Algorithm 1. EB-UCB dispatch at each event')
alg=doc.add_table(rows=1,cols=1);alg.alignment=WD_TABLE_ALIGNMENT.CENTER;alg.autofit=False
tw=OxmlElement('w:tblW');tw.set(qn('w:w'),str(int(3.35*1440)));tw.set(qn('w:type'),'dxa');alg._tbl.tblPr.append(tw);alg.columns[0].width=Inches(3.4)
cell=alg.cell(0,0);cell.width=Inches(3.4);set_cell_shading(cell,'F5F5F5');set_cell_margins(cell,65,80,65,80)
lines=['1  Update per-backend state using observations with time ≤ now.','2  Release completed tasks; update workflow cost and successors.','3  Dispatch ready classical tasks to available CPU workers.','4  For each ready QPU task and idle backend, compute robust cost ℓ̂.','5  Prefer budget-feasible pairs; rank by deadline slack, then normalized cost.','6  If none is feasible, use least slack, then least normalized overshoot.','7  Reserve estimated cost; on completion, replace reservation with realized proxy.']
p=cell.paragraphs[0];p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.0
for i,line in enumerate(lines):
    if i:p.add_run().add_break()
    r=p.add_run(line);r.font.name='Courier New';r.font.size=Pt(7.15)
add_body(doc,'With k=min(|R(t)|,|B(t)|) assignments made before the event loop is re-entered, the implementation rescans the ready-task/backend pairs after each assignment, giving worst-case O(k|R(t)||B(t)|) comparisons per scheduling event after backend estimates are cached. This is a direct reading of the nested loops in the released code; it is not a global optimizer and can be locally myopic.')
add_body(doc,'The reservation ledger provides a simple implementation invariant. For workflow w, let I_w=C_w+R_w, with C_w the realized costs of completed QPU tasks and R_w the predicted reservations of in-flight tasks. Every dispatched task appears once in R_w; on completion the code removes its reservation and adds its start-time calibration proxy to C_w, so I_w changes by −ℓ̂_reserved+ℓ_actual and the task is counted once on each side of the event. The feasibility test only constrains the predicted ledger at dispatch. Estimation error and the explicit urgency fallback mean it does not certify the realized budget. This is an accounting argument for the implementation, not a theorem about hardware reliability.')
add_body(doc,'The prototype is implemented as an event-driven Python simulator and was replayed with the same generated DAGs across five policies. It releases resources, enforces precedence, dispatches classical and QPU tasks, and emits per-workflow rows plus a run manifest. The accompanying pipeline figure shows exactly which observations and state feed each decision; no QPU access is required to reproduce this proof of concept.')
add_body(doc,'No calibration lookahead is used: future error changes are not visible to the policy, even though they are present in the replay trace for later simulated decision times. The policy does not delay an otherwise feasible task in anticipation of a lower-error interval. This isolates dispatch from temporal planning, while limiting possible gains from scheduling into future calibration states.')


add_head(doc,'V.','Experimental Protocol')
add_subhead(doc,'A','Calibration data')
add_body(doc,'The source is the CC BY 4.0 Qiskit Calibration Drift dataset at pinned revision bfaaec7ee00a091b1710153e251f4c69e1ac9254, which records IBM-derived properties at observation time and component calibration age [8]. We selected ibm_fez, ibm_kingston, and ibm_marrakesh, and the CZ and readout error families. The extractor removes invalid or failure-ceiling values, canonicalizes undirected CZ pairs while preserving separate directional values, deduplicates identical rows, and summarizes each backend/snapshot/property family by median and 90th-percentile error and age across components. The compact file contains 10,020 rows spanning January 31–November 8, 2026. The raw 101-MiB Parquet is not bundled; README provides its immutable revision URL and extraction command.')
# dataset count table
add_table_caption(doc,'TABLE I. TELEMETRY COVERAGE USED')
counts=DATA.groupby(['backend','property_family']).size().unstack(fill_value=0)
t=doc.add_table(rows=1,cols=3)
for i,h in enumerate(['IBM backend','CZ snapshots','Readout snapshots']):t.rows[0].cells[i].text=h
for b in ['ibm_fez','ibm_kingston','ibm_marrakesh']:
    row=t.add_row().cells;row[0].text=b;row[1].text=str(int(counts.loc[b,'cz_error']));row[2].text=str(int(counts.loc[b,'readout_error']))
fmt_table(t,[1.28,.78,.90],font=7.1)
add_body(doc,'Across the observation series, the median CZ errors range from 0.00186 to 0.00262 by backend, but the maximum snapshot medians reach 0.08501 on marrakesh and 0.05321 on kingston. Median readout errors range from 0.00922 to 0.01227, with maximum snapshot medians of 0.33032 on kingston and 0.32031 on fez. These high-tail events motivate robust summaries, but the current evaluation does not classify their physical cause or validate them with circuit outcomes.')
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(2);p.paragraph_format.space_after=Pt(0);p.paragraph_format.keep_with_next=True
r=p.add_run();r.add_picture(str(ROOT/'figures'/'calibration_medians.png'),width=Inches(3.25))
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(0);p.paragraph_format.space_after=Pt(2);p.paragraph_format.keep_with_next=True
r=p.add_run('Fig. 1. Snapshot median CZ and readout errors across the three selected Heron backends; logarithmic vertical axes.');r.font.name='Times New Roman';r.font.size=Pt(7.5)
add_subhead(doc,'B','Workloads and scenarios')
add_body(doc,'We generate fork-join workflows with seven tasks: classical preparation; two parallel QPU tasks; classical postprocessing on both branches; a join-dependent QPU refinement; and final classical processing. Each workflow has three QPU and four classical tasks. Branch tasks use 24–100 CZ gates and 8–24 measurements; the join task uses 60–180 CZs and 12–32 measurements. QPU durations are uniform within 240–480 s for branch tasks and 300–600 s for the join; classical durations range from 120–480 s by stage. Deadline slack is uniform from 1.5–4 h; ε_w is uniform from 0.40–0.75. The simulator has 16 CPU workers and three single-server QPUs. These service times and operation counts are synthetic scenario inputs, not measurements from IBM. Arrivals are Poisson at 4 or 8 workflows/hour. Five seven-day UTC windows start on 2026-02-15, 2026-04-15, 2026-06-15, 2026-08-15, and 2026-10-15; each is paired with seeds 17, 23, and 42. All policies receive identical tasks and chronological calibration inputs.')
low_n=int(SUMMARY[SUMMARY.rate_per_hour==4.0].n.iloc[0]);high_n=int(SUMMARY[SUMMARY.rate_per_hour==8.0].n.iloc[0])
add_body(doc,f'The scheduler drains each arrival interval until every admitted workflow completes, including the post-arrival queue tail. Because arrivals are Poisson, counts vary by seed. Across five windows and three seeds, each policy processes {low_n:,} workflows at 4/hour and {high_n:,} at 8/hour ({low_n+high_n:,} total), with {(low_n+high_n)*3:,} QPU tasks and {(low_n+high_n)*4:,} classical tasks across both rates. Five policies run on the same 30 window/seed/load blocks. The simulator samples task runtimes rather than predicting them, so response time measures dispatch and queueing under the declared service distributions.')

add_subhead(doc,'C','Baselines and metrics')
add_body(doc,'For completed workflow w, response is R_w=F_w−a_w, where F_w is the final task completion time. QPU wait for task j is S_j−r_j, where r_j is its time of dependency readiness; the workflow value is the mean across its three QPU tasks. With aggregate proxy cost L_w, the failure proxy is q_w=1−exp(−L_w). Budget exceedance is 1[L_w>B_w], and a deadline miss is 1[F_w>d_w]. These metrics describe simulated service and the model-derived error proxy; they do not represent user-visible completion on an external quantum cloud.')
add_body(doc,'FIFO-current dispatches the oldest ready QPU task and assigns the idle backend with lowest current median cost. EDF-current selects the ready task with the earliest workflow deadline and uses the same backend rule. Min-risk chooses the ready task/backend pair with lowest current median cost, without workflow-budget feasibility. EB-current applies the budget and reservation rule to current medians; EB-UCB substitutes the robust estimate. We report response time, P95 response, mean QPU wait, independent-error failure proxy, budget-exceedance fraction, and deadline-miss rate. Tables report the mean and sample SD across 15 window/seed blocks at each arrival rate; these are descriptive block summaries, not confidence intervals or hypothesis tests. A separate causal check predicts each next telemetry summary using only the prior seven days and tests whether the upper estimate covers it.')

add_head(doc,'VI.','Results and Discussion')
add_body(doc,'Table I reports telemetry coverage, Table II the median of snapshot medians, Table III the causal estimator check, Table IV the same-time budget diagnostic, and Tables V–VI the simulator outputs. For reference, across the extracted telemetry rows the median of snapshot-level CZ medians was 0.00262, 0.00186, and 0.00250 for fez, kingston, and marrakesh; the corresponding readout medians were 0.01227, 0.00922, and 0.01172. These are summaries of calibration properties, not circuit outcomes.')
add_table_caption(doc,'TABLE II. MEDIAN OF SNAPSHOT MEDIANS')
med=DATA.groupby(['backend','property_family']).median_error.median().unstack()
t=doc.add_table(rows=1,cols=3)
for i,h in enumerate(['Backend','CZ error','Readout error']):t.rows[0].cells[i].text=h
for b in ['ibm_fez','ibm_kingston','ibm_marrakesh']:
    row=t.add_row().cells;row[0].text=b.replace('ibm_','');row[1].text=f'{med.loc[b,"cz_error"]:.5f}';row[2].text=f'{med.loc[b,"readout_error"]:.5f}'
fmt_table(t,[1.35,.9,1.0],font=7.2)
add_body(doc,f"A causal next-observation check uses {int(FVALID.n.sum())} one-step predictions. The seven-day upper estimate covers {pct(float(FVALID[FVALID.property_family=='cz_error'].coverage.iloc[0]))} of next CZ summaries and {pct(float(FVALID[FVALID.property_family=='readout_error'].coverage.iloc[0]))} of next readout summaries. The corresponding point-prediction MAEs are {FVALID[FVALID.property_family=='cz_error'].point_mae.iloc[0]:.6f} and {FVALID[FVALID.property_family=='readout_error'].point_mae.iloc[0]:.6f}; mean upper-estimate margins are {FVALID[FVALID.property_family=='cz_error'].mean_upper_margin.iloc[0]:.6f} and {FVALID[FVALID.property_family=='readout_error'].mean_upper_margin.iloc[0]:.6f}. These are empirical coverages over the observed sequence and do not imply a formal probabilistic guarantee.")
add_table_caption(doc,'TABLE III. CAUSAL NEXT-SNAPSHOT FORECAST CHECK')
t=doc.add_table(rows=1,cols=5)
for i,h in enumerate(['Property','N','Coverage','Point MAE','Mean upper margin']):t.rows[0].cells[i].text=h
for _,rr in FVALID.iterrows():
    row=t.add_row().cells;row[0].text=rr.property_family.replace('_error','');row[1].text=str(int(rr.n));row[2].text=pct(rr.coverage);row[3].text=f"{rr.point_mae:.6f}";row[4].text=f"{rr.mean_upper_margin:.6f}"
fmt_table(t,[.75,.35,.55,.62,.8],font=6.8)
add_body(doc,'We also measured an optimistic same-time budget diagnostic: at workflow arrival, each of its three QPU tasks is independently assigned the backend with the lowest current-median cost. This ignores QPU capacity, future calibration changes, and physical mapping. It is therefore a useful lower-cost comparison within the snapshot model, not an end-to-end schedule.')
add_table_caption(doc,'TABLE IV. OPTIMISTIC SAME-TIME BUDGET FEASIBILITY')
t=doc.add_table(rows=1,cols=5)
for i,h in enumerate(['Arrival rate','N','Over budget','Median min. log-cost','Median budget']):t.rows[0].cells[i].text=h
for _,rr in FEAS.iterrows():
    row=t.add_row().cells;row[0].text=f"{rr.rate_per_hour:.0f}/h";row[1].text=str(int(rr.n));row[2].text=pct(rr.infeasible_rate);row[3].text=f"{rr.median_min_cost:.3f}";row[4].text=f"{rr.median_budget:.3f}"
fmt_table(t,[.65,.35,.62,.85,.68],font=6.8)
add_body(doc,f"Even this optimistic per-task backend choice exceeds the sampled workflow budget for {pct(float(FEAS[FEAS.rate_per_hour==4.0].infeasible_rate.iloc[0]))} of workflows at 4/hour and {pct(float(FEAS[FEAS.rate_per_hour==8.0].infeasible_rate.iloc[0]))} at 8/hour. Because it assumes no queueing and keeps one arrival-time calibration snapshot throughout each workflow, this diagnosis points to an incompatibility between the sampled budget range and circuit-size/error proxy for a substantial fraction of generated workflows. It helps explain why online dispatch alone cannot drive violations close to zero.")
add_body(doc,f"At 4 workflows/hour, pooled across completed workflows, FIFO has a {fmt(val('FIFO-current',4.0,'response_mean_min'))}-min mean response, {fmt(val('FIFO-current',4.0,'response_p95_min'))}-min 95th percentile, and {fmt(val('FIFO-current',4.0,'risk_mean'),3)} failure proxy. EB-UCB is {fmt(val('EB-UCB',4.0,'response_mean_min'))} min mean response and {fmt(val('EB-UCB',4.0,'risk_mean'),3)} proxy. Its budget-exceedance rate is {pct(val('EB-UCB',4.0,'budget_violation_rate'))}, versus {pct(val('FIFO-current',4.0,'budget_violation_rate'))} for FIFO. All methods have low deadline misses at this load.")
add_results_table(doc,4.0)
add_body(doc,f"At 8 workflows/hour, the queue is substantially loaded. Pooled across completed workflows, FIFO averages {fmt(val('FIFO-current',8.0,'response_mean_min'))} min response, while EB-UCB averages {fmt(val('EB-UCB',8.0,'response_mean_min'))} min and has a {fmt(val('EB-UCB',8.0,'response_p95_min'))}-min 95th percentile. EB-UCB's budget-exceedance rate is {pct(val('EB-UCB',8.0,'budget_violation_rate'))}, versus {pct(val('FIFO-current',8.0,'budget_violation_rate'))} for FIFO; the failure proxies are {fmt(val('EB-UCB',8.0,'risk_mean'),3)} and {fmt(val('FIFO-current',8.0,'risk_mean'),3)}, respectively. Deadline misses are {pct(val('EB-UCB',8.0,'deadline_miss_rate'))} for EB-UCB and {pct(val('FIFO-current',8.0,'deadline_miss_rate'))} for FIFO. The outcome is a measurable trade-off rather than dominance.")
add_results_table(doc,8.0)
add_body(doc,f"Pooled across completed workflows, against the median-only error-budget policy EB-UCB has {fmt(val('EB-UCB',8.0,'response_mean_min'))} versus {fmt(val('EB-current',8.0,'response_mean_min'))} min mean response, while its budget-exceedance rate is {pct(val('EB-UCB',8.0,'budget_violation_rate'))} versus {pct(val('EB-current',8.0,'budget_violation_rate'))}; their failure proxies are {fmt(val('EB-UCB',8.0,'risk_mean'),3)} and {fmt(val('EB-current',8.0,'risk_mean'),3)}. In this replay, the rolling upper margin slightly changes queue order but does not improve budget compliance over using the same reservation rule with current medians. The independent forecast check supports using a margin as a conservative heuristic, but its imperfect coverage and the lack of circuit-specific edge mapping prevent translating coverage into guaranteed workflow reliability.")
add_body(doc,f"At high load, EB-current reduces budget exceedance by {100*(val('FIFO-current',8.0,'budget_violation_rate')-val('EB-current',8.0,'budget_violation_rate')):.1f} percentage points against FIFO, while its mean response rises from {fmt(val('FIFO-current',8.0,'response_mean_min'))} to {fmt(val('EB-current',8.0,'response_mean_min'))} min. EB-UCB lowers exceedance by {100*(val('FIFO-current',8.0,'budget_violation_rate')-val('EB-UCB',8.0,'budget_violation_rate')):.1f} points, but its deadline-miss rate is {pct(val('EB-UCB',8.0,'deadline_miss_rate'))} compared with {pct(val('FIFO-current',8.0,'deadline_miss_rate'))} under FIFO. Thus a reliability target tracked by a local reservation does not automatically translate into a better deadline/reliability operating point; users need an explicit service-level trade-off policy.")
add_body(doc,f"Across the low-load scenarios, mean response ranges from {fmt(val('FIFO-current',4.0,'response_mean_min'))} min for FIFO to {fmt(val('Min-risk',4.0,'response_mean_min'))} min for Min-risk, while the failure proxy spans only {fmt(min(val(p,4.0,'risk_mean') for p in ('FIFO-current','EDF-current','Min-risk','EB-current','EB-UCB')),3)}–{fmt(max(val(p,4.0,'risk_mean') for p in ('FIFO-current','EDF-current','Min-risk','EB-current','EB-UCB')),3)}. This is consistent with a lightly queued pool: little room exists for task reordering to change completion time, and the three backend-level median error profiles are relatively close for typical snapshots.")
add_body(doc,f"At high load, FIFO and EDF remain close in average response ({fmt(val('FIFO-current',8.0,'response_mean_min'))} and {fmt(val('EDF-current',8.0,'response_mean_min'))} min), while Min-risk reaches {fmt(val('Min-risk',8.0,'response_mean_min'))} min and a {fmt(val('Min-risk',8.0,'response_p95_min'))}-min P95. The two error-budget policies lower their budget-exceedance rate, but their 95th-percentile responses exceed 1,000 min. This suggests that a local budget check can interact poorly with dependency-driven queue selection when the pool is saturated; a whole-workflow admission test or a future-cost reserve is a more appropriate next design than merely increasing the local error penalty.")
low_excess=100*SUMMARY.budget_violation_rate.min(); high_excess=100*SUMMARY.budget_violation_rate.max()
add_body(doc,f'The principal result is a trade-off, not dominance. The tested policy’s feasibility-first dispatch orders ready work around predicted workflow allowance, but can leave a long response-time tail when the system is heavily loaded. Its fallback keeps the system work-conserving and avoids permanent starvation while permitting budget overshoots. Budget-exceedance rates range from {low_excess:.0f}% to {high_excess:.0f}% across the tested policy/load combinations, showing that the sampled ε_w range and independent-error proxy often disagree with generated circuit costs; the budget is not a guarantee. The modest high-load reduction in exceedances should not be read as a robust fidelity improvement. Admission control, workflow-level lookahead, circuit-specific mapping, calibrated uncertainty, and workload traces from additional application classes are needed before deployment claims.')

add_head(doc,'VII.','Limitations, Threats to Validity, and Conclusion')
add_body(doc,'This study is limited to three IBM superconducting processors and two calibration families. It does not submit circuits to hardware, replay actual QPU queue traces, model provider outages, or use physical-qubit mappings for each circuit. The workload DAGs and durations are synthetic; the evaluation covers five weekly telemetry windows and three generated-workload seeds per window/load. The product-of-independent-error proxy omits correlated noise, coherence, crosstalk, error mitigation, shot counts, and statistical uncertainty. Median–MAD drift inflation is a heuristic and carries no formal coverage guarantee. Accordingly, neither the reported proxy nor budget-exceedance rate is a measured quantum fidelity or operational service-level guarantee.')
add_subhead(doc,'A','Threats to validity')
add_body(doc,'Construct validity is the main threat. A product of calibrated gate and readout error probabilities approximates the chance that a sequence of operations is error-free only under independence and the selected operation counts. It excludes coherent errors, correlated events, crosstalk, circuit topology, physical-qubit placement, idle coherence, error mitigation, and finite-shot statistics. A device-level median cannot predict a mapped circuit’s actual edge-level error, and a failure proxy cannot be substituted for fidelity against an ideal output distribution. We therefore use “budget exceedance” only as a model-internal comparison.')
add_body(doc,'Internal validity is limited by generated workloads and service-time distributions. The fork-join graph is reproducible but does not represent a trace collected from an HPC system or a particular QAOA/VQE application. The 4- and 8-workflow/hour rates, deadline range, error allowances, and service times are controlled scenario settings. We used common random-number inputs across policies to support paired comparisons, with three workload seeds per weekly window and load. Aggregate means and standard deviations are descriptive and can mask week-to-week variation; no significance tests or confidence intervals are claimed.')
add_body(doc,'External validity is limited to public calibration histories for three IBM Heron backends over the observed period. The simulator uses idealized single-server QPUs, no cloud queue or network latency trace, no downtime, and a fixed 16-worker classical pool. Another provider, processor generation, QPU multiplexing model, queue discipline, or workload mix can produce a different trade-off. The causal forecast check measures prediction coverage on the same historical series used to drive scheduling, not generalization to a future hardware release. These limits mean that the policy should be treated as a reproducible design baseline, not an operational recommendation.')
add_body(doc,f"We implemented and evaluated an online scheduler that uses timestamped calibration summaries and per-workflow log-error budgets. At 8 workflows/hour, EB-UCB reduces model-internal budget exceedance from {pct(val('FIFO-current',8.0,'budget_violation_rate'))} under FIFO to {pct(val('EB-UCB',8.0,'budget_violation_rate'))}, while pooled mean response rises from {fmt(val('FIFO-current',8.0,'response_mean_min'))} to {fmt(val('EB-UCB',8.0,'response_mean_min'))} min and deadline misses rise from {pct(val('FIFO-current',8.0,'deadline_miss_rate'))} to {pct(val('EB-UCB',8.0,'deadline_miss_rate'))}. The experiment establishes a reproducible simulator baseline and exposes a concrete failure mode of conservative local budget handling. The immediate next step is to couple admission control and workflow lookahead with circuit-specific calibration data, then validate against real execution and queue telemetry before making stronger performance claims.")
add_subhead(doc,'','Data and code availability')
add_body(doc,'The manuscript package contains the simulator, calibration-summary CSV, per-workflow outputs, run manifest, and extraction code. The raw source dataset is available at the pinned Hugging Face revision specified in the package README under CC BY 4.0. Reported tables are generated directly from results/summary.csv; no result values were manually inserted.')

# References
p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.CENTER;p.paragraph_format.space_before=Pt(5);p.paragraph_format.space_after=Pt(2);p.paragraph_format.keep_with_next=True
r=p.add_run('REFERENCES');r.bold=True;r.font.name='Times New Roman';r.font.size=Pt(10)
refs=[
'[1] S. S. Cranganore, V. De Maio, I. Brandic, and E. Deelman, “Paving the way to hybrid quantum–classical scientific workflows,” Future Generation Computer Systems, vol. 158, pp. 346–366, Sep. 2024, doi: 10.1016/j.future.2024.04.030.',
'[2] G. S. Ravi, K. N. Smith, P. Murali, and F. T. Chong, “Adaptive job and resource management for the growing quantum cloud,” in Proc. IEEE Int. Conf. Quantum Computing and Engineering (QCE), 2021, pp. 301–312, doi: 10.1109/QCE52317.2021.00047.',
'[3] P. D. Nation and M. Treinish, “Suppressing quantum circuit errors due to system variability,” PRX Quantum, vol. 4, no. 1, Art. no. 010327, Mar. 2023, doi: 10.1103/PRXQuantum.4.010327.',
'[4] H. Kurniawan, L. Rodríguez-Soriano, D. Cuomo, C. G. Almudéver, and F. Garcia-Herrero, “On the use of calibration data in error-aware compilation techniques for NISQ devices,” in Proc. IEEE Int. Conf. Quantum Computing and Engineering (QCE), 2024, pp. 338–348, doi: 10.1109/QCE60285.2024.00048.',
'[5] E. Giortamis, F. Romão, N. Tornow, and P. Bhatotia, “QOS: Quantum operating system,” in Proc. 19th USENIX Symp. Operating Systems Design and Implementation (OSDI), 2025, pp. 429–447.',
'[6] S. Pehlivanoglu, U. de Muelenaere, P. Kogge, and A. Sabry, “Qurator: Scheduling hybrid quantum-classical workflows across heterogeneous cloud providers,” arXiv:2604.05505, v3, Oct. 25, 2026.',
'[7] IBM Quantum, “View backend details,” IBM Quantum Documentation. [Online]. Available: https://quantum.cloud.ibm.com/docs/en/guides/qpu-information. Accessed: Nov. 8, 2026.',
'[8] phanerozoic, “Qiskit Calibration Drift,” Hugging Face Datasets, revision bfaaec7ee00a091b1710153e251f4c69e1ac9254, 2026. [Online]. Available: https://huggingface.co/datasets/phanerozoic/qiskit-calibration-drift. License: CC BY 4.0.',
'[9] G. Acampora, R. Schiattarella, and A. Troiano, “A dataset for quantum circuit mapping,” Data in Brief, vol. 39, Art. no. 107526, Dec. 2021, doi: 10.1016/j.dib.2021.107526.',
'[10] V. Jha et al., “Choreography and profiling of quantum-classical FaaS workflows on hybrid clouds,” in Proc. 25th IEEE Int. Symp. Cluster, Cloud and Internet Computing (CCGrid), 2025, pp. 1–11, doi: 10.1109/CCGRID64434.2025.00069.',
'[11] N. Tornow, E. Giortamis, D. Sprokholt, C. Mendl, and P. Bhatotia, “qTPU: Hybrid tensor networks for quantum-classical acceleration,” in Proc. 20th USENIX Symp. Operating Systems Design and Implementation (OSDI), 2026, pp. 2367–2388.',
'[12] I. Fulginiti et al., “Fidelity-aware scheduling of quantum circuits on multi-QPU systems,” arXiv:2609.09980, Sep. 2026. Accepted to the 2nd Int. Workshop on Software Frameworks and Workload Management for Quantum and HPC Ecosystems (SFWM), co-located with SC26.'
]
for ref in refs:
    p=doc.add_paragraph();p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY;p.paragraph_format.left_indent=Inches(.18);p.paragraph_format.first_line_indent=Inches(-.18);p.paragraph_format.space_after=Pt(0);p.paragraph_format.line_spacing=1.0
    r=p.add_run(ref);r.font.name='Times New Roman';r.font.size=Pt(7.3)

# Set page language and remove any default header/footer content.
for s in doc.sections:
    for hp in (s.header,s.footer,s.first_page_header,s.first_page_footer):
        for p in hp.paragraphs: p.text=''
doc.core_properties.title='Noise-Aware, Error-Budgeted Scheduling of Hybrid Quantum–Classical Workflows Using Time-Varying QPU Calibration Data'
doc.core_properties.author='Dr. Tanvir H. Sardar'
doc.core_properties.subject='IEEE conference manuscript with reproducible calibration-driven scheduler simulation'
doc.save(OUT)
print(OUT)
