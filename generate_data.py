"""
HR Dashboard - Sample Data Generator
Generates a realistic Excel dataset for the HR Dashboard.
Run: python3 generate_data.py
"""

import pandas as pd
import numpy as np
from datetime import datetime
import random
import os

random.seed(42)
np.random.seed(42)

# ── Company Configuration ─────────────────────────────────────
ANNUAL_REVENUE     = 85_000_000
AVG_SALARY         = 85_000
HEADCOUNT          = 480
MONTHS = pd.date_range(start='2024-01-01', periods=12, freq='MS').strftime('%Y-%m').tolist()

DEPTS = {
    'Engineering':       120,
    'Sales':              80,
    'Operations':         60,
    'Marketing':          40,
    'HR':                 30,
    'Finance':            35,
    'Customer Success':   55,
    'Product':            60,
}

ROLES = {
    'Engineering':       ['Software Engineer', 'Senior Engineer', 'Staff Engineer', 'Engineering Manager', 'DevOps Engineer'],
    'Sales':             ['SDR', 'Account Executive', 'Senior AE', 'Sales Manager', 'VP Sales'],
    'Operations':        ['Operations Analyst', 'Operations Manager', 'Process Specialist', 'Director of Operations'],
    'Marketing':         ['Marketing Analyst', 'Content Manager', 'Growth Manager', 'Marketing Director'],
    'HR':                ['HR Generalist', 'Recruiter', 'HR Business Partner', 'HR Director'],
    'Finance':           ['Financial Analyst', 'Senior Analyst', 'Finance Manager', 'Controller'],
    'Customer Success':  ['CS Specialist', 'CS Manager', 'CS Director', 'Account Manager'],
    'Product':           ['Product Analyst', 'Product Manager', 'Senior PM', 'Director of Product'],
}

SALARY_RANGES = {
    'Engineering':      (80_000, 185_000),
    'Sales':            (60_000, 160_000),
    'Operations':       (55_000, 120_000),
    'Marketing':        (60_000, 130_000),
    'HR':               (55_000, 120_000),
    'Finance':          (65_000, 140_000),
    'Customer Success': (55_000, 110_000),
    'Product':          (80_000, 160_000),
}

GENDERS    = ['Male', 'Female', 'Non-Binary']
AGE_GROUPS = ['Under 25', '25-34', '35-44', '45-54', '55+']
LEVELS     = ['Individual Contributor', 'Senior', 'Lead/Principal', 'Manager', 'Director', 'VP+']


# ── Sheet 1 : KPI_Summary ────────────────────────────────────
def gen_kpi_summary():
    rows = []
    for i, month in enumerate(MONTHS):
        t = i / 11          # progress 0→1
        rows.append({
            'Month':                  month,
            'Headcount':              int(465 + t * 15 + random.randint(-3, 3)),
            'TurnoverRate':           round(13.2  - t * 2.1  + random.uniform(-0.3,  0.3),  1),
            'TimeToHire_Days':        round(31.0  - t * 5.0  + random.uniform(-1.0,  1.0),  1),
            'CostPerHire':            int( 4800   - t * 400  + random.randint(-100,  100)),
            'EngagementScore':        round(68.0  + t * 8.0  + random.uniform(-0.5,  0.5),  1),
            'AbsenteeismRate':        round(3.8   - t * 0.6  + random.uniform(-0.1,  0.1),  1),
            'RevenuePerEmployee':     int(174_000 + t * 8000 + random.randint(-1000, 1000)),
            'PerformanceAvg':         round(3.4   + t * 0.3  + random.uniform(-0.05, 0.05), 2),
            'TrainingHoursPerEmp':    round(24.0  + t * 12.0 + random.uniform(-1.0,  1.0),  1),
            'TrainingROI':            int( 220    + t * 80   + random.randint(-10,   10)),
            'OfferAcceptRate':        round(78.0  + t * 8.0  + random.uniform(-1.0,  1.0),  1),
            'QualityOfHire':          round(72.0  + t * 6.0  + random.uniform(-0.5,  0.5),  1),
            'RetentionRate':          round(86.8  + t * 2.1  + random.uniform(-0.3,  0.3),  1),
            'eNPS':                   int( 22     + t * 18   + random.randint(-2,    2)),
        })
    return pd.DataFrame(rows)


# ── Sheet 2 : Employees ──────────────────────────────────────
def gen_employees():
    rows = []
    emp_id = 1001
    for dept, count in DEPTS.items():
        smin, smax = SALARY_RANGES[dept]
        for _ in range(count):
            perf = float(np.clip(np.random.normal(3.5, 0.65), 1.0, 5.0))
            eng  = float(np.clip(np.random.normal(72,  12),   30, 100))
            rows.append({
                'EmpID':            f'EMP{emp_id:04d}',
                'Department':       dept,
                'Role':             random.choice(ROLES[dept]),
                'Level':            random.choices(LEVELS, weights=[30,25,18,14,9,4])[0],
                'Salary':           int(random.uniform(smin, smax)),
                'Tenure_Years':     round(random.uniform(0.3, 14.0), 1),
                'Status':           random.choices(['Active','Resigned','Terminated'],
                                                   weights=[88, 8, 4])[0],
                'Gender':           random.choices(GENDERS,    weights=[52, 45, 3])[0],
                'AgeGroup':         random.choices(AGE_GROUPS, weights=[12, 35, 30, 16, 7])[0],
                'PerformanceScore': round(perf, 1),
                'EngagementScore':  round(eng,  1),
            })
            emp_id += 1
    return pd.DataFrame(rows)


# ── Sheet 3 : Recruitment ────────────────────────────────────
def gen_recruitment():
    rows = []
    for month in MONTHS:
        for dept in DEPTS:
            hires = random.randint(0, 4)
            if hires > 0:
                rows.append({
                    'Month':           month,
                    'Department':      dept,
                    'NewHires':        hires,
                    'TimeToHire_Days': round(max(7, np.random.normal(28, 8)), 1),
                    'CostPerHire':     int(max(1000, np.random.normal(4500, 800))),
                    'OfferAcceptRate': round(np.clip(np.random.normal(80, 10), 50, 100), 1),
                    'QualityScore':    round(np.clip(np.random.normal(73, 10), 40, 100), 1),
                })
    return pd.DataFrame(rows)


# ── Sheet 4 : Attendance ─────────────────────────────────────
def gen_attendance():
    rows = []
    for i, month in enumerate(MONTHS):
        for dept, base in DEPTS.items():
            workdays   = 21 if (i % 12) in [1, 3, 5, 8, 10] else 22
            headcount  = base + random.randint(-3, 3)
            total_days = headcount * workdays
            absence    = int(total_days * random.uniform(0.025, 0.045))
            rows.append({
                'Month':            month,
                'Department':       dept,
                'Headcount':        headcount,
                'AbsenceDays':      absence,
                'TotalWorkdays':    workdays,
                'AbsenteeismRate':  round(absence / total_days * 100, 2),
            })
    return pd.DataFrame(rows)


# ── Sheet 5 : Training ───────────────────────────────────────
PROGRAMS = [
    'Leadership Development', 'Technical Skills', 'Compliance & Ethics',
    'Sales Effectiveness', 'Customer Service Excellence', 'Digital Transformation',
    'Project Management', 'Communication & Influence', 'DEI & Belonging',
]

def gen_training():
    rows = []
    for month in MONTHS:
        for dept in DEPTS:
            participants = random.randint(4, 28)
            hours        = participants * random.uniform(4, 16)
            cost         = hours * random.uniform(50, 150)
            pre          = float(np.clip(np.random.normal(65, 10), 40, 85))
            post         = float(np.clip(pre + np.random.normal(13, 4), 50, 100))
            rows.append({
                'Month':          month,
                'Department':     dept,
                'Program':        random.choice(PROGRAMS),
                'Participants':   participants,
                'TotalHours':     round(hours, 1),
                'TotalCost':      int(cost),
                'CompletionRate': round(np.clip(np.random.normal(88, 7), 60, 100), 1),
                'PreScore':       round(pre,  1),
                'PostScore':      round(post, 1),
            })
    return pd.DataFrame(rows)


# ── Sheet 6 : Business_Config ────────────────────────────────
def gen_config():
    data = {
        'Metric': [
            'AnnualRevenue', 'AverageSalary', 'CurrentHeadcount',
            'IndustryTurnoverBenchmark', 'IndustryTimeToHireBenchmark',
            'IndustryCostPerHireBenchmark', 'IndustryEngagementBenchmark',
            'IndustryAbsenteeismBenchmark', 'IndustryRevenuePerEmpBenchmark',
            'TurnoverReplacementCostMultiplier', 'AbsenteeismIndirectCostMultiplier',
            'TargetTurnoverRate', 'TargetEngagementScore', 'TargetTimeToHire',
            'TargetCostPerHire', 'TargetRevenuePerEmployee',
        ],
        'Value': [
            85_000_000, 85_000, 480,
            14.0, 36, 4_700, 66, 3.5, 160_000,
            1.5, 1.5,
            10.0, 78, 25, 4_000, 185_000,
        ],
        'Unit': [
            'USD', 'USD', 'Employees',
            '%', 'Days', 'USD', 'Score/100', '%', 'USD',
            'multiplier', 'multiplier',
            '%', 'Score/100', 'Days', 'USD', 'USD',
        ],
        'Description': [
            'Total Annual Revenue',
            'Average Annual Salary across all employees',
            'Total Current Headcount',
            'Industry Average Annual Turnover Rate',
            'Industry Average Time to Hire (days)',
            'Industry Average Cost per Hire',
            'Industry Average Employee Engagement Score',
            'Industry Average Absenteeism Rate',
            'Industry Average Revenue per Employee',
            'Cost to replace one employee as multiple of annual salary',
            'Multiplier for indirect / productivity absenteeism costs',
            'Target Annual Turnover Rate',
            'Target Employee Engagement Score',
            'Target Time to Hire (days)',
            'Target Cost per Hire ($)',
            'Target Revenue per Employee ($)',
        ],
    }
    return pd.DataFrame(data)


# ── Write Excel ───────────────────────────────────────────────
def main():
    os.makedirs('data', exist_ok=True)
    path = 'data/hr_sample_data.xlsx'

    dfs = {
        'KPI_Summary':    gen_kpi_summary(),
        'Employees':      gen_employees(),
        'Recruitment':    gen_recruitment(),
        'Attendance':     gen_attendance(),
        'Training':       gen_training(),
        'Business_Config':gen_config(),
    }

    with pd.ExcelWriter(path, engine='openpyxl') as writer:
        for sheet, df in dfs.items():
            df.to_excel(writer, sheet_name=sheet, index=False)

    print(f"\n✅  Excel file generated → {path}")
    for sheet, df in dfs.items():
        print(f"   {sheet:<22}: {len(df):>4} rows")

if __name__ == '__main__':
    main()
