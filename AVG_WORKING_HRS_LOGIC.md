# Average Working Hours Calculator

Script: `avg_workingHrs_calcultor.py`

## What this script does

It estimates **each task’s hours for a selected month** using a fair average-daily-hours method — **not** by dividing `Σ Time Spent` by calendar days, and **not** by only summing that month’s worklogs.

### Core formula

```text
TOTAL_TASK_HOURS
    = SUM(all worklogs for the issue)
      (fallback: Jira Σ Time Spent if worklogs have no Issue Key)

TOTAL_VALID_ACTIVE_DAYS
    = valid working days between Created and Updated
      (exclude weekoffs, leave, holidays, no-work days)

AVERAGE_DAILY_TASK_HOURS
    = TOTAL_TASK_HOURS / TOTAL_VALID_ACTIVE_DAYS

SELECTED_MONTH_ACTIVE_DAYS
    = valid working days between
        max(Created, MonthStart)  and  min(Updated, MonthEnd)

SELECTED_MONTH_TASK_HOURS
    = AVERAGE_DAILY_TASK_HOURS × SELECTED_MONTH_ACTIVE_DAYS
```

### Example

```text
Created = 25-Jun-2026
Updated = 06-Jul-2026
Total worklog/Sigma hours = 50

Valid active working days (25 Jun → 6 Jul) = 10
Average daily = 50 / 10 = 5 hrs/day

July overlap = 01 Jul → 06 Jul
Valid July working days = 4

July task hours = 4 × 5 = 20
```

---

## Task eligibility (`--month 07` = July)

A task is included when:

```text
Created <= July 31
AND
Updated >= July 1
```

| Created | Updated | In July report? |
|---------|---------|-----------------|
| 25-Jun | 06-Jul | Yes |
| 25-Jun | 25-Jun | No |
| 10-Aug | 15-Aug | No |
| 02-Jul | 07-Aug | Yes |

Created date alone does **not** decide the month’s hours — only eligibility and the active date range do.

---

## Working days (what is removed)

A day is **not** a working day if any of these apply (priority order):

1. Company holiday  
2. Employee leave (from leave sheet)  
3. Configured weekoff (default: **Sunday only**)  
4. No work recorded that day (when Tempo timesheet covers that date)

### Weekoffs

Default:

```python
WEEKOFF_DAYS = {6}   # Sunday
```

Change via CLI:

```bash
--weekoff sun
--weekoff sat,sun
```

Your leave sheet already marks Saturday holidays (e.g. `SH`), so default is **Sunday only**; other Saturdays can still be workdays unless leave/attendance marks them off.

---

## Input files

### 1. Jira task sheet (`--jira`)

Example: `Jira sheet 100.xlsx`

Uses: Issue key, Issue Type, Summary, Assignee, Created, Updated, Σ Time Spent (fallback hours).

### 2. Worklogs (`--worklogs`)

Supported formats:

| Format | Behavior |
|--------|----------|
| **Issue-level** (Issue Key + Date + Time Spent + Author) | Sum hours per Issue Key (preferred) |
| **Tempo user/day** (User + daily columns) — your `worklogs_01.07.2026_31.07.2026.xlsx` | No Issue Key → hours fall back to Σ Time Spent; daily columns used for “employee worked that day” |

### 3. Leave / attendance (`--leave`)

Example: `Sample leave sheet.xlsx` (calendar matrix)

- Row: `Employee Name` + weekday headers  
- Row: day numbers 1–31  
- Employee rows with codes:

| Code | Meaning in script |
|------|-------------------|
| `AV`, `OS`, `WFH` | Present / working |
| `LL`, `CL`, `CL1`, `CL2`, `AL`, `SH`, `ComOff`, `FH`, … | Non-working (leave / off) |
| `SUN` | Sunday / weekoff marker |
| blank | Treated as non-working for that employee |

Also supports a flat leave table: `Employee` + `Leave Date` or `From Date`/`To Date`.

---

## How to run

From the project folder:

```powershell
python avg_workingHrs_calcultor.py `
  --jira ".\Jira sheet 100.xlsx" `
  --worklogs ".\worklogs_01.07.2026_31.07.2026.xlsx" `
  --leave ".\Sample leave sheet.xlsx" `
  --month 07 `
  --output ".\Jira_task_avg_July.xlsx"
```

August example:

```powershell
python avg_workingHrs_calcultor.py `
  --jira ".\Jira sheet 100.xlsx" `
  --worklogs ".\worklogs_01.07.2026_31.07.2026.xlsx" `
  --leave ".\Sample leave sheet.xlsx" `
  --month 08 `
  --output ".\Jira_task_avg_August.xlsx"
```

Every Saturday off:

```powershell
python avg_workingHrs_calcultor.py --jira ".\Jira sheet 100.xlsx" --worklogs ".\worklogs_01.07.2026_31.07.2026.xlsx" --leave ".\Sample leave sheet.xlsx" --month 07 --weekoff sat,sun --output ".\out.xlsx"
```

`--month` default = current calendar month. Year for `07` = system year (e.g. 2026).

---

## Output Excel (3 sheets)

### Sheet 1 — Task Details

| Column | Meaning |
|--------|---------|
| Issue Key / Type / Summary / Assignee | From Jira |
| Created / Updated | Task active range |
| Total Worklog Hours | All worklogs (or Sigma fallback) |
| Total Active Working Days | Valid days Created→Updated |
| Average Daily Task Hours | Total ÷ active days |
| Selected Month … | Month window |
| Selected Month Working Days | Valid days in month overlap |
| **Selected Month Task Hours** | **Final July/Aug hours** |

### Sheet 2 — Employee Summary

Employee, task count, working days, total month hours, average hours per working day.

### Sheet 3 — Calculation Details

Audit trail: active days list, excluded leave/holidays/weekoffs/no-work days, averages, final hours, notes.

---

## Terminal summary (example)

```text
Selected Month: July 2026
Month window: 2026-07-01 .. 2026-07-31
Weekoffs: [6] (0=Mon .. 6=Sun)
Worklog source: Tempo user/day timesheet ...
Leave source: Attendance calendar leave sheet ...

Total Jira Tasks Read: 619
Tasks Included: 87
Tasks Excluded: 532
Tasks using Sigma Time Spent fallback: 87

Total Worklog/Sigma Hours (included tasks): 1250.50
Total Selected Month Task Hours: 986.25

Output:
Jira_task_avg_July.xlsx
```

---

## Important limitations

1. **Your current Tempo worklogs file has no Issue Key**, so per-task total hours use **Σ Time Spent** (lifetime). The average-daily method still allocates a **portion** of that total into the selected month via working-day ratios.
2. For **true** worklog-based totals per issue, export issue-level worklogs (`Issue Key`, `Started`, `Time Spent`, `Author`).
3. Leave calendar day numbers are applied to the **selected `--month`** year/month.
4. Never invent hours when both worklogs and Σ Time Spent are empty → month hours = 0.

---

## Files created

| File | Role |
|------|------|
| `avg_workingHrs_calcultor.py` | Calculator script |
| `AVG_WORKING_HRS_LOGIC.md` | This documentation |
