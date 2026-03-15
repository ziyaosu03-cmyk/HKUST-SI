# -*- coding: utf-8 -*-
"""
Modifies CBP_P3B.mpp per ISOM5460 Assignment 2 requirements:
  1. Saves current schedule as Baseline 0
  2. Inserts actual progress data (status date: Jan 1, 2011)
  3. Sets StatusDate to 2011-01-01

Output: CBP_P3B_with_actuals.xml (Microsoft Project XML / MSPDI format)
  -> Can be opened directly in MS Project and saved as .mpp
     (mpxj supports reading MPP but only writing MSPDI XML, MPX, etc.)

Actual data source: ISOM5460+Assignment+2.pdf, table on page 1
"""

import jpype
import jpype.imports
import mpxj
import glob
import os

jar_dir = mpxj.mpxj_dir
jars = glob.glob(os.path.join(jar_dir, "*.jar"))
classpath = ":".join(jars)
jpype.startJVM(jpype.getDefaultJVMPath(), f"-Djava.class.path={classpath}")

from org.mpxj.reader import UniversalProjectReader
from org.mpxj.mspdi import MSPDIWriter
from org.mpxj import Duration, TimeUnit
from java.time import LocalDateTime
from java.io import FileOutputStream
from java.math import BigDecimal

# ---- Read source file --------------------------------------------------------
reader = UniversalProjectReader()
project = reader.read("/workspace/CBP_P3B.mpp")
props = project.getProjectProperties()
print(f"Loaded: {props.getStartDate()}  ->  {props.getFinishDate()}")

# ---- Step 1: Save current schedule as Baseline 0 ----------------------------
print("\nSaving Baseline 0...")
for task in project.getTasks():
    if not task.getName():
        continue
    task.setBaselineStart(task.getStart())
    task.setBaselineFinish(task.getFinish())
    task.setBaselineDuration(task.getDuration())
    task.setBaselineCost(task.getCost())
    task.setBaselineWork(task.getWork())

# ---- Step 2: Set status date ------------------------------------------------
STATUS = LocalDateTime.of(2011, 1, 1, 8, 0)
props.setStatusDate(STATUS)
print(f"Status date set to: {STATUS}")

# ---- Step 3: Insert actuals from assignment PDF table -----------------------
# task_id: (actual_start, actual_finish or None, actual_dur_days, remaining_days)
ACTUALS = {
    3:  (LocalDateTime.of(2010,  2,  8, 8, 0), LocalDateTime.of(2010,  4,  9, 17, 0),  45.0,   0.0),
    4:  (LocalDateTime.of(2010,  4, 12, 8, 0), LocalDateTime.of(2010,  7, 15, 17, 0),  67.0,   0.0),
    5:  (LocalDateTime.of(2010,  7, 16, 8, 0), LocalDateTime.of(2010,  8, 23, 17, 0),  27.0,   0.0),
    6:  (LocalDateTime.of(2010, 11,  3, 8, 0), None,                                   34.0,  44.0),
    10: (LocalDateTime.of(2010,  2,  8, 8, 0), LocalDateTime.of(2010,  3, 12, 17, 0),  25.0,   0.0),
    12: (LocalDateTime.of(2010,  3, 15, 8, 0), LocalDateTime.of(2010,  8, 17, 17, 0), 110.0,   0.0),
    13: (LocalDateTime.of(2010, 11,  3, 8, 0), None,                                   34.0, 119.0),
    14: (LocalDateTime.of(2010,  3, 15, 8, 0), LocalDateTime.of(2010,  7, 30, 17, 0),  98.0,   0.0),
    15: (LocalDateTime.of(2010,  4,  5, 8, 0), LocalDateTime.of(2010,  5,  7, 17, 0),  25.0,   0.0),
    18: (LocalDateTime.of(2010,  3,  8, 8, 0), LocalDateTime.of(2010,  3, 29, 17, 0),  16.0,   0.0),
    19: (LocalDateTime.of(2010,  3, 30, 8, 0), LocalDateTime.of(2010,  6, 29, 17, 0),  65.0,   0.0),
    20: (LocalDateTime.of(2010,  3, 30, 8, 0), LocalDateTime.of(2010,  8, 11, 17, 0),  95.0,   0.0),
    21: (LocalDateTime.of(2010,  5,  4, 8, 0), LocalDateTime.of(2010,  6,  3, 17, 0),  22.0,   0.0),
    24: (LocalDateTime.of(2010,  1,  4, 8, 0), LocalDateTime.of(2010,  2,  5, 17, 0),  25.0,   0.0),
    25: (LocalDateTime.of(2010,  8, 24, 8, 0), LocalDateTime.of(2010, 11,  9, 17, 0),  55.0,   0.0),
}

print("\nInserting actual progress:")
for task in project.getTasks():
    if not task.getName():
        continue
    tid = int(str(task.getID()))
    if tid not in ACTUALS:
        continue
    act_start, act_finish, act_d, rem_d = ACTUALS[tid]
    task.setActualStart(act_start)
    task.setActualDuration(Duration.getInstance(act_d, TimeUnit.DAYS))
    task.setRemainingDuration(Duration.getInstance(rem_d, TimeUnit.DAYS))
    if act_finish is not None:
        task.setActualFinish(act_finish)
        task.setPercentageComplete(BigDecimal(100))
    else:
        total = act_d + rem_d
        pct = (act_d / total * 100.0) if total > 0 else 0.0
        task.setPercentageComplete(BigDecimal(round(pct, 1)))
    status = "COMPLETE" if act_finish else f"IN PROGRESS ({act_d:.0f}d done / {rem_d:.0f}d rem)"
    print(f"  Task {tid:2d} {str(task.getName()):<42} [{status}]")

# ---- Step 4: Write MSPDI XML ------------------------------------------------
OUTPUT = "/workspace/CBP_P3B_with_actuals.xml"
writer = MSPDIWriter()
fos = FileOutputStream(OUTPUT)
writer.write(project, fos)
fos.close()

print(f"\nWritten: {OUTPUT}  ({os.path.getsize(OUTPUT):,} bytes)")
print("\nVerification:")

import xml.etree.ElementTree as ET
NS = "http://schemas.microsoft.com/project"

tree = ET.parse(OUTPUT)
root = tree.getroot()

sd_el = root.find(f"{{{NS}}}StatusDate")
print(f"  StatusDate : {sd_el.text if sd_el is not None else 'MISSING'}")

checks = {"Hardware specifications": None, "Serial I/O drivers": None}
for task_el in root.findall(f"{{{NS}}}Tasks/{{{NS}}}Task"):
    name_el = task_el.find(f"{{{NS}}}Name")
    if name_el is None or name_el.text not in checks:
        continue
    n = name_el.text
    bl = task_el.find(f"{{{NS}}}Baseline")
    bl_start = bl.find(f"{{{NS}}}Start").text if bl is not None and bl.find(f"{{{NS}}}Start") is not None else None
    ast_el = task_el.find(f"{{{NS}}}ActualStart")
    adur_el = task_el.find(f"{{{NS}}}ActualDuration")
    rem_el = task_el.find(f"{{{NS}}}RemainingDuration")
    pct_el = task_el.find(f"{{{NS}}}PercentComplete")
    print(f"  {n}:")
    print(f"    BaselineStart={bl_start}  ActualStart={ast_el.text if ast_el is not None else None}")
    print(f"    ActualDuration={adur_el.text if adur_el is not None else None}  "
          f"RemainingDuration={rem_el.text if rem_el is not None else None}  "
          f"PercentComplete={pct_el.text if pct_el is not None else None}")
