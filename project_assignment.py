#!/usr/bin/python3
import openpyxl
import numpy as np
from scipy.optimize import linear_sum_assignment

wb = openpyxl.load_workbook(filename="test.xlsx")
ws = wb.active
last_worker_line = [x[2].value for x in ws.rows].index(None)
quota_sum = sum(row[0].value for row in ws.iter_rows(min_row=2, max_row=last_worker_line, min_col=2, max_col=2))
keyword_names = [x.value for x in ws[1][2:]]
W = np.zeros((quota_sum, len(keyword_names)))
i = 0
repeated_worker_names = []
for row in ws[2:last_worker_line]:
    quota = row[1].value
    W[i:i+quota] = [c.value for c in row[2:]]
    i += quota
    repeated_worker_names.extend([row[0].value]*quota)
num_projects = ws.max_row - last_worker_line - 1
P = np.array([[x.value for x in row[2:]] for row in ws.iter_rows(min_row=last_worker_line+2)], dtype=float)
project_names = [row[0].value for row in ws.iter_rows(min_row=last_worker_line+2)]
workers = {row[0].value: row[1].value for row in ws[2:last_worker_line]}
W /= np.linalg.norm(W, axis=1, keepdims=True)
P /= np.linalg.norm(P, axis=1, keepdims=True)
C = 1 - W @ P.transpose()
worker_ind, project_ind = linear_sum_assignment(C)
for w, p in zip(worker_ind, project_ind):
    print("{}: {} (score: {:.4f})".format(repeated_worker_names[w], project_names[p], 1 - C[w,p]))
