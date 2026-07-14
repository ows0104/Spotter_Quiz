# Neuro Spotter Bell Exam Practice

Static GitHub Pages build for the Neuro Spotter Bell Exam mock practice sets.

## Contents

- `index.html`: exam selection page
- `exams/01` through `exams/10`: bundled mock exam web apps

## GitHub Pages

Publish this folder from the root of a GitHub repository, or copy these files
into the root of an existing Pages repository.

## Exam trend weighting

The 1st, 2nd, and 3rd anatomy sections use the integrated anatomy frequency
workbook to set each question's sampling weight to `sqrt(n + 1)`, where `n` is
the historical exam frequency. A 30-question set allocates region quotas by
`sqrt(unique structures in region)`, then samples within each region by the exam
weight. The same anatomical structure appears at most once per set.

Refresh the metadata after updating the workbook:

```powershell
python update_exam_trends.py "C:\path\to\인체해부학_차시별_구조물_기출빈도.xlsx"
```
