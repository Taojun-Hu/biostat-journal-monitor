# Biostatistics Journal Monitor

This repository checks ten leading statistics and biostatistics journals every
month, retrieves recently published articles from OpenAlex, screens them for
biostatistical relevance, and stores DOI-based reports.

## Journals

1. Journal of the American Statistical Association
2. Biometrika
3. Journal of the Royal Statistical Society: Series B
4. The Annals of Statistics
5. Biometrics
6. Biostatistics
7. Statistics in Medicine
8. Statistical Methods in Medical Research
9. Journal of the Royal Statistical Society: Series C
10. The Annals of Applied Statistics

## Monthly behavior

The workflow runs at 08:15 Asia/Shanghai on the first day of each month. It
queries the preceding 45 days and reports only previously unseen DOIs. The
overlap protects against delayed indexing.

Each run produces:

- an Excel report with `Included`, `Needs_Review`, `Excluded`, and `Summary`
  worksheets;
- a Markdown summary;
- a cumulative Excel workbook;
- a persistent DOI history;
- a GitHub notification issue when new papers are found.

OpenAlex citation counts are included and are not Google Scholar counts.

## Manual run

Open `Actions`, select `Monthly biostatistics journal check`, and choose
`Run workflow`. The lookback period and optional end date can be overridden.

For local testing:

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python scripts/monthly_checker.py --lookback-days 45
```

## Screening

All research articles from Biometrics, Biostatistics, Statistics in Medicine,
and Statistical Methods in Medical Research are included. Papers from the six
broader statistics journals are:

- included when an explicit biomedical focus or a method central to
  biostatistics is identified;
- placed in `Needs_Review` when they concern adjacent statistical methodology
  without an explicit biomedical focus;
- otherwise recorded as excluded in the monthly screening log.

Screening is rule-based and auditable, but not a substitute for manual review.

