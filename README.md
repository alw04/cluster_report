# cluster_report

Generate an Excel report for Ceres or Atlas with CPU utilization, job statistics, filesystem usage, and project storage. Atlas reports also include GPU job statistics.

## Requirements

- Python 3.9 or newer with `XlsxWriter` installed.
- Run on the cluster you want to report on. The script detects the cluster automatically.
- On Ceres, you need `sudo` permissions to run `sacct` and SSH access to the LTS host.

## Usage

Pass the start and end dates in `YYYY-MM-DD` format. The start date is inclusive and the end date is exclusive.

For example, this command generates a report for October 2026.

```bash
python3 cluster_report.py --start 2026-10-01 --end 2026-11-01
```

The default filename is `<cluster>-report-<start>-to-<end>.xlsx`.
Use `--output FILE.xlsx` to choose a different filename.

Storage figures reflect the snapshot available when the script runs.
On Atlas, the GPU job section reports the number of jobs that requested GPUs and their allocated CPU-hours.
