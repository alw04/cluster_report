#!/usr/bin/env python3

import argparse
from datetime import date
from pathlib import Path
from subprocess import run

import xlsxwriter


def detect_cluster():
    result = run(
        ["scontrol", "show", "config"],
        capture_output=True,
        text=True,
        check=True,
    )

    for line in result.stdout.splitlines():
        key, separator, value = line.partition("=")
        if separator and key.strip() == "ClusterName":
            return value.strip().lower()

    raise RuntimeError("Could not find ClusterName in Slurm's configuration.")


def collect_utilization(start_date, end_date):
    fields = [
        "Allocated",
        "Down",
        "PlannedDown",
        "Idle",
        "Planned",
        "Reported",
    ]

    result = run(
        [
            "sreport",
            "cluster",
            "utilization",
            "-t",
            "hour",
            f"start={start_date}",
            f"end={end_date}",
            f"format={','.join(fields)}",
            "-P",
            "-n",
        ],
        capture_output=True,
        text=True,
        check=True,
    )

    rows = [line for line in result.stdout.splitlines() if line.strip()]
    if len(rows) != 1:
        raise ValueError(f"Expected one sreport utilization row, got {len(rows)}.")

    values = rows[0].split("|")
    if len(values) != len(fields):
        raise ValueError(
            f"Expected {len(fields)} sreport fields, got {len(values)}: {rows[0]!r}"
        )

    return dict(zip(fields, map(int, values)))


def collect_jobs(start_date, end_date, *, cluster, use_sudo=False):
    fields = ["Account", "CPUTimeRAW"]
    command = ["sudo", "sacct"] if use_sudo else ["sacct"]

    if cluster == "atlas":
        command.append("-a")
        fields.append("ReqTRES")

    command.extend(
        [
            "--truncate",
            f"--start={start_date}",
            f"--end={end_date}",
            f"--format={','.join(fields)}",
            "-P",
            "-X",
            "-n",
        ]
    )

    result = run(command, capture_output=True, text=True, check=True)

    jobs = []
    for line in result.stdout.splitlines():
        if not line.strip():
            continue

        values = line.split("|")
        if len(values) != len(fields):
            raise ValueError(f"Unexpected sacct row: {line!r}")

        job = dict(zip(fields, values))
        job["CPUTimeRAW"] = int(job["CPUTimeRAW"])
        jobs.append(job)

    return jobs


def collect_filesystem_usage(*, include_lts):
    fields = "--output=target,size,used,avail,pcent"
    commands = [
        ["df", "-B1", fields, "/project", "/90daydata"],
    ]

    if include_lts:
        commands.append(
            [
                "ssh",
                "-o",
                "BatchMode=yes",
                "-o",
                "ConnectTimeout=15",
                "nal-dtn-0.scinet.usda.gov",
                f"df -B1 {fields} /LTS/project",
            ]
        )

    filesystems = []

    for command in commands:
        result = run(
            command,
            capture_output=True,
            text=True,
            check=True,
        )

        for line in result.stdout.splitlines()[1:]:
            mount, size, used, available, percent = line.split()

            filesystems.append(
                {
                    "Mount": mount,
                    "SizeBytes": int(size),
                    "UsedBytes": int(used),
                    "AvailableBytes": int(available),
                    "UsedPercent": int(percent.rstrip("%")),
                }
            )

    return filesystems


def collect_ceres_project_storage():
    text = Path("/system/vast_quotas/view_vast_quotas").read_text()

    projects = []
    fields = None
    required_fields = {
        "Path",
        "Pretty-state",
        "Used-effective-capacity",
        "Hard-limit",
    }

    for line_number, line in enumerate(text.splitlines(), start=1):
        line = line.strip()
        if not line.startswith("|"):
            continue

        values = [value.strip() for value in line[1:].removesuffix("|").split("|")]

        if values[0] == "Id":
            fields = values
            missing_fields = required_fields.difference(fields)
            if missing_fields:
                missing_names = ", ".join(sorted(missing_fields))
                raise ValueError(
                    f"Ceres quota header is missing fields: {missing_names}"
                )
            continue

        if fields is None:
            continue

        if len(values) != len(fields):
            raise ValueError(
                f"Unexpected Ceres quota row on line {line_number}: "
                f"expected {len(fields)} fields, got {len(values)}."
            )

        row = dict(zip(fields, values))
        path = row["Path"]

        if not path.startswith("/project/"):
            continue

        usage = row["Used-effective-capacity"]
        quota = row["Hard-limit"]

        projects.append(
            {
                "Project": path.removeprefix("/project/"),
                "State": row["Pretty-state"],
                "UsedBytes": None if usage == "None" else int(usage),
                "QuotaBytes": None if quota == "None" else int(quota),
            }
        )

    if fields is None:
        raise ValueError(
            "Could not find the Ceres quota table header "
            "(expected 'Id' as the first field.)"
        )

    if not projects:
        raise ValueError("No /project/ rows found in the Ceres quota report.")

    return projects


def collect_atlas_project_storage():
    result = run(
        ["my_quota", "-r"],
        capture_output=True,
        text=True,
        check=True,
    )

    projects = []
    found_header = False

    for line in result.stdout.splitlines():
        line = line.strip()

        if line.startswith("Directory/Group"):
            found_header = True
            continue

        if not found_header or not line or set(line) == {"-"}:
            continue

        fields = line.split()
        if len(fields) != 4:
            raise ValueError(f"Unexpected Atlas quota row: {line!r}")

        group, usage, quota, percent = fields
        projects.append(
            {
                "Project": group,
                "UsedGB": usage if usage == "<1" else int(usage),
                "QuotaGB": quota if quota == "<1" else int(quota),
                "UsedPercent": int(percent),
            }
        )

    if not found_header:
        raise ValueError(
            "Could not find the Atlas quota table header "
            "(expected it to start with 'Directory/Group')."
        )

    if not projects:
        raise ValueError("No project rows found in the Atlas quota report.")

    return projects


def summarize_jobs(jobs):
    job_count = len(jobs)
    total_seconds = sum(job["CPUTimeRAW"] for job in jobs)

    usage_by_account = {}

    for job in jobs:
        account = job["Account"]
        usage_by_account[account] = usage_by_account.get(account, 0) + job["CPUTimeRAW"]

    accounts = [
        {
            "Account": account,
            "CPUHours": seconds / 3600,
        }
        for account, seconds in usage_by_account.items()
    ]

    accounts.sort(key=lambda account: account["CPUHours"], reverse=True)

    return {
        "JobCount": job_count,
        "TotalCPUHours": total_seconds / 3600,
        "AverageCPUHours": (total_seconds / job_count / 3600 if job_count else None),
        "MaxCPUHours": (
            max(job["CPUTimeRAW"] for job in jobs) / 3600 if job_count else None
        ),
        "UsageByAccount": accounts,
        "Top10Accounts": accounts[:10],
    }


def summarize_project_storage(projects, *, cluster):
    storage = []

    if cluster == "atlas":
        unit = "GB"
        for project in projects:
            storage.append(
                {
                    "Project": project["Project"],
                    "Used": project["UsedGB"],
                    "Quota": project["QuotaGB"],
                }
            )
    elif cluster == "ceres":
        unit = "TiB"
        for project in projects:
            if (
                project["Project"] == "rm_projects"
                or project["State"] == "FAILED"
                or project["UsedBytes"] is None
            ):
                continue

            quota = project["QuotaBytes"]
            storage.append(
                {
                    "Project": project["Project"],
                    "Used": project["UsedBytes"] / 1024**4,
                    "Quota": quota / 1024**4 if quota is not None else None,
                }
            )
    else:
        raise ValueError(f"Unsupported cluster: {cluster!r}")

    def usage_sort_key(project):
        used = project["Used"]
        return 0.5 if used == "<1" else used

    storage.sort(key=usage_sort_key, reverse=True)

    return {
        "Unit": unit,
        "UsageByProject": storage,
        "Top10Projects": storage[:10],
    }


def summarize_utilization(utilization):
    reported = utilization["Reported"]

    return {
        "ReportedCPUHours": reported,
        "AllocatedCPUHours": utilization["Allocated"],
        "Utilization": (utilization["Allocated"] / reported if reported else None),
    }


def summarize_filesystem_usage(filesystems):
    return [
        {
            "Mount": filesystem["Mount"],
            "SizeTiB": filesystem["SizeBytes"] / 1024**4,
            "UsedTiB": filesystem["UsedBytes"] / 1024**4,
            "AvailableTiB": filesystem["AvailableBytes"] / 1024**4,
            "UsedFraction": filesystem["UsedPercent"] / 100,
        }
        for filesystem in filesystems
    ]


def summarize_gpu_jobs(jobs):
    job_count = 0
    cpu_seconds = 0

    for job in jobs:
        for resource in job["ReqTRES"].split(","):
            name, separator, count = resource.strip().partition("=")

            if separator and (name == "gres/gpu" or name.startswith("gres/gpu:")):
                if int(count) > 0:
                    job_count += 1
                    cpu_seconds += job["CPUTimeRAW"]
                    break

    return {
        "JobCount": job_count,
        "CPUHours": cpu_seconds / 3600,
    }


def write_excel_report(summary, start_date, end_date, output_path):
    with xlsxwriter.Workbook(output_path) as workbook:
        sheet = workbook.add_worksheet("Cluster Report")

        project_storage = summary["project_storage"]
        storage_unit = project_storage["Unit"]
        used_header = f"Used ({storage_unit})"
        quota_header = f"Quota ({storage_unit})"

        heading = workbook.add_format({"bold": True})
        integer = workbook.add_format({"num_format": "#,##0"})
        decimal = workbook.add_format({"num_format": "#,##0.00"})
        percent = workbook.add_format({"num_format": "0.00%"})

        sheet.set_column("A:A", 38)
        sheet.set_column("B:C", 22)

        sheet.write("A1", "Cluster Usage Report", heading)
        sheet.write("A2", f"Period: {start_date} to {end_date} (end exclusive)")

        sheet.write_row("A4", ["CPU Utilization", "Period", "YTD"], heading)

        metrics = [
            ("Reported CPU Hours", "ReportedCPUHours", integer),
            ("Allocated CPU Hours", "AllocatedCPUHours", integer),
            ("Utilization", "Utilization", percent),
        ]

        for row, (label, key, cell_format) in enumerate(metrics, start=4):
            sheet.write(row, 0, label)

            for column, section in enumerate(
                ["utilization", "utilization_ytd"], start=1
            ):
                value = summary[section][key]
                sheet.write(
                    row,
                    column,
                    value if value is not None else "N/A",
                    cell_format,
                )

        sheet.write_row("A9", ["Job Statistics", "Period"], heading)

        metrics = [
            ("Total Jobs", "JobCount", integer),
            ("Total Allocated CPU Hours", "TotalCPUHours", decimal),
            ("Average Allocated CPU Hours per Job", "AverageCPUHours", decimal),
            ("Maximum Allocated CPU Hours per Job", "MaxCPUHours", decimal),
        ]

        for row, (label, key, cell_format) in enumerate(metrics, start=9):
            sheet.write(row, 0, label)
            value = summary["jobs"][key]
            sheet.write(
                row,
                1,
                value if value is not None else "N/A",
                cell_format,
            )

        row = 14

        if "gpu_jobs" in summary:
            gpu = summary["gpu_jobs"]
            sheet.write_row(row, 0, ["GPU-Requesting Jobs", "Period"], heading)
            sheet.write(row + 1, 0, "Job Count")
            sheet.write_number(row + 1, 1, gpu["JobCount"], integer)
            sheet.write(row + 2, 0, "Allocated CPU Hours")
            sheet.write_number(row + 2, 1, gpu["CPUHours"], decimal)
            row += 4

        sheet.set_column("D:E", 22)
        sheet.write_row(
            row,
            0,
            ["Filesystem", "Size (TiB)", "Used (TiB)", "Available (TiB)", "Used (%)"],
            heading,
        )

        for fs_row, filesystem in enumerate(summary["filesystems"], start=row + 1):
            sheet.write(fs_row, 0, filesystem["Mount"])
            sheet.write(fs_row, 1, filesystem["SizeTiB"], decimal)
            sheet.write(fs_row, 2, filesystem["UsedTiB"], decimal)
            sheet.write(fs_row, 3, filesystem["AvailableTiB"], decimal)
            sheet.write(fs_row, 4, filesystem["UsedFraction"], percent)

        row += len(summary["filesystems"]) + 2

        sheet.write_row(
            row,
            0,
            ["Top 10 Compute Accounts", "Allocated CPU Hours"],
            heading,
        )
        row += 1

        for account in summary["jobs"]["Top10Accounts"]:
            sheet.write(row, 0, account["Account"])
            sheet.write(row, 1, account["CPUHours"], decimal)
            row += 1

        row += 1

        sheet.write_row(
            row,
            0,
            ["Top 10 Storage Projects", used_header, quota_header],
            heading,
        )
        row += 1

        for project in project_storage["Top10Projects"]:
            sheet.write_string(row, 0, project["Project"])
            sheet.write(row, 1, project["Used"], decimal)
            quota = project["Quota"]
            sheet.write(row, 2, quota if quota is not None else "N/A", decimal)
            row += 1

        compute_sheet = workbook.add_worksheet("Compute Usage")
        compute_sheet.set_column("A:A", 50)
        compute_sheet.set_column("B:B", 24)
        compute_sheet.freeze_panes(1, 0)

        compute_sheet.write_row(
            "A1",
            ["Account", "Allocated CPU Hours"],
            heading,
        )

        accounts = summary["jobs"]["UsageByAccount"]

        for row, account in enumerate(accounts, start=1):
            compute_sheet.write_string(row, 0, account["Account"])
            compute_sheet.write(row, 1, account["CPUHours"], decimal)

        if accounts:
            compute_sheet.autofilter(0, 0, len(accounts), 1)

        storage_sheet = workbook.add_worksheet("Project Storage")
        storage_sheet.set_column("A:A", 50)
        storage_sheet.set_column("B:C", 22)
        storage_sheet.freeze_panes(1, 0)

        storage_sheet.write_row(
            "A1",
            ["Project", used_header, quota_header],
            heading,
        )

        projects = project_storage["UsageByProject"]

        for row, project in enumerate(projects, start=1):
            storage_sheet.write_string(row, 0, project["Project"])
            storage_sheet.write(row, 1, project["Used"], decimal)
            quota = project["Quota"]
            storage_sheet.write(row, 2, quota if quota is not None else "N/A", decimal)

        if projects:
            storage_sheet.autofilter(0, 0, len(projects), 2)

        compute_count = min(10, len(accounts))

        if compute_count:
            chart = workbook.add_chart({"type": "bar"})

            chart.add_series(
                {
                    "name": "Allocated CPU Hours",
                    "categories": ["Compute Usage", 1, 0, compute_count, 0],
                    "values": ["Compute Usage", 1, 1, compute_count, 1],
                }
            )

            chart.set_title({"name": "Top 10 Compute Accounts"})
            chart.set_x_axis({"name": "Allocated CPU Hours", "min": 0})
            chart.set_y_axis({"reverse": True})
            chart.set_legend({"none": True})
            chart.set_size({"width": 800, "height": 500})

            sheet.insert_chart("G4", chart)

        storage_count = min(10, len(projects))

        if storage_count:
            chart = workbook.add_chart({"type": "bar"})

            for column, name in [(1, "Used"), (2, "Quota")]:
                chart.add_series(
                    {
                        "name": name,
                        "categories": ["Project Storage", 1, 0, storage_count, 0],
                        "values": ["Project Storage", 1, column, storage_count, column],
                    }
                )

            chart.set_title({"name": "Top 10 Projects by Storage Used"})
            chart.set_x_axis({"name": storage_unit, "min": 0})
            chart.set_y_axis({"reverse": True})
            chart.set_legend({"position": "bottom"})
            chart.set_size({"width": 800, "height": 500})

            sheet.insert_chart("G31", chart)

        pie_configs = [
            (
                compute_sheet,
                summary["jobs"]["UsageByAccount"],
                "Account",
                "CPUHours",
                "Allocated CPU Hours",
                "Compute Usage Share",
                "T4",
            ),
            (
                storage_sheet,
                projects,
                "Project",
                "Used",
                used_header,
                (
                    "Project Storage Share (excludes usage <1 GB)"
                    if storage_unit == "GB"
                    else "Project Storage Usage Share"
                ),
                "T31",
            ),
        ]

        for (
            source_sheet,
            records,
            name_key,
            value_key,
            unit,
            title,
            position,
        ) in pie_configs:
            records = [
                record
                for record in records
                if isinstance(record[value_key], (int, float))
            ]

            slices = [(record[name_key], record[value_key]) for record in records[:10]]

            other = sum(record[value_key] for record in records[10:])
            if other > 0:
                slices.append(("Other", other))

            if not slices or sum(value for _, value in slices) == 0:
                continue

            source_sheet.set_column("E:E", 50)
            source_sheet.set_column("F:F", 24)
            source_sheet.write_row("E1", ["Usage Share", unit], heading)

            for row, (name, value) in enumerate(slices, start=1):
                source_sheet.write_string(row, 4, name)
                source_sheet.write(row, 5, value, decimal)

            chart = workbook.add_chart({"type": "pie"})

            chart.add_series(
                {
                    "name": title,
                    "categories": [source_sheet.name, 1, 4, len(slices), 4],
                    "values": [source_sheet.name, 1, 5, len(slices), 5],
                    "data_labels": {
                        "percentage": True,
                        "num_format": "0.0%",
                    },
                }
            )

            chart.set_title({"name": title})
            chart.set_legend({"position": "right"})
            chart.set_size({"width": 800, "height": 500})

            sheet.insert_chart(position, chart)


def parse_date(value):
    try:
        return date.fromisoformat(value)
    except ValueError:
        raise argparse.ArgumentTypeError(
            f"invalid date {value!r}; use YYYY-MM-DD, for example 2026-10-01"
        ) from None


def parse_args():
    parser = argparse.ArgumentParser(
        description="Generate a cluster usage report.",
        epilog=(
            "Example: %(prog)s --start 2026-10-01 --end 2026-11-01\n"
            "The end date is exclusive: this example reports October."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--start",
        type=parse_date,
        required=True,
        metavar="YYYY-MM-DD",
        help="Start date, inclusive (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--end",
        type=parse_date,
        required=True,
        metavar="YYYY-MM-DD",
        help="End date, exclusive (YYYY-MM-DD)",
    )
    parser.add_argument(
        "--output",
        metavar="FILE.xlsx",
        help="Output Excel filename "
        "(defaults to <cluster>-report-<start>-to-<end>.xlsx)",
    )

    args = parser.parse_args()

    if args.end <= args.start:
        parser.error("--end must be after --start")

    return args


def main():
    args = parse_args()
    cluster = detect_cluster()

    if cluster == "ceres":
        collect_project_storage = collect_ceres_project_storage
    elif cluster == "atlas":
        collect_project_storage = collect_atlas_project_storage
    else:
        raise SystemExit(f"Unsupported Slurm cluster: {cluster!r}")

    jobs = collect_jobs(
        args.start,
        args.end,
        cluster=cluster,
        use_sudo=(cluster == "ceres"),
    )

    summary = {
        "utilization": summarize_utilization(collect_utilization(args.start, args.end)),
        "utilization_ytd": summarize_utilization(
            collect_utilization(args.start.replace(month=1, day=1), args.end)
        ),
        "jobs": summarize_jobs(jobs),
        "filesystems": summarize_filesystem_usage(
            collect_filesystem_usage(include_lts=(cluster == "ceres"))
        ),
        "project_storage": summarize_project_storage(
            collect_project_storage(), cluster=cluster
        ),
    }

    if cluster == "atlas":
        summary["gpu_jobs"] = summarize_gpu_jobs(jobs)

    output_path = args.output or f"{cluster}-report-{args.start}-to-{args.end}.xlsx"

    write_excel_report(summary, args.start, args.end, output_path)
    print(f"Report saved to {output_path}")


if __name__ == "__main__":
    main()
