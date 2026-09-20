#!/usr/bin/env python3
"""Generate the two VP daily-report narratives from an XLSX workbook."""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import math
import posixpath
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Iterable


SHEET_TREND = "销售过程数据走势"
SHEET_ONLINE = "线上官渠潜客客流"
SHEET_DAILY = "各车系当日数据"
SHEET_MONTH = "月累环比 (2)"
REQUIRED_SHEETS = (SHEET_TREND, SHEET_ONLINE, SHEET_DAILY, SHEET_MONTH)

NS = {
    "x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
}


class ReportDataError(RuntimeError):
    """Raised when the workbook cannot support a reliable report."""


@dataclass(frozen=True)
class CompareConfig:
    rule: str
    label: str
    current_dates: tuple[dt.date, ...]
    baseline_dates: tuple[dt.date, ...]
    baseline_mode: str


@dataclass(frozen=True)
class Comparison:
    current: float | None
    baseline: float | None
    rate: float | None
    current_found: int
    baseline_found: int
    current_expected: int
    baseline_expected: int

    @property
    def baseline_partial(self) -> bool:
        return 0 < self.baseline_found < self.baseline_expected


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="从 VP 日报 Excel 生成概览解析和车系数据解析文案。"
    )
    parser.add_argument("workbook", type=Path, help="输入的 VP日报_YYYYMMDD.xlsx")
    parser.add_argument(
        "run_date",
        nargs="?",
        type=dt.date.fromisoformat,
        help="运行日期 YYYY-MM-DD；目标日期为运行日期前一天。默认从文件名推断。",
    )
    parser.add_argument(
        "output",
        nargs="?",
        type=Path,
        default=Path("outputs/vp_daily_fields.csv"),
        help="输出 CSV 路径。",
    )
    parser.add_argument(
        "--target-date",
        type=dt.date.fromisoformat,
        help="直接指定目标日期 YYYY-MM-DD；不能与 run_date 同时使用。",
    )
    parser.add_argument(
        "--stdout-only",
        action="store_true",
        help="仅在终端显示文案，不写入 CSV 文件。",
    )
    args = parser.parse_args()
    if args.run_date and args.target_date:
        parser.error("run_date 和 --target-date 只能指定一个")
    return args


def infer_target_date(workbook: Path, run_date: dt.date | None, target_date: dt.date | None) -> dt.date:
    if target_date:
        return target_date
    if run_date:
        return run_date - dt.timedelta(days=1)
    match = re.search(r"VP日报_(\d{8})", workbook.stem)
    if not match:
        raise ReportDataError("无法从文件名推断日期，请传入 run_date 或 --target-date。")
    inferred_run_date = dt.datetime.strptime(match.group(1), "%Y%m%d").date()
    return inferred_run_date - dt.timedelta(days=1)


def compare_config(target: dt.date) -> CompareConfig:
    weekday = target.weekday()
    if weekday == 0:
        previous_monday = target - dt.timedelta(days=7)
        return CompareConfig(
            "prev_week_weekday_avg",
            "上一周周中日均",
            (target,),
            tuple(previous_monday + dt.timedelta(days=i) for i in range(5)),
            "average",
        )
    if weekday in (1, 2, 3, 4):
        return CompareConfig(
            "previous_day",
            "前一日",
            (target,),
            (target - dt.timedelta(days=1),),
            "sum",
        )
    if weekday == 5:
        previous_saturday = target - dt.timedelta(days=7)
        return CompareConfig(
            "prev_weekend_avg",
            "上周末日均",
            (target,),
            (previous_saturday, previous_saturday + dt.timedelta(days=1)),
            "average",
        )
    current_monday = target - dt.timedelta(days=6)
    previous_monday = current_monday - dt.timedelta(days=7)
    return CompareConfig(
        "weekly_vs_last_week",
        "上周",
        tuple(current_monday + dt.timedelta(days=i) for i in range(7)),
        tuple(previous_monday + dt.timedelta(days=i) for i in range(7)),
        "sum",
    )


def col_to_num(ref: str) -> int:
    letters = re.match(r"[A-Z]+", ref)
    if not letters:
        raise ReportDataError(f"无效单元格坐标：{ref}")
    result = 0
    for char in letters.group(0):
        result = result * 26 + ord(char) - 64
    return result


def cell_pos(ref: str) -> tuple[int, int]:
    row = re.search(r"\d+", ref)
    if not row:
        raise ReportDataError(f"无效单元格坐标：{ref}")
    return int(row.group(0)), col_to_num(ref)


def as_date(value: object) -> dt.date | None:
    if isinstance(value, (int, float)):
        text = str(int(value))
        if len(text) == 8:
            try:
                return dt.datetime.strptime(text, "%Y%m%d").date()
            except ValueError:
                pass
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(value))
    if isinstance(value, str):
        text = value.strip()
        for pattern, length in (("%Y-%m-%d", 10), ("%Y%m%d", 8)):
            try:
                return dt.datetime.strptime(text[:length], pattern).date()
            except ValueError:
                continue
    return value if isinstance(value, dt.date) else None


def to_num(value: object) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_rate(current: float | None, baseline: float | None) -> float | None:
    if current is None or baseline in (None, 0):
        return None
    return current / baseline - 1


class XlsxReader:
    def __init__(self, workbook: Path) -> None:
        if not workbook.is_file():
            raise ReportDataError(f"找不到输入文件：{workbook}")
        try:
            with zipfile.ZipFile(workbook) as archive:
                shared = self._read_shared_strings(archive)
                paths = self._sheet_paths(archive)
                missing = [name for name in REQUIRED_SHEETS if name not in paths]
                if missing:
                    raise ReportDataError("缺少工作表：" + "、".join(missing))
                self.sheets = {
                    name: self._read_sheet(archive, paths[name], shared)
                    for name in REQUIRED_SHEETS
                }
        except zipfile.BadZipFile as exc:
            raise ReportDataError(f"不是有效的 XLSX 文件：{workbook}") from exc

    @staticmethod
    def _read_shared_strings(archive: zipfile.ZipFile) -> list[str]:
        if "xl/sharedStrings.xml" not in archive.namelist():
            return []
        root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
        return ["".join(node.text or "" for node in item.findall(".//x:t", NS)) for item in root.findall("x:si", NS)]

    @staticmethod
    def _sheet_paths(archive: zipfile.ZipFile) -> dict[str, str]:
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        targets = {}
        for relation in relationships.findall("rel:Relationship", NS):
            target = relation.attrib["Target"]
            if target.startswith("/"):
                path = target.lstrip("/")
            else:
                path = posixpath.normpath(posixpath.join("xl", target))
            targets[relation.attrib["Id"]] = path
        return {
            sheet.attrib["name"]: targets[sheet.attrib[f"{{{NS['r']}}}id"]]
            for sheet in workbook.findall("x:sheets/x:sheet", NS)
        }

    @staticmethod
    def _read_sheet(
        archive: zipfile.ZipFile, path: str, shared: list[str]
    ) -> dict[tuple[int, int], object]:
        root = ET.fromstring(archive.read(path))
        cells: dict[tuple[int, int], object] = {}
        for cell in root.findall(".//x:c", NS):
            kind = cell.attrib.get("t")
            value_node = cell.find("x:v", NS)
            if kind == "inlineStr":
                value: object = "".join(node.text or "" for node in cell.findall(".//x:t", NS))
            elif value_node is None:
                value = None
            else:
                raw = value_node.text
                if kind == "s":
                    value = shared[int(raw)]
                else:
                    try:
                        numeric = float(raw)
                        value = int(numeric) if numeric.is_integer() else numeric
                    except (TypeError, ValueError):
                        value = raw
            cells[cell_pos(cell.attrib["r"])] = value
        return cells


class ReportBuilder:
    def __init__(self, reader: XlsxReader, target: dt.date) -> None:
        self.target = target
        self.config = compare_config(target)
        self.trend = reader.sheets[SHEET_TREND]
        self.online = reader.sheets[SHEET_ONLINE]
        self.daily = reader.sheets[SHEET_DAILY]
        self.month = reader.sheets[SHEET_MONTH]
        self.warnings: list[str] = []
        self._max_trend_row = max(row for row, _ in self.trend)
        self._max_trend_col = max(col for _, col in self.trend)
        self._max_online_col = max(col for _, col in self.online)
        self._scope_rows = self._index_scope_rows()

    def _index_scope_rows(self) -> dict[str, int]:
        rows: dict[str, int] = {}
        for row in range(1, self._max_trend_row + 1):
            value = self.trend.get((row, 1))
            label = str(value).strip() if value is not None else ""
            if label in {"全系", "GX", "L03", "G9L"}:
                rows[label] = row
        missing = {"全系", "GX", "L03", "G9L"} - rows.keys()
        if missing:
            raise ReportDataError("销售过程数据走势缺少车系区块：" + "、".join(sorted(missing)))
        return rows

    def trend_value(self, scope: str, metric: str, date: dt.date) -> float | None:
        scope_row = self._scope_rows[scope]
        metric_row = None
        for row in range(scope_row + 2, min(scope_row + 12, self._max_trend_row + 1)):
            label = self.trend.get((row, 1))
            if label is not None and str(label).strip() == metric:
                metric_row = row
                break
        if metric_row is None:
            raise ReportDataError(f"{SHEET_TREND} 的 {scope} 区块缺少指标：{metric}")
        for col in range(2, self._max_trend_col + 1):
            if as_date(self.trend.get((scope_row + 1, col))) == date:
                return to_num(self.trend.get((metric_row, col)))
        return None

    def online_value(self, date: dt.date) -> float | None:
        for col in range(2, self._max_online_col + 1):
            if as_date(self.online.get((1, col))) == date:
                return to_num(self.online.get((2, col)))
        return None

    @staticmethod
    def _aggregate(values: Iterable[float | None], mode: str) -> tuple[float | None, int]:
        present = [value for value in values if value is not None]
        if not present:
            return None, 0
        total = sum(present)
        return (total / len(present) if mode == "average" else total), len(present)

    def compare_getter(
        self,
        getter: Callable[[dt.date], float | None],
        *,
        name: str,
        allow_missing_current: bool = False,
        allow_missing_baseline: bool = False,
    ) -> Comparison:
        current_values = [getter(date) for date in self.config.current_dates]
        baseline_values = [getter(date) for date in self.config.baseline_dates]
        current, current_found = self._aggregate(current_values, "sum")
        baseline, baseline_found = self._aggregate(baseline_values, self.config.baseline_mode)
        if not allow_missing_current and current_found != len(current_values):
            missing = [str(date) for date, value in zip(self.config.current_dates, current_values) if value is None]
            raise ReportDataError(f"{name} 缺少当前周期数据：{', '.join(missing)}")
        if not allow_missing_baseline and baseline_found != len(baseline_values):
            missing = [str(date) for date, value in zip(self.config.baseline_dates, baseline_values) if value is None]
            raise ReportDataError(f"{name} 缺少对比周期数据：{', '.join(missing)}")
        return Comparison(
            current,
            baseline,
            safe_rate(current, baseline),
            current_found,
            baseline_found,
            len(current_values),
            len(baseline_values),
        )

    def compare_trend(
        self,
        scope: str,
        metric: str,
        *,
        allow_missing_current: bool = False,
        allow_missing_baseline: bool = False,
    ) -> Comparison:
        return self.compare_getter(
            lambda date: self.trend_value(scope, metric, date),
            name=f"{scope}{metric}",
            allow_missing_current=allow_missing_current,
            allow_missing_baseline=allow_missing_baseline,
        )

    def compare_derived(self, metric: str, *, exclude_g9l: bool = False) -> Comparison:
        def getter(date: dt.date) -> float | None:
            all_value = self.trend_value("全系", metric, date)
            gx_value = self.trend_value("GX", metric, date)
            l03_value = self.trend_value("L03", metric, date)
            g9l_value = self.trend_value("G9L", metric, date)
            if all_value is None:
                return None
            if exclude_g9l:
                return all_value - (g9l_value or 0)
            if gx_value is None or l03_value is None:
                return None
            return all_value - (g9l_value or 0) - gx_value - l03_value

        name = f"非G9L{metric}" if exclude_g9l else f"其他车型{metric}"
        return self.compare_getter(getter, name=name)

    def has_prior_trend_value(self, scope: str, metric: str) -> bool:
        scope_row = self._scope_rows[scope]
        metric_row = None
        for row in range(scope_row + 2, min(scope_row + 12, self._max_trend_row + 1)):
            label = self.trend.get((row, 1))
            if label is not None and str(label).strip() == metric:
                metric_row = row
                break
        if metric_row is None:
            raise ReportDataError(f"{SHEET_TREND} 的 {scope} 区块缺少指标：{metric}")
        for col in range(2, self._max_trend_col + 1):
            date = as_date(self.trend.get((scope_row + 1, col)))
            if date is not None and date < self.target and to_num(self.trend.get((metric_row, col))) is not None:
                return True
        return False

    def build(self) -> tuple[str, str]:
        online = self.compare_getter(self.online_value, name="线上官渠潜客客流")
        traffic = self.compare_trend("全系", "客流")
        leads = self.compare_trend("全系", "线索")
        g9l_leads = self.compare_trend("G9L", "线索")
        non_g9l_leads = self.compare_derived("线索", exclude_g9l=True)
        gx_leads = self.compare_trend("GX", "线索")
        l03_leads = self.compare_trend("L03", "线索")
        other_leads = self.compare_derived("线索")

        drives = self.compare_trend("全系", "试驾")
        g9l_drives = self.compare_trend(
            "G9L", "试驾", allow_missing_current=True, allow_missing_baseline=True
        )
        l03_drives = self.compare_trend("L03", "试驾")
        gx_drives = self.compare_trend("GX", "试驾")
        other_drives = self.compare_derived("试驾")

        orders = self.compare_trend("全系", "锁单")
        g9l_orders = self.compare_trend(
            "G9L", "锁单", allow_missing_current=True, allow_missing_baseline=True
        )
        l03_orders = self.compare_trend("L03", "锁单")
        gx_orders = self.compare_trend("GX", "锁单")
        other_orders = self.compare_derived("锁单")

        g9l_leads_share = (
            g9l_leads.current / leads.current
            if g9l_leads.current is not None and leads.current
            else None
        )
        g9l_drive_share = (
            g9l_drives.current / drives.current
            if g9l_drives.current is not None and drives.current
            else None
        )
        g9l_order_share = (
            g9l_orders.current / orders.current
            if g9l_orders.current is not None and orders.current
            else None
        )
        g9l_drive_label = self.config.label
        if g9l_drives.baseline_partial and self.config.baseline_mode == "average":
            g9l_drive_label = self.config.label.replace("日均", "有值日均")
            self.warnings.append(
                f"G9L试驾对比周期仅 {g9l_drives.baseline_found}/{g9l_drives.baseline_expected} 天有值，按有值日均计算。"
            )
        if (
            g9l_orders.current is not None
            and g9l_orders.rate is None
            and not self.has_prior_trend_value("G9L", "锁单")
        ):
            g9l_order_vs = f"{self.target.month}/{self.target.day}首次有锁单数据，暂无历史周期可比"
        else:
            g9l_order_vs = fmt_vs(self.config.label, g9l_orders.rate)

        overview = (
            f"线上官渠潜客客流：官渠客流{trend_word(online.rate)}，VS {self.config.label}{fmt_pct(online.rate)}；\n\n"
            f"门店客流：门店客流{trend_word(traffic.rate)}，VS {self.config.label}{fmt_pct(traffic.rate)}；\n\n"
            f"线索总量：线索总量{trend_word(leads.rate)}，VS {self.config.label}{fmt_pct(leads.rate)}；"
            f"G9L 线索{fmt_num(g9l_leads.current)}（占线索总量比例{fmt_pct_value(g9l_leads_share)}），"
            f"VS {self.config.label}{fmt_pct(g9l_leads.rate)}；"
            f"全系（非 G9L）线索 VS {self.config.label}{fmt_pct(non_g9l_leads.rate)}；"
            f"其中 GX{fmt_pct(gx_leads.rate)}，L03{fmt_pct(l03_leads.rate)}，"
            f"其他车型合计{fmt_pct(other_leads.rate)}；\n\n"
            f"试驾总量：试驾总量{trend_word(drives.rate)}，VS {self.config.label}{fmt_pct(drives.rate)}；"
            f"G9L 试驾{fmt_num(g9l_drives.current)}（占试驾总量比例{fmt_pct_value(g9l_drive_share)}），"
            f"{fmt_vs(g9l_drive_label, g9l_drives.rate)}；"
            f"L03{fmt_pct(l03_drives.rate)}，GX{fmt_pct(gx_drives.rate)}，"
            f"其他车型合计{fmt_pct(other_drives.rate)}；\n\n"
            f"锁单总量：全系锁单总量{fmt_num(orders.current)}台，VS {self.config.label}{fmt_pct(orders.rate)}；"
            f"其中 G9L 净锁单{fmt_num(g9l_orders.current)}台（占全系锁单比例{fmt_pct_value(g9l_order_share)}），"
            f"{g9l_order_vs}；"
            f"L03 净锁单{fmt_num(l03_orders.current)}台，{fmt_pct(l03_orders.rate)}；"
            f"GX 净锁单{fmt_num(gx_orders.current)}台，{fmt_pct(gx_orders.rate)}；"
            f"其他车型合计{fmt_num(other_orders.current)}台，{fmt_pct(other_orders.rate)}。"
        )

        daily_trial = to_num(self.daily.get((4, 13)))
        target_trial = self.trend_value("全系", "试驾", self.target)
        if daily_trial is not None and target_trial is not None and round(daily_trial) != round(target_trial):
            self.warnings.append(
                f"试驾日值存在口径差异：{SHEET_DAILY}!M4={fmt_num(daily_trial)}，"
                f"{SHEET_TREND}全系试驾={fmt_num(target_trial)}。"
            )

        day = self.target.day
        month_orders = to_num(self.month.get((3, 20)))
        month_leads = to_num(self.daily.get((16, 4)))
        month_conversion = (
            month_orders / month_leads if month_orders is not None and month_leads else None
        )
        model = (
            f"客流：{day}日客流{fmt_num(to_num(self.daily.get((4, 2))))}，"
            f"月累客流{fmt_num(to_num(self.daily.get((16, 2))))}，"
            f"月累环比{fmt_pct(to_num(self.daily.get((20, 2))), False)}，"
            f"月累同比{fmt_pct(to_num(self.daily.get((65, 2))), False)}\n"
            f"线索：{day}日线索{fmt_num(to_num(self.daily.get((4, 4))))}，"
            f"月累线索{fmt_num(to_num(self.daily.get((16, 4))))}，"
            f"月累环比{fmt_pct(to_num(self.daily.get((20, 4))), False)}，"
            f"月累同比{fmt_pct(to_num(self.daily.get((65, 4))), False)}\n"
            f"试驾：{day}日试驾{fmt_num(daily_trial)}，"
            f"月累试驾{fmt_num(to_num(self.daily.get((16, 13))))}，"
            f"月累环比{fmt_pct(to_num(self.daily.get((20, 13))), False)}，"
            f"月累同比{fmt_pct(to_num(self.daily.get((65, 13))), False)}\n"
            f"锁单：{day}日锁单{fmt_num(self.trend_value('全系', '锁单', self.target))}，"
            f"月累锁单{fmt_num(month_orders)}，"
            f"月累环比{fmt_pct(to_num(self.month.get((3, 21))), False)}，"
            f"月累同比{fmt_pct(to_num(self.month.get((3, 22))), False)}\n"
            f"转化率：月累转化率{fmt_pct_value(month_conversion)}，"
            f"月累环比{fmt_pct(to_num(self.month.get((3, 24))), False)}，"
            f"月累同比{fmt_pct(to_num(self.month.get((3, 25))), False)}"
        )
        return overview, model


def fmt_num(value: float | None) -> str:
    return "/" if value is None or math.isnan(value) else f"{round(value):,}"


def fmt_pct(value: float | None, space_before: bool = True) -> str:
    if value is None or math.isnan(value) or math.isinf(value):
        return "/"
    percentage = round(value * 100, 1)
    if percentage == 0:
        percentage = 0.0
    sign = "+" if percentage > 0 else ""
    prefix = " " if space_before else ""
    return f"{prefix}{sign}{percentage:.1f}%"


def fmt_pct_value(value: float | None) -> str:
    return "/" if value is None or math.isnan(value) else f"{value * 100:.1f}%"


def fmt_vs(label: str, rate: float | None) -> str:
    if rate is None or math.isnan(rate) or math.isinf(rate):
        return f"{label}暂无可比数据"
    return f"VS {label}{fmt_pct(rate)}"


def trend_word(rate: float | None) -> str:
    if rate is None or math.isnan(rate):
        return "变化"
    if rate > 0.001:
        return "有增长"
    if rate < -0.001:
        return "下降"
    return "基本持平"


def write_csv(output: Path, target: dt.date, config: CompareConfig, overview: str, model: str) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["field", "target_date", "compare_rule", "compare_label", "content"])
        writer.writerow(["概览解析-AI颜色版", target, config.rule, config.label, overview])
        writer.writerow(["车系数据解析-AI颜色版", target, config.rule, config.label, model])


def main() -> int:
    args = parse_args()
    try:
        target = infer_target_date(args.workbook, args.run_date, args.target_date)
        builder = ReportBuilder(XlsxReader(args.workbook), target)
        overview, model = builder.build()
        if not args.stdout_only:
            write_csv(args.output, target, builder.config, overview, model)
    except (OSError, KeyError, ET.ParseError, ReportDataError) as exc:
        print(f"错误：{exc}", file=sys.stderr)
        return 1

    if not args.stdout_only:
        print(f"输出：{args.output.resolve()}")
    for warning in builder.warnings:
        print(f"提示：{warning}", file=sys.stderr)
    print("\n概览解析-AI颜色版")
    print(overview)
    print("\n车系数据解析-AI颜色版")
    print(model)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
