"""
Gallery of every chart type xlchart supports, with styling examples.
Run:  python demo_all_charts.py      ->  chart_gallery.xlsx
"""
import datetime as dt
import random

from xlchart import Workbook

random.seed(7)
wb = Workbook()
data = wb.add_worksheet("Data")
classic = wb.add_worksheet("Classic charts")
modern = wb.add_worksheet("Excel 2016+ charts")
data.set_column("A:Z", 13)

# ---------------------------------------------------------------- data ------
months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
rev = [120, 135, 128, 150, 162, 171, 168, 180, 175, 190, 205, 220]
cost = [80, 88, 90, 95, 101, 108, 110, 112, 115, 118, 125, 130]
margin = [round((r - c) / r, 3) for r, c in zip(rev, cost)]
data.write_rows("A1", [["Month", "Revenue", "Cost", "Margin"]] +
                [[m, r, c, g] for m, r, c, g in zip(months, rev, cost, margin)])
for i in range(12):
    data.write((i + 1, 3), margin[i], "percent")

# pie / funnel data  (F:G)
data.write_rows("F1", [["Region", "Loans ($M)"], ["Toronto", 420], ["Ottawa", 180],
                       ["Hamilton", 140], ["London", 95], ["Other", 165]])
# funnel (I:J)
data.write_rows("I1", [["Stage", "Applications"], ["Received", 1200], ["Screened", 860],
                       ["Approved", 540], ["Funded", 410], ["Active", 395]])
# scatter / bubble (L:N)
xs = [round(random.uniform(1, 10), 2) for _ in range(25)]
ys = [round(2.1 * x + random.gauss(0, 2), 2) for x in xs]
sz = [random.randint(5, 40) for _ in xs]
data.write_rows("L1", [["Unemployment %", "Delinquency bp", "Assets"]] +
                [[a, b, c] for a, b, c in zip(xs, ys, sz)])
# radar (P:R)
data.write_rows("P1", [["Metric", "CU A", "CU B"], ["Capital", 8, 6], ["Liquidity", 7, 9],
                       ["Asset quality", 6, 7], ["Earnings", 9, 5], ["Governance", 7, 8]])
# stock HLC with dates (T:W)
data.write_row("T1", ["Date", "High", "Low", "Close"], "header")
d0, px = dt.date(2026, 9, 1), 50.0
for i in range(20):
    px += random.gauss(0, 1)
    hi, lo = px + random.uniform(0.3, 1.5), px - random.uniform(0.3, 1.5)
    data.write_row((i + 1, 19), [d0 + dt.timedelta(days=i), round(hi, 2), round(lo, 2), round(px, 2)])
data.set_column("T", 12)
# waterfall (Y:Z) - row 0 and last row are totals
wf = [["Item", "Change ($M)"], ["Opening capital", 500], ["Net income", 85],
      ["Dividends", -40], ["OCI", -25], ["Share issuance", 60], ["Q2 subtotal", None],
      ["Credit losses", -55], ["Other", 12], ["Closing capital", None]]
data.write_rows("Y1", wf)
# histogram / box (AE:AG)
data.write_row("AE1", ["PD sample A", "PD sample B", "Group"], "header")
for i in range(60):
    data.write((i + 1, 30), round(random.lognormvariate(0.5, 0.5), 2))
    data.write((i + 1, 31), round(random.lognormvariate(0.8, 0.35), 2))
    data.write((i + 1, 32), "Retail" if i % 2 else "Commercial")
# treemap / sunburst hierarchy (AI:AK)
tree = [["Sector", "Sub-sector", "Exposure"],
        ["Real estate", "Residential", 420], ["Real estate", "Commercial", 260],
        ["Consumer", "Auto", 150], ["Consumer", "Cards", 90], ["Consumer", "LOC", 70],
        ["Business", "Agriculture", 120], ["Business", "SME", 160]]
data.write_rows("AI1", tree)

W, H = 470, 300                         # chart size in px
def at(r, c):                           # grid position helper
    return (r * 16, c * 8)

# ================================================================ CLASSIC ====
# 1 clustered column with custom axis colours / fonts / gridlines
ch = wb.add_chart({"type": "column", "gap": 80})
ch.add_series({"name": "=Data!$B$1", "categories": "=Data!$A$2:$A$13",
               "values": "=Data!$B$2:$B$13", "color": "#1F4E79"})
ch.add_series({"name": "=Data!$C$1", "categories": "=Data!$A$2:$A$13",
               "values": "=Data!$C$2:$C$13", "color": "#9DC3E6"})
ch.set_title({"name": "1. Column - custom axes", "font": {"size": 13, "bold": True, "color": "1F4E79"}})
ch.set_x_axis({"line": {"color": "000000", "width": 1.25},
               "font": {"color": "C00000", "size": 9, "bold": True},
               "major_tick_mark": "out", "title": "Month"})
ch.set_y_axis({"min": 0, "max": 250, "major_unit": 50, "num_format": "$#,##0",
               "line": {"color": "000000", "width": 1.25},
               "major_gridlines": {"color": "BFBFBF", "dash": "dash"},
               "font": {"color": "375623"}, "title": {"name": "$M", "font": {"color": "375623"}}})
ch.set_legend({"position": "top"})
classic.insert_chart(at(0, 0), ch, W, H)

# 2 stacked bar (horizontal) with data labels
ch = wb.add_chart({"type": "bar", "subtype": "stacked"})
for col, colr in (("B", "4472C4"), ("C", "ED7D31")):
    ch.add_series({"name": f"=Data!${col}$1", "categories": "=Data!$A$2:$A$7",
                   "values": f"=Data!${col}$2:${col}$7", "color": colr,
                   "data_labels": {"value": True, "position": "center",
                                   "font": {"color": "FFFFFF", "bold": True}}})
ch.set_title({"name": "2. Stacked bar + labels"})
ch.set_x_axis({"reverse": True})
classic.insert_chart(at(0, 1), ch, W, H)

# 3 percent-stacked column
ch = wb.add_chart({"type": "column", "subtype": "percent_stacked"})
for col in "BC":
    ch.add_series({"name": f"=Data!${col}$1", "categories": "=Data!$A$2:$A$13",
                   "values": f"=Data!${col}$2:${col}$13"})
ch.set_title({"name": "3. 100% stacked column"})
classic.insert_chart(at(0, 2), ch, W, H)

# 4 line with markers, dashes, trendline
ch = wb.add_chart({"type": "line"})
ch.add_series({"name": "Revenue", "categories": "=Data!$A$2:$A$13", "values": "=Data!$B$2:$B$13",
               "line": {"color": "7030A0", "width": 2.5},
               "marker": {"type": "diamond", "size": 7, "fill": {"color": "FFFFFF"},
                          "border": {"color": "7030A0", "width": 1.5}},
               "trendline": {"type": "linear", "name": "Trend", "forward": 2,
                             "line": {"color": "7030A0", "dash": "dot", "width": 1}}})
ch.add_series({"name": "Cost", "categories": "=Data!$A$2:$A$13", "values": "=Data!$C$2:$C$13",
               "line": {"color": "808080", "width": 1.75, "dash": "dash"}, "smooth": True})
ch.set_title({"name": "4. Line, markers, trendline"})
ch.set_plotarea({"fill": {"color": "F7F7F7"}})
classic.insert_chart(at(1, 0), ch, W, H)

# 5 stacked area with transparency
ch = wb.add_chart({"type": "area", "subtype": "stacked"})
ch.add_series({"name": "Cost", "categories": "=Data!$A$2:$A$13", "values": "=Data!$C$2:$C$13",
               "fill": {"color": "5B9BD5", "transparency": 30}})
ch.add_series({"name": "Revenue", "categories": "=Data!$A$2:$A$13", "values": "=Data!$B$2:$B$13",
               "fill": {"color": "FFC000", "transparency": 30}})
ch.set_title({"name": "5. Stacked area"})
ch.set_x_axis({"position_axis": "on_tick"})
classic.insert_chart(at(1, 1), ch, W, H)

# 6 combo: column + line on secondary axis
ch = wb.add_chart({"type": "column"})
ch.add_series({"name": "Revenue", "categories": "=Data!$A$2:$A$13", "values": "=Data!$B$2:$B$13",
               "color": "A5A5A5"})
ln = wb.add_chart({"type": "line"})
ln.add_series({"name": "Margin", "categories": "=Data!$A$2:$A$13", "values": "=Data!$D$2:$D$13",
               "y2_axis": True, "line": {"color": "C00000", "width": 2.5},
               "marker": {"type": "circle", "size": 6}})
ch.combine(ln)
ch.set_title({"name": "6. Combo, secondary axis"})
ch.set_y2_axis({"num_format": "0%", "min": 0, "max": 0.5, "font": {"color": "C00000"},
                "line": {"color": "C00000"}})
classic.insert_chart(at(1, 2), ch, W, H)

# 7 pie with custom slice colours + exploded slice
ch = wb.add_chart({"type": "pie", "first_slice_angle": 30})
ch.add_series({"name": "Loans", "categories": "=Data!$F$2:$F$6", "values": "=Data!$G$2:$G$6",
               "colors": ["1F4E79", "2E75B6", "9DC3E6", "DEEBF7", "FFC000"],
               "points": {4: {"explosion": 12}},
               "data_labels": {"percentage": True, "category": True, "position": "outside_end",
                               "separator": "newline", "leader_lines": True}})
ch.set_title({"name": "7. Pie"})
ch.set_legend({"position": "right"})
classic.insert_chart(at(2, 0), ch, W, H)

# 8 doughnut
ch = wb.add_chart({"type": "doughnut", "hole_size": 60})
ch.add_series({"name": "Loans", "categories": "=Data!$F$2:$F$6", "values": "=Data!$G$2:$G$6",
               "data_labels": {"value": True, "font": {"color": "FFFFFF", "bold": True}}})
ch.set_title({"name": "8. Doughnut"})
classic.insert_chart(at(2, 1), ch, W, H)

# 9 scatter with polynomial trendline + equation, R²
ch = wb.add_chart({"type": "scatter"})
ch.add_series({"name": "Credit unions", "categories": "=Data!$L$2:$L$26",
               "values": "=Data!$M$2:$M$26",
               "marker": {"type": "circle", "size": 7, "fill": {"color": "ED7D31", "transparency": 30},
                          "border": {"color": "C55A11"}},
               "trendline": {"type": "polynomial", "order": 2, "display_equation": True,
                             "display_r_squared": True}})
ch.set_title({"name": "9. Scatter + trendline"})
ch.set_x_axis({"title": "Unemployment %", "min": 0, "max": 11})
ch.set_y_axis({"title": "Delinquency (bp)"})
ch.set_legend({"position": "none"})
classic.insert_chart(at(2, 2), ch, W, H)

# 10 bubble
ch = wb.add_chart({"type": "bubble", "bubble_scale": 60})
ch.add_series({"name": "CUs", "categories": "=Data!$L$2:$L$26", "values": "=Data!$M$2:$M$26",
               "sizes": "=Data!$N$2:$N$26", "color": "70AD47"})
ch.set_title({"name": "10. Bubble"})
ch.set_legend({"position": "none"})
classic.insert_chart(at(3, 0), ch, W, H)

# 11 radar filled
ch = wb.add_chart({"type": "radar", "subtype": "filled"})
ch.add_series({"name": "=Data!$Q$1", "categories": "=Data!$P$2:$P$6", "values": "=Data!$Q$2:$Q$6"})
ch.add_series({"name": "=Data!$R$1", "categories": "=Data!$P$2:$P$6", "values": "=Data!$R$2:$R$6"})
ch.set_title({"name": "11. Radar (filled)"})
classic.insert_chart(at(3, 1), ch, W, H)

# 12 radar with markers
ch = wb.add_chart({"type": "radar", "subtype": "with_markers"})
ch.add_series({"name": "=Data!$Q$1", "categories": "=Data!$P$2:$P$6", "values": "=Data!$Q$2:$Q$6"})
ch.add_series({"name": "=Data!$R$1", "categories": "=Data!$P$2:$P$6", "values": "=Data!$R$2:$R$6"})
ch.set_title({"name": "12. Radar (markers)"})
classic.insert_chart(at(3, 2), ch, W, H)

# 13 stock high-low-close on a date axis
ch = wb.add_chart({"type": "stock", "high_low_lines": {"color": "404040", "width": 1}})
for col in "UVW":
    ch.add_series({"name": f"=Data!${col}$1", "categories": "=Data!$T$2:$T$21",
                   "values": f"=Data!${col}$2:${col}$21"})
ch.series[2]["marker"] = {"type": "dash", "size": 6, "fill": {"color": "C00000"},
                          "border": {"color": "C00000"}}
ch.set_title({"name": "13. Stock (HLC), date axis"})
ch.set_x_axis({"date_axis": True, "num_format": "mmm d", "font": {"rotation": -45}})
ch.set_legend({"position": "none"})
classic.insert_chart(at(4, 0), ch, W, H)

# 14 scatter smooth lines, log y-axis, error bars
ch = wb.add_chart({"type": "scatter", "subtype": "smooth_with_markers"})
ch.add_series({"name": "Revenue", "categories": [1, 2, 3, 4, 5, 6],
               "values": [3, 10, 30, 100, 300, 1000], "color": "264478",
               "y_error_bars": {"type": "percentage", "value": 20}})
ch.set_title({"name": "14. Smooth scatter, log axis, error bars"})
ch.set_y_axis({"log_base": 10, "minor_gridlines": {"color": "F2F2F2"}})
ch.set_legend({"position": "none"})
classic.insert_chart(at(4, 1), ch, W, H)

# 15 classic waterfall (works in every Excel version)
ch = wb.add_waterfall_chart(labels="=Data!$Y$2:$Y$10", values="=Data!$Z$2:$Z$10",
                            totals=[0, 5, 8], helper_cell="AP1", hide_helpers=False,
                            title="15. Waterfall (classic, any Excel)")
ch.set_y_axis({"min": 400, "num_format": "#,##0"})
ch.set_x_axis({"font": {"rotation": -30}})
classic.insert_chart(at(4, 2), ch, W, H)

# 16 column with per-point colours + data table + dark theme
ch = wb.add_chart({"type": "column", "gap": 40})
ch.add_series({"name": "Revenue", "categories": "=Data!$A$2:$A$7", "values": "=Data!$B$2:$B$7",
               "fill": {"gradient": {"colors": ["9DC3E6", "1F4E79"], "angle": 90}},
               "points": {5: {"fill": "FFC000"}}})
ch.set_title({"name": "16. Dark style, gradient, data table", "font": {"color": "FFFFFF"}})
ch.set_chartarea({"fill": {"color": "262626"}, "border": {"none": True}})
ch.set_plotarea({"fill": {"color": "262626"}})
ch.set_font({"color": "D9D9D9"})
ch.set_y_axis({"major_gridlines": {"color": "404040"}})
ch.set_x_axis({"line": {"color": "808080"}})
ch.set_table()
ch.set_legend({"position": "none"})
classic.insert_chart(at(5, 0), ch, W, H)

# ================================================================ MODERN =====
# 17 native waterfall
# native waterfall: deltas, and total rows hold the total value itself
data.write_rows("AM1", [["Item", "Value"], ["Opening capital", 500], ["Net income", 85],
                        ["Dividends", -40], ["OCI", -25], ["Share issuance", 60],
                        ["Q2 subtotal", 580], ["Credit losses", -55], ["Other", 12],
                        ["Closing capital", 537]])
ch = wb.add_chart({"type": "waterfall"})
ch.add_series({"name": "Capital", "categories": "=Data!$AM$2:$AM$10",
               "values": "=Data!$AN$2:$AN$10", "totals": [0, 5, 8],
               "increase_color": "70AD47", "decrease_color": "C00000", "total_color": "1F4E79",
               "data_labels": {"value": True, "position": "outside_end", "num_format": "#,##0"}})
ch.set_title({"name": "17. Waterfall (native)"})
ch.set_y_axis({"min": 400, "font": {"color": "1F4E79"}})
ch.set_x_axis({"line": {"color": "000000"}})
modern.insert_chart(at(0, 0), ch, W, H)

# 18 histogram
ch = wb.add_chart({"type": "histogram"})
ch.add_series({"name": "PD A", "values": "=Data!$AE$2:$AE$61", "color": "5B9BD5",
               "border": {"color": "FFFFFF", "width": 1}, "bins": {"count": 10}})
ch.set_title({"name": "18. Histogram"})
modern.insert_chart(at(0, 1), ch, W, H)

# 19 pareto
ch = wb.add_chart({"type": "pareto"})
ch.add_series({"name": "Loans", "categories": "=Data!$F$2:$F$6", "values": "=Data!$G$2:$G$6",
               "color": "4472C4", "pareto_line": {"color": "C00000", "width": 2}})
ch.set_title({"name": "19. Pareto"})
modern.insert_chart(at(0, 2), ch, W, H)

# 20 box & whisker (two series)
ch = wb.add_chart({"type": "box_whisker"})
ch.add_series({"name": "PD A", "values": "=Data!$AE$2:$AE$61", "color": "9DC3E6",
               "border": {"color": "1F4E79"}, "mean_marker": True, "outliers": True})
ch.add_series({"name": "PD B", "values": "=Data!$AF$2:$AF$61", "color": "F8CBAD",
               "border": {"color": "C55A11"}})
ch.set_title({"name": "20. Box & whisker"})
modern.insert_chart(at(1, 0), ch, W, H)

# 21 funnel
ch = wb.add_chart({"type": "funnel"})
ch.add_series({"name": "Pipeline", "categories": "=Data!$I$2:$I$6", "values": "=Data!$J$2:$J$6",
               "color": "2E75B6", "data_labels": {"value": True, "font": {"color": "FFFFFF", "bold": True}}})
ch.set_title({"name": "21. Funnel"})
modern.insert_chart(at(1, 1), ch, W, H)

# 22 treemap (two-level hierarchy)
ch = wb.add_chart({"type": "treemap"})
ch.add_series({"name": "Exposure", "categories": "=Data!$AI$2:$AJ$8", "values": "=Data!$AK$2:$AK$8",
               "parent_labels": "banner"})
ch.set_title({"name": "22. Treemap"})
modern.insert_chart(at(1, 2), ch, W, H)

# 23 sunburst
ch = wb.add_chart({"type": "sunburst"})
ch.add_series({"name": "Exposure", "categories": "=Data!$AI$2:$AJ$8", "values": "=Data!$AK$2:$AK$8"})
ch.set_title({"name": "23. Sunburst"})
modern.insert_chart(at(2, 0), ch, W, H)

wb.save("chart_gallery.xlsx")
print("wrote chart_gallery.xlsx")
