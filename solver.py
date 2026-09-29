import json
import boto3
import openpyxl
from io import BytesIO
from ortools.linear_solver import pywraplp
 
s3 = boto3.client("s3")
 
S3_BUCKET = "230-hcp-territories-data"
S3_KEY = "Input_UK_HCP_Universe_MockData_TerritoryZero_v1.xlsx"
 
 
def load_sheet_as_dicts(wb, sheet_name):
    ws = wb[sheet_name]
    headers = [cell.value for cell in ws[1]]
    return [dict(zip(headers, row)) for row in ws.iter_rows(min_row=2, values_only=True)]
 
 
def to_float(v):
    try:
        return float(v) if v is not None else 0.0
    except (TypeError, ValueError):
        return 0.0
 
 
def lambda_handler(event, context):
    body = json.loads(event["body"]) if "body" in event else event
    brief = body.get("brief", body)
    run_id = body.get("run_id" , "manual")
    try:
        num_territories = int(brief.get("num_territories", 8))
    except (TypeError, ValueError):
        num_territories = 8
 
    obj = s3.get_object(Bucket=S3_BUCKET, Key=S3_KEY)
    wb = openpyxl.load_workbook(BytesIO(obj["Body"].read()), data_only=True)
    hcp_rows = load_sheet_as_dicts(wb, "HCP_Universe")
    geo_rows = load_sheet_as_dicts(wb, "Geography_Reference")
    geo_lookup = {g["Postcode_Sector"]: g for g in geo_rows if g.get("Postcode_Sector")}
 
    hcps = []
    for r in hcp_rows:
        status = (r.get("Record_Status") or "").strip().lower()
        if status and status != "active":
            continue
        sector = r.get("Postcode_Sector")
        geo = geo_lookup.get(sector, {})
        hcps.append({
            "hcp_id": r.get("HCP_ID"),
            "primary_specialty": r.get("Primary_Specialty"),
            "primary_hco_name": r.get("Primary_HCO_Name"),
            "best_segment": r.get("Best_Segment"),
            "postcode_sector": sector,
            "nhs_region": geo.get("NHS_Region"),
            "workload_units": to_float(r.get("Workload_Units")),
            "value_units": to_float(r.get("Value_Units")),
        })
 
    # Roll HCPs up to postcode sector, then balance sectors across territories
    sector_totals, sector_members = {}, {}
    for h in hcps:
        sec = h["postcode_sector"] or "UNKNOWN"
        sector_totals[sec] = sector_totals.get(sec, 0.0) + h["workload_units"]
        sector_members.setdefault(sec, []).append(h)
 
    sectors = list(sector_totals.keys())
    n = len(sectors)
    target = sum(sector_totals.values()) / num_territories
 
    solver = pywraplp.Solver.CreateSolver("CBC")
    x = {(i, t): solver.BoolVar(f"x_{i}_{t}") for i in range(n) for t in range(num_territories)}
    for i in range(n):
        solver.Add(sum(x[i, t] for t in range(num_territories)) == 1)
    max_dev = solver.NumVar(0, solver.infinity(), "max_dev")
    for t in range(num_territories):
        load = sum(sector_totals[sectors[i]] * x[i, t] for i in range(n))
        solver.Add(load - target <= max_dev)
        solver.Add(target - load <= max_dev)
    solver.Minimize(max_dev)
    solver.SetTimeLimit(20000)
    solver.Solve()
 
    names = [f"Territory_{i+1}" for i in range(num_territories)]
    load_by_t = {t: 0.0 for t in names}
    value_by_t = {t: 0.0 for t in names}
    assignments = []
    for i in range(n):
        for t in range(num_territories):
            if x[i, t].solution_value() > 0.5:
                for h in sector_members[sectors[i]]:
                    load_by_t[names[t]] += h["workload_units"]
                    value_by_t[names[t]] += h["value_units"]
                    assignments.append({**{k: h[k] for k in (
                        "hcp_id", "primary_specialty", "primary_hco_name",
                        "best_segment", "postcode_sector", "nhs_region",
                        "workload_units", "value_units")},
                        "assigned_territory": names[t]})
 
    assignments.sort(key=lambda a: (a["assigned_territory"], str(a["hcp_id"])))
    territory_summary = {
        t: {"total_workload_units": round(load_by_t[t], 1),
            "total_value_units": round(value_by_t[t], 1)} for t in names}
 
    result = {"hcp_count": len(hcps), "assignments": assignments,
              "territory_summary": territory_summary}
    s3.put_object(Bucket=S3_BUCKET, Key="outputs/territory_assignment_result.json",
                  Body=json.dumps(result), ContentType="application/json")
 
    # ---- Excel export ----
    out = openpyxl.Workbook()
    ws1 = out.active
    ws1.title = "HCP_Territory_Assignment"
    cols = ["hcp_id", "primary_specialty", "primary_hco_name", "best_segment",
            "postcode_sector", "nhs_region", "assigned_territory",
            "workload_units", "value_units"]
    ws1.append(cols)
    for a in assignments:
        ws1.append([a[c] for c in cols])
    ws2 = out.create_sheet("Territory_Summary")
    ws2.append(["territory", "total_workload_units", "total_value_units"])
    for t, s in territory_summary.items():
        ws2.append([t, s["total_workload_units"], s["total_value_units"]])
    buf = BytesIO()
    out.save(buf)
    s3.put_object(Bucket=S3_BUCKET, Key="outputs/territory_assignment_result.xlsx",
                  Body=buf.getvalue(),
                  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
 
    return {"run_id": run_id, "hcp_count": len(hcps),
            "territory_summary": territory_summary,
            "json_key": f"outputs/{run_id}/territory_assignment_result.json",
            "xlsx_key": f"outputs/{run_id}/territory_assignment_result.xlsx"}
 
 
