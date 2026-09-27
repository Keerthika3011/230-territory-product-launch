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
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rows.append(dict(zip(headers, row)))
    return rows
 
 
def lambda_handler(event, context):
    if "body" in event:
        body = json.loads(event["body"])
    else:
        body = event
 
    brief = body.get("brief", body)
    num_territories = brief.get("num_territories", 8)
 
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
 
        postcode_sector = r.get("Postcode_Sector")
        geo = geo_lookup.get(postcode_sector, {})
 
        workload = r.get("Workload_Units")
        value = r.get("Value_Units")
        try:
            workload = float(workload) if workload is not None else 0.0
        except (TypeError, ValueError):
            workload = 0.0
        try:
            value = float(value) if value is not None else 0.0
        except (TypeError, ValueError):
            value = 0.0
 
        hcps.append({
            "hcp_id": r.get("HCP_ID"),
            "primary_specialty": r.get("Primary_Specialty"),
            "primary_hco_name": r.get("Primary_HCO_Name"),
            "best_segment": r.get("Best_Segment"),
            "postcode_sector": postcode_sector,
            "nhs_region": geo.get("NHS_Region"),
            "workload_units": workload,
            "value_units": value,
        })
 
    # --- Real OR-Tools optimization ---
    # NOTE: solving a pure per-HCP assignment problem with ~1000 binary
    # variables x territories can be slow/heavy for CBC in a Lambda's time
    # limit. To keep this practical, we pre-group HCPs into buckets by
    # postcode_sector (summed workload per sector), and let the solver
    # balance sectors across territories instead of individual HCPs.
    sector_totals = {}
    sector_members = {}
    for h in hcps:
        sec = h["postcode_sector"] or "UNKNOWN"
        sector_totals[sec] = sector_totals.get(sec, 0.0) + h["workload_units"]
        sector_members.setdefault(sec, []).append(h)
 
    sectors = list(sector_totals.keys())
    n = len(sectors)
    total_workload = sum(sector_totals.values())
    target_per_territory = total_workload / num_territories if num_territories else 0
 
    solver = pywraplp.Solver.CreateSolver("CBC")
    x = {}
    for i in range(n):
        for t in range(num_territories):
            x[i, t] = solver.BoolVar(f"x_{i}_{t}")
 
    # Each sector assigned to exactly one territory
    for i in range(n):
        solver.Add(sum(x[i, t] for t in range(num_territories)) == 1)
 
    # Minimize the maximum deviation from the ideal balanced load
    max_dev = solver.NumVar(0, solver.infinity(), "max_dev")
    for t in range(num_territories):
        load = sum(sector_totals[sectors[i]] * x[i, t] for i in range(n))
        solver.Add(load - target_per_territory <= max_dev)
        solver.Add(target_per_territory - load <= max_dev)
 
    solver.Minimize(max_dev)
    solver.SetTimeLimit(20000)  # 20 second cap so it can't run past Lambda's timeout
    solver.Solve()
 
    territory_names = [f"Territory_{i+1}" for i in range(num_territories)]
    territory_load = {t: 0.0 for t in territory_names}
    territory_value = {t: 0.0 for t in territory_names}
    assignments = []
 
    for i in range(n):
        for t in range(num_territories):
            if x[i, t].solution_value() > 0.5:
                tname = territory_names[t]
                sec = sectors[i]
                for h in sector_members[sec]:
                    territory_load[tname] += h["workload_units"]
                    territory_value[tname] += h["value_units"]
                    assignments.append({
                        "hcp_id": h["hcp_id"],
                        "primary_specialty": h["primary_specialty"],
                        "primary_hco_name": h["primary_hco_name"],
                        "best_segment": h["best_segment"],
                        "postcode_sector": h["postcode_sector"],
                        "nhs_region": h["nhs_region"],
                        "assigned_territory": tname,
                        "workload_units": h["workload_units"],
                        "value_units": h["value_units"],
                    })
 
    territory_summary = {
        t: {"total_workload_units": round(territory_load[t], 1),
            "total_value_units": round(territory_value[t], 1)}
        for t in territory_names
    }
 
    result_payload = {
        "hcp_count": len(hcps),
        "assignments": assignments,
        "territory_summary": territory_summary
    }
 
    output_key = "outputs/territory_assignment_result.json"
    s3.put_object(
        Bucket=S3_BUCKET,
        Key=output_key,
        Body=json.dumps(result_payload),
        ContentType="application/json"
    )
 
    return {
        "statusCode": 200,
        "body": json.dumps({
            "message": "Territory assignment complete (OR-Tools optimized)",
            "hcp_count": len(hcps),
            "sectors_balanced": n,
            "result_location": f"s3://{S3_BUCKET}/{output_key}"
        })
    }
 
