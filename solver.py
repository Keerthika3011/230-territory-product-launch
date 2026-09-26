import json
import boto3
import openpyxl
from io import BytesIO
 
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
 
    # Real join key is Postcode_Sector, not Postcode_Prefix
    geo_lookup = {g["Postcode_Sector"]: g for g in geo_rows if g.get("Postcode_Sector")}
 
    hcps = []
    for r in hcp_rows:
        # Only include active records
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
            "target_flag": r.get("Target_Flag"),
            "target_tier": r.get("Target_Tier"),
            "postcode_sector": postcode_sector,
            "nhs_region": geo.get("NHS_Region"),
            "nhs_nation": geo.get("NHS_Nation"),
            "workload_units": workload,
            "value_units": value,
        })
 
    territory_names = [f"Territory_{i+1}" for i in range(num_territories)]
    territory_load = {t: 0.0 for t in territory_names}
    territory_value = {t: 0.0 for t in territory_names}
    assignments = []
 
    # Greedy balance on Workload_Units — the metric the source data is designed around
    for h in sorted(hcps, key=lambda x: -x["workload_units"]):
        lightest = min(territory_load, key=territory_load.get)
        territory_load[lightest] += h["workload_units"]
        territory_value[lightest] += h["value_units"]
        assignments.append({
            "hcp_id": h["hcp_id"],
            "primary_specialty": h["primary_specialty"],
            "primary_hco_name": h["primary_hco_name"],
            "best_segment": h["best_segment"],
            "postcode_sector": h["postcode_sector"],
            "nhs_region": h["nhs_region"],
            "assigned_territory": lightest,
            "workload_units": h["workload_units"],
            "value_units": h["value_units"],
        })
 
    territory_summary = {
        t: {"total_workload_units": round(territory_load[t], 1),
            "total_value_units": round(territory_value[t], 1)}
        for t in territory_names
    }
 
    return {
        "statusCode": 200,
        "body": json.dumps({
            "hcp_count": len(hcps),
            "assignments": assignments,
            "territory_summary": territory_summary
        })
    }
 
