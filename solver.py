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
    hco_rows = load_sheet_as_dicts(wb, "HCO_Master")
    geo_rows = load_sheet_as_dicts(wb, "Geography_Reference")
 
    # Build lookup: HCO_ID -> HCO details
    hco_lookup = {h["HCO_ID"]: h for h in hco_rows}
 
    # Build lookup: Postcode_Prefix -> Geography zone
    geo_lookup = {g["Postcode_Prefix"]: g for g in geo_rows}
 
    priority_weight = {"High": 3, "Medium": 2, "Low": 1}
    segment_bonus = {"Target": 2, "Non-target": 0}
 
    hcps = []
    for r in hcp_rows:
        hco = hco_lookup.get(r.get("Primary_HCO_ID"), {})
        postcode_prefix = (r.get("Postcode") or "").split(" ")[0]
        geo = geo_lookup.get(postcode_prefix, {})
 
        score = (
            priority_weight.get(r.get("Call_Priority", "Medium"), 2)
            + segment_bonus.get(r.get("Segment", "Non-target"), 0)
        )
 
        hcps.append({
            "hcp_id": r["HCP_ID"],
            "call_priority": r.get("Call_Priority"),
            "segment": r.get("Segment"),
            "postcode": r.get("Postcode"),
            "hco_name": r.get("Primary_HCO_Name"),
            "formulary_status": hco.get("Formulary_Status_Cardiology"),
            "geo_zone": geo.get("Suggested_Territory_Zone"),
            "workload_score": score
        })
 
    territory_names = [f"Territory_{i+1}" for i in range(num_territories)]
    territory_load = {t: 0 for t in territory_names}
    assignments = []
 
    for h in sorted(hcps, key=lambda x: -x["workload_score"]):
        lightest = min(territory_load, key=territory_load.get)
        territory_load[lightest] += h["workload_score"]
        assignments.append({
            "hcp_id": h["hcp_id"],
            "assigned_territory": lightest,
            "workload_score": h["workload_score"],
            "geo_zone": h["geo_zone"]
        })
 
    return {
        "statusCode": 200,
        "body": json.dumps({
            "assignments": assignments,
            "territory_summary": territory_load
        })
    }
 
