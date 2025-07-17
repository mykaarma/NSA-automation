import os
import requests
import time
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from dotenv import load_dotenv
import openpyxl
from openpyxl import Workbook
from dealer_info import DEALERS, get_dealer_by_id
import json

try:
    from tqdm import tqdm
    USE_TQDM = True
except ImportError:
    USE_TQDM = False

# Load environment variables
load_dotenv()

BASE_URL = os.getenv('MYKAARMA_BASE_URL')
USERNAME = os.getenv('MYKAARMA_USERNAME')
PASSWORD = os.getenv('MYKAARMA_PASSWORD')
AUTH = (USERNAME, PASSWORD)

def fetch_slot_size(dealer_uuid):
    url = f"{BASE_URL}/appointment/v2/dealer/{dealer_uuid}/hoursOfOperation"
    headers = {"accept": "application/json"}
    cookies = {"rollout.stage": "canary"}
    resp = requests.get(url, auth=AUTH, headers=headers, cookies=cookies)
    resp.raise_for_status()
    data = resp.json()
    return int(data.get('slotSizeInMins', 15))  # Default to 15 if not present

def load_opcodes_from_xlsx(filename):
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    opcodes = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0]:
            opcode = str(row[0]).strip()
            description = str(row[1]).strip() if row[1] else ""
            opcodes[opcode] = description
    return opcodes

def get_first_available_slot_firstapi(row, department_uuid, filtered_opcodes, target_date):
    url = f"{BASE_URL}/appointment/v2/department/{department_uuid}/first-available-slot"
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    cookies = {"rollout.stage": "canary"}
    # If target_date is in the past, use today
    from_date = max(target_date, datetime.today())
    date_str = from_date.strftime('%Y-%m-%d')
    body = {
        "dates": [date_str],
        "customerInformation": {
            "firstName": row.get('Customer First Name'),
            "lastName": row.get('Customer Last Name'),
            "uuid": row.get('Customer UUID'),
            "key": row.get('Customer Key'),
        },
        "vehicleInformation": {
            "uuid": row.get('Vehicle UUID'),
            "vin": row.get('VIN'),
        },
        "laborOpcodeList": filtered_opcodes,
        "selectedAvailabilityAttributes": {},
        "allAvailabilityAttributes": {}
    }
    resp = requests.post(url, json=body, auth=AUTH, headers=headers, cookies=cookies)
    resp.raise_for_status()
    data = resp.json()
    dt = data.get('dateTime')
    if dt:
        # dt is like '2024-08-01 09:15:00'
        appt_date, appt_time = dt.split(' ')
        return appt_date, appt_time
    return None, None

def create_appointment(row, appt_date, appt_time, slot_size, dealer_uuid, filtered_opcodes, opcode_descriptions):
    url = f"{BASE_URL}/appointment/v2/dealer/{dealer_uuid}/appointment"
    headers = {"accept": "application/json", "Content-Type": "application/json"}
    cookies = {"rollout.stage": "canary"}
    service_list = []
    for op in filtered_opcodes:
        service_item = {"title": op, "operationType": "OPCODE"}
        if opcode_descriptions.get(op):
            service_item["description"] = opcode_descriptions[op]
        service_list.append(service_item)
    # Calculate end time
    start_dt = datetime.strptime(f"{appt_date}T{appt_time}", "%Y-%m-%dT%H:%M:%S")
    end_dt = start_dt + timedelta(minutes=slot_size) - timedelta(seconds=1)
    end_time = end_dt.strftime("%H:%M:%S")
    body = {
        "customerUuid": row['Customer UUID'],
        "vehicleInformation": {
            "vehicleUuid": row['Vehicle UUID'],
            "vin": row['VIN']
        },
        "appointmentInformation": {
            "appointmentStartDateTime": f"{appt_date}T{appt_time}",
            "appointmentEndDateTime": f"{appt_date}T{end_time}",
            "serviceList": service_list,
            "comments": "",
            "internalNotes": "Next Service Appointment scheduled automatically by script.",
            "customerAppointmentPreference": {
                "notifyCustomer": False,
                "emailConfirmation": False,
                "textConfirmation": False,
                "emailReminder": False,
                "textReminder": False
            },
            "status": None,
            "recall": False,
            "pushToDms": True
        }
    }
    resp = requests.post(url, json=body, auth=AUTH, headers=headers, cookies=cookies)
    resp.raise_for_status()
    return resp.json()

def prefetch_dealer_context(rows):
    unique_dealer_ids = set(row['Dealer ID'] for row in rows)
    dealer_context = {}
    for dealer_id in unique_dealer_ids:
        dealer_info = get_dealer_by_id(dealer_id)
        if not dealer_info:
            dealer_context[dealer_id] = None
            continue
        slot_size = fetch_slot_size(dealer_info['dealer_uuid'])
        valid_opcodes = load_opcodes_from_xlsx(dealer_info['opcode_xlsx'])
        dealer_context[dealer_id] = {
            'dealer_info': dealer_info,
            'slot_size': slot_size,
            'valid_opcodes': valid_opcodes
        }
    return dealer_context

def main():
    xlsx_file = "closed_ros.xlsx"
    if not os.path.exists(xlsx_file):
        print("Error: closed_ros.xlsx not found in the current directory.")
        return
    wb = openpyxl.load_workbook(xlsx_file)
    ws = wb.active
    rows = [dict(zip([cell.value for cell in ws[1]], [cell.value for cell in row])) for row in ws.iter_rows(min_row=2)]
    total = len(rows)
    dealer_context = prefetch_dealer_context(rows)
    results = []
    if USE_TQDM:
        iterator = tqdm(rows, desc="Scheduling Appointments", unit="appt")
    else:
        iterator = rows
    for idx, row in enumerate(iterator, 1):
        dealer_id = row['Dealer ID']
        context = dealer_context.get(dealer_id)
        if not context or not context['dealer_info']:
            print(f"  [{idx}/{total}] Dealer ID {dealer_id} not found in dealer_info.py. Skipping.")
            row['NSA Status'] = 'FAILED'
            row['NSA Date'] = ''
            row['NSA UUID'] = ''
            results.append(row)
            continue
        dealer_info = context['dealer_info']
        slot_size = context['slot_size']
        valid_opcodes = context['valid_opcodes']
        opcodes = row['Opcodes'].split(',') if row.get('Opcodes') else []
        filtered_opcodes = [op for op in opcodes if op in valid_opcodes]
        
        # Add default NSA opcode for this dealer (for reporting and identification)
        default_nsa_opcode = dealer_info.get('default_nsa_opcode')
        if default_nsa_opcode and default_nsa_opcode not in filtered_opcodes:
            filtered_opcodes.append(default_nsa_opcode)
        
        # Create opcode_descriptions dict for filtered opcodes
        opcode_descriptions = {op: valid_opcodes.get(op, '') for op in filtered_opcodes}
        close_date = row['RO Close Date']
        months = dealer_info.get('next_service_interval_in_months', 6)
        target_date = datetime.strptime(close_date, '%Y-%m-%d') + relativedelta(months=months)
        appt_uuid = None
        appt_date = None
        appt_time = None
        for attempt in range(1, 3):
            appt_date, appt_time = get_first_available_slot_firstapi(row, dealer_info['department_uuid'], filtered_opcodes, target_date)
            if appt_date and appt_time:
                try:
                    resp = create_appointment(row, appt_date, appt_time, slot_size, dealer_info['dealer_uuid'], filtered_opcodes, opcode_descriptions)
                    appt_uuid = resp.get('appointmentUuid', 'Success')
                    row['NSA Status'] = 'SUCCESS'
                    row['NSA Date'] = f"{appt_date} {appt_time}"
                    row['NSA UUID'] = appt_uuid
                    results.append(row)
                    print(f"  [{idx}/{total}] Appointment created for {appt_date} {appt_time}: {appt_uuid}")
                    break
                except Exception as e:
                    print(f"  [{idx}/{total}] Attempt {attempt} failed: {e}")
                    time.sleep(2)
            else:
                print(f"  [{idx}/{total}] No available slot found after {target_date.date()} (using first-available-slot endpoint). Attempt {attempt}.")
                time.sleep(2)
        if not appt_uuid:
            row['NSA Status'] = 'FAILED'
            row['NSA Date'] = ''
            row['NSA UUID'] = ''
            results.append(row)
        time.sleep(2)
    # Write results to xlsx
    out_filename = xlsx_file.replace('closed_ros', 'schedule_results')
    out_wb = Workbook()
    out_ws = out_wb.active
    out_fields = list(results[0].keys())
    out_ws.append(out_fields)
    for r in results:
        out_ws.append([r.get(f) for f in out_fields])
    out_wb.save(out_filename)
    print(f"\nWrote appointment scheduling results to {out_filename}")

if __name__ == "__main__":
    main() 
