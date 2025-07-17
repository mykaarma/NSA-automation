# SAMPLE DEALER CONFIGURATION
# Replace this with your actual dealer information

DEALERS = {
    '1': {
        'name': 'Sample Dealer 1',
        'dealer_uuid': 'REPLACE_WITH_YOUR_DEALER_UUID',
        'department_uuid': 'REPLACE_WITH_YOUR_DEPARTMENT_UUID',
        'opcode_xlsx': 'sample_opcodes.xlsx',
        'next_service_interval_in_months': 12,
        'default_nsa_opcode': 'REPLACE_WITH_DEFAULT_NSA_OPCODE',
    },
    '2': {
        'name': 'Sample Dealer 2',
        'dealer_uuid': 'REPLACE_WITH_YOUR_DEALER_UUID',
        'department_uuid': 'REPLACE_WITH_YOUR_DEPARTMENT_UUID',
        'opcode_xlsx': 'sample_opcodes.xlsx',
        'next_service_interval_in_months': 10,
        'default_nsa_opcode': 'REPLACE_WITH_DEFAULT_NSA_OPCODE',
    }
}

def get_dealer_by_id(dealer_id):
    return DEALERS.get(dealer_id) 