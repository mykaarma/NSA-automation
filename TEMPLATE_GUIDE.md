# Template Configuration Wiki

Quick guide for configuring email and text notification templates.

## Table of Contents

- [Template Structure](#template-structure)
- [Available Variables](#available-variables)
- [Date Formatting](#date-formatting)
- [Template Examples](#template-examples)
- [Troubleshooting](#troubleshooting)

## Template Structure

Both templates use XML format with underscore-prefixed variables:

### Email Template
```xml
<email>
<subject>Your appointment at _dealer_name</subject>
<body>Dear _customer_firstname, your appointment at _dealer_name is confirmed for _appt_date @ _appt_start_time.</body>
<date_format>EEEE, MMMM dd, yyyy#_appt_date</date_format>
<date_format>hh:mm a#_appt_start_time</date_format>
</email>
```

### Text Template
```xml
<text>
<body>Hi _customer_firstname, your _dealer_name appointment is set for _appt_date @ _appt_start_time.</body>
<date_format>EEEE, MMMM dd, yyyy#_appt_date</date_format>
<date_format>hh:mm a#_appt_start_time</date_format>
</text>
```

## Available Variables

- **`_customer_firstname`**: Customer's first name
- **`_customer_lastname`**: Customer's last name  
- **`_dealer_name`**: Dealership name
- **`_appt_date`**: Appointment date (auto-formatted)
- **`_appt_start_time`**: Appointment time (auto-formatted)

**Example**: `Hi _customer_firstname, your _dealer_name appointment is _appt_date @ _appt_start_time.`  
**Result**: "Hi John, your Mercedes-Benz appointment is Monday, January 15, 2024 @ 09:30 AM."

## Date Formatting

Dates and times are automatically formatted using the `<date_format>` tags:

### Common Patterns
- `EEEE, MMMM dd, yyyy#_appt_date` → "Monday, January 15, 2024"
- `hh:mm a#_appt_start_time` → "09:30 AM"
- `MM/dd/yyyy#_appt_date` → "01/15/2024" 
- `MMM dd#_appt_date` → "Jan 15"

### Pattern Reference
| Pattern | Output | Pattern | Output |
|---------|---------|---------|---------|
| `EEEE` | Monday | `hh` | 09 |
| `MMMM` | January | `mm` | 30 |
| `dd` | 15 | `a` | AM |
| `yyyy` | 2024 | `HH` | 21 (24hr) |

## Template Examples

### Professional Email
```xml
<email>
<subject>Service Appointment Confirmation - _dealer_name</subject>
<body>Dear _customer_firstname _customer_lastname, Your service appointment at _dealer_name is confirmed for _appt_date at _appt_start_time. Please arrive 15 minutes early.</body>
<date_format>EEEE, MMMM dd, yyyy#_appt_date</date_format>
<date_format>hh:mm a#_appt_start_time</date_format>
</email>
```

### Casual Text Message
```xml
<text>
<body>Hi _customer_firstname! Your _dealer_name appointment: _appt_date at _appt_start_time. Questions? Call us!</body>
<date_format>MMM dd#_appt_date</date_format>
<date_format>h:mm a#_appt_start_time</date_format>
</text>
```

## Troubleshooting

### Common Issues

**Variables not replaced**: Check spelling and underscore prefix (`_customer_firstname`)
**Date formatting issues**: Verify `<date_format>` tags: `EEEE, MMMM dd, yyyy#_appt_date`
**XML parsing errors**: Validate XML structure and check for unclosed tags
**Template not loading**: Ensure files exist in `templates/` directory with UTF-8 encoding

### Quick Reference

**Template Files**: `templates/email_template.txt`, `templates/text_template.txt`

**Variable Format**: All variables use underscore prefix and are case-sensitive
- ✅ `_customer_firstname`
- ❌ `_Customer_FirstName` or `customer_firstname`

**Required XML Structure**:
```xml
<email><subject>...</subject><body>...</body><date_format>...</date_format></email>
<text><body>...</body><date_format>...</date_format></text>
```
