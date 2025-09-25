# NSA Automation Tool

This tool helps you automatically extract closed Repair Orders (ROs) from the myKaarma system and schedule follow-up appointments for customers. It's designed to save time by automating the process of identifying customers who need next service appointments (NSA) and creating those appointments in the system.

## What This Tool Does

1. **Extract Closed ROs**: Looks at repair orders that were closed on a specific date or date range and filters them based on specific service codes (opcodes)
2. **Schedule Appointments**: Automatically creates follow-up service appointments for customers whose ROs contained specific service codes
3. **Track Appointments**: Maintains a cache of created appointments using RO numbers to prevent duplicates and prompt users about existing appointments
4. **Send Notifications**: Automatically sends customizable text and email notifications to customers about their scheduled appointments


## Table of Contents


- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Project Setup](#project-setup)
- [Configuration](#configuration)
- [How to Use](#how-to-use)
- [Notifications & Templates](#notifications--templates)
- [File Structure](#file-structure)
- [Support](#support)

## Prerequisites

### What You Need Before Starting

1. **Computer**: Windows, Mac, or Linux computer
2. **Internet Connection**: Required to download software and access myKaarma APIs
3. **myKaarma Account**: You need valid credentials to access the myKaarma system
4. **Basic Computer Skills**: Ability to download files, create folders, and run simple commands

## Installation

### Step 1: Install Python

Python is a programming language that this tool is built with. You need to install it first.

#### For Windows:
1. Go to [python.org](https://python.org/downloads/)
2. Click "Download Python" (it will download the latest version)
3. Run the downloaded file
4. **IMPORTANT**: Check the box "Add Python to PATH" during installation
5. Click "Install Now"
6. When installation is complete, click "Close"

#### For Mac:
1. Go to [python.org](https://python.org/downloads/)
2. Click "Download Python" 
3. Open the downloaded `.pkg` file
4. Follow the installation wizard
5. Click "Install" and enter your password when prompted

#### For Linux (Ubuntu/Debian):
Open Terminal and run:
```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

### Step 2: Verify Python Installation

1. Open Command Prompt (Windows) or Terminal (Mac/Linux)
2. Type this command and press Enter:
   ```bash
   python --version
   ```
   OR
   ```bash
   python3 --version
   ```
3. You should see something like "Python 3.9.0" or similar
4. If you get an error, Python wasn't installed correctly - repeat Step 1

## Project Setup

### Step 1: Download the Project

1. Download all the project files to a folder on your computer
2. Create a new folder called `nsa-automation` on your Desktop
3. Put all the downloaded files in this folder

### Step 2: Open Command Line in Project Folder

#### For Windows:
1. Open File Explorer
2. Navigate to your `nsa-automation` folder
3. In the address bar, type `cmd` and press Enter
4. A black Command Prompt window will open

#### For Mac:
1. Open Finder
2. Navigate to your `nsa-automation` folder
3. Right-click in the folder and select "Services" > "New Terminal at Folder"

#### For Linux:
1. Open your file manager
2. Navigate to your `nsa-automation` folder
3. Right-click and select "Open in Terminal"

### Step 3: Create Virtual Environment

A virtual environment keeps this project's requirements separate from other Python projects.

1. In the command line, type:
   ```bash
   python -m venv venv
   ```
   OR (if python command doesn't work):
   ```bash
   python3 -m venv venv
   ```

2. Wait for it to finish (you'll see the cursor return)

### Step 4: Activate Virtual Environment

#### For Windows:
```bash
venv\Scripts\activate
```

#### For Mac/Linux:
```bash
source venv/bin/activate
```

You should see `(venv)` at the beginning of your command line prompt, which means the virtual environment is active.

### Step 5: Install Required Packages

Copy and paste this command and press Enter:
```bash
pip install -r requirements.txt
```

Wait for all packages to install. You should see messages about successfully installing various packages.

## Configuration

### Step 1: Create Environment File

1. In your project folder, create a new file called `.env` (note the dot at the beginning)
2. Open this file with a text editor (Notepad on Windows, TextEdit on Mac)
3. Add the following content, replacing the placeholder values with your actual myKaarma credentials:

```
MYKAARMA_BASE_URL=https://api.mykaarma.com
MYKAARMA_USERNAME=your_username
MYKAARMA_PASSWORD=your_password
PAGE_SIZE=500
```

**Where to get these values:**
- `MYKAARMA_BASE_URL`: `https://api.mykaarma.com`)
- `MYKAARMA_USERNAME`: Your myKaarma username
- `MYKAARMA_PASSWORD`: Your myKaarma password
- `PAGE_SIZE`: Number of ROs to fetch for a date at once (500 is good for most cases)

### Step 2: Create Opcode Files

Opcodes are service codes that identify specific types of repairs or services. You need to create Excel files containing the opcodes you want to track for each dealer.

**Note**: A sample opcode file (`sample_opcodes.xlsx`) is provided in the project for reference. You can copy this file and modify it for your specific needs.

#### How to Create Opcode Files:

1. **Open Excel** (or Google Sheets, then save as .xlsx)
2. **Create a new spreadsheet**
3. **In cell A1**, type: `Opcode` and **in cell B1**, type: `Description`
4. **Starting from A2**, list all the opcodes you want to track, one per row, with optional descriptions in column B:
   ```
   A1: Opcode          B1: Description
   A2: 001150          B2: Oil Change Service
   A3: 002100          B3: Brake Inspection
   A4: 003050          B4: Engine Diagnostic
   ... and so on
   ```
5. **Save the file** with a descriptive name like:
   - `mercedes_opcodes.xlsx` (for Mercedes dealers)
   - `audi_riverside_opcodes.xlsx` (for Audi Riverside)
   - `porsche_ontario_opcodes.xlsx` (for Porsche Ontario)

**Example Opcode File Structure:**
```
| Opcode  | Description         |
|---------|---------------------|
| 001150  | Oil Change Service  |
| 002100  | Brake Inspection    |
| 003050  | Engine Diagnostic   |
| 004200  | Tire Rotation       |
```

#### Create These Files:
- Create one opcode file for each brand/dealer you work with
- Name them clearly so you know which dealer they belong to
- Put all opcode files in the same folder as your Python scripts
- **Note**: Descriptions in column B are optional. If provided, they will be included in the appointment details for better clarity

### Step 3: Configure Dealer Information

The `dealer_info.py` file contains information about each dealer. **This file currently contains sample data only.** You need to replace it with your actual dealer information.

#### Understanding the Dealer Information Structure:

Open `dealer_info.py` and you'll see sample entries like this:
```python
'1': {
    'name': 'Sample Dealer 1',
    'dealer_uuid': 'REPLACE_WITH_YOUR_DEALER_UUID',
    'department_uuid': 'REPLACE_WITH_YOUR_DEPARTMENT_UUID',
    'opcode_xlsx': 'sample_opcodes.xlsx',
    'next_service_interval_in_months': 12,
    'default_nsa_opcode': 'REPLACE_WITH_DEFAULT_NSA_OPCODE',
}
```

#### What Each Field Means:

- **Dealer ID** (`'1'`): DealerID for the Dealer in mykaarma system
- **name**: The dealer's name (for your reference)
- **dealer_uuid**: Unique identifier for the dealer in myKaarma system
- **department_uuid**: Unique identifier for the service department of the dealer
- **opcode_xlsx**: Name of the Excel file containing opcodes for this dealer
- **next_service_interval_in_months**: How many months after service to schedule the next appointment
- **default_nsa_opcode**: A default NSA (Next Service Appointment) opcode that will always be included in every appointment for reporting and identification purposes

#### How to Add a New Dealer:

1. **Get the UUIDs**: You'll need to get the dealer_uuid and department_uuid from myKaarma
2. **Choose a Dealer ID**: Again , you have to get this from mykaarma
3. **Add the entry** to the DEALERS dictionary:

```python
'3': {
    'name': 'Your Dealer Name',
    'dealer_uuid': 'your-dealer-uuid-here',
    'department_uuid': 'your-department-uuid-here',
    'opcode_xlsx': 'your_dealer_opcodes.xlsx',
    'next_service_interval_in_months': 12,
    'default_nsa_opcode': 'your-default-nsa-opcode',
}
```

#### Where to Find UUIDs:
You'll need to get these from your myKaarma administrator or by checking the myKaarma API documentation.

#### Default NSA Opcode:
The `default_nsa_opcode` is a service code that will be automatically included in every Next Service Appointment created by this automation tool. This opcode serves as an identifier for reporting purposes and can be useful for:
- Tracking appointments created by the NSA automation script
- Generating reports on automated vs manual appointments
- Standard NSA-specific services that should always be included
- Dealer-specific NSA procedures

**Note**: The default NSA opcode will be added to the service list along with any opcodes from the repair order. If the default NSA opcode is already present in the RO opcodes, it won't be duplicated. This opcode can be used in reporting to identify which appointments were created through this automation system.

## How to Use

### Part 1: Extract Closed ROs

This step finds all repair orders that were closed on a specific date and filters them based on your opcodes.

1. **Make sure your virtual environment is active** (you should see `(venv)` in your command prompt)

2. **Run the extraction script**:
   ```bash
   python extract_closed_ros_to_xlsx.py
   ```

3. **Follow the prompts**:
   - **Select a dealer**: You'll see a numbered list of dealers. Type the number and press Enter
   - **Choose extraction method**: Select between single date or date range
     - **Single date**: Extract ROs closed on a specific date (e.g., `2024-03-15`)
     - **Date range**: Extract ROs closed within a date range (e.g., from `2024-03-10` to `2024-03-15`)
   - **Enter date(s)**: Type the date(s) in YYYY-MM-DD format
   - **Handle existing files**: If `closed_ros.xlsx` already exists, you can choose to keep existing data or start fresh

4. **Wait for completion**: The script will:
   - Fetch all closed ROs for the specified date(s)
   - Check each RO for your specified opcodes
   - Save matching ROs to `closed_ros.xlsx`

5. **Check the results**: A file called `closed_ros.xlsx` will be created with all matching repair orders

### Optional: Review and Filter Results

Before proceeding to Part 2, you may want to review the extracted repair orders and exclude certain ones from appointment creation.

1. **Open `closed_ros.xlsx`** in Excel or any spreadsheet program
2. **Review the data**: Look through the repair orders that were extracted
3. **Delete unwanted rows**: If you want to exclude certain repair orders from having appointments created:
   - Simply delete the entire row for any RO you don't want to schedule
   - For example, if a customer recently had service or shouldn't be contacted
   - **Important**: Do NOT delete the header row (the first row with column names)
4. **Save the file**: Save your changes to `closed_ros.xlsx`

**Note**: Only the repair orders remaining in the file will have appointments created for them in Part 2.

### Part 2: Schedule Appointments

This step takes the extracted ROs and creates follow-up appointments.

1. **Make sure you have `closed_ros.xlsx`** from Part 1

2. **Run the scheduling script**:
   ```bash
   python schedule_appointments_from_extracted_ros.py
   ```

3. **Duplicate Detection**: The script will:
   - Check for RO numbers that already have appointments created
   - If duplicates are found, prompt you to choose whether to create new appointments
   - Display details about existing appointments (customer name, creation date, appointment UUID)

4. **Wait for completion**: The script will:
   - Read each RO from `closed_ros.xlsx`
   - Calculate when the next service should be (based on months configured)
   - Add the dealer's default NSA opcode to the service list (for reporting identification)
   - Find available appointment slots
   - Create appointments in myKaarma with both RO opcodes and default NSA opcode
   - Track created appointments in a cache file (`appointment_cache.json`)
   - Save results to `schedule_results.xlsx`

5. **Check the results**: A file called `schedule_results.xlsx` will be created showing:
   - Which appointments were successfully created
   - Which ones failed and why
   - Appointment dates and times

### Appointment Tracking

The tool automatically maintains a cache file (`appointment_cache.json`) that tracks:
- RO Numbers that have had appointments created
- Customer information (first name, last name)
- Dealer ID
- Creation date and time
- Appointment UUID

This prevents accidental duplicate appointments and provides visibility into what has already been processed. The cache file is automatically managed by the script and helps ensure data integrity across multiple runs.

## Notifications & Templates

The tool automatically sends text and email notifications to customers when appointments are created. Both notification types use customizable XML-based templates with support for:

- **Dynamic Variables**: Customer names, dealer information, appointment dates/times
- **Automatic Date Formatting**: Converts dates to user-friendly formats (e.g., "Monday, January 15, 2024")
- **Consistent Formatting**: Both templates follow the same XML structure for easy maintenance

### Template Customization

If you want to customize the notification messages sent to customers, modify the template files in the `templates/` directory:
- `templates/email_template.txt` - Email notification template
- `templates/text_template.txt` - Text message template

**📖 For complete template configuration instructions, variable reference, date formatting options, and examples, see the [Template Configuration Guide](TEMPLATE_GUIDE.md).** This guide covers:
- How to use available variables like customer names and appointment details
- Date and time formatting patterns
- Template structure and XML requirements
- Common customization examples
- Troubleshooting template issues

## File Structure

```
nsa-automation/
├── extract_closed_ros_to_xlsx.py      # Script to extract closed ROs
├── schedule_appointments_from_extracted_ros.py  # Script to schedule appointments
├── dealer_info.py                     # Dealer configuration (sample data provided)
├── requirements.txt                   # Python packages needed
├── .env                              # Your credentials (you create this)
├── .gitignore                        # Files to ignore in version control
├── README.md                         # This file
├── sample_opcodes.xlsx               # Sample opcode file (provided)
├── templates/                        # Notification templates
│   ├── email_template.txt           # Email notification template
│   └── text_template.txt            # Text notification template
├── venv/                             # Virtual environment (created automatically)
├── [dealer]_opcodes.xlsx             # Your opcode files (you create these)
├── closed_ros.xlsx                   # Generated by extract script (ignored by git)
├── schedule_results.xlsx             # Generated by schedule script (ignored by git)
└── appointment_cache.json            # Appointment tracking cache (ignored by git)
```

### What Each File Does:

#### Files You Run:
- **`extract_closed_ros_to_xlsx.py`**: Extracts repair orders from myKaarma
- **`schedule_appointments_from_extracted_ros.py`**: Creates appointments for customers

#### Configuration Files:
- **`dealer_info.py`**: Contains dealer information and settings
- **`.env`**: Contains your myKaarma credentials
- **`requirements.txt`**: Lists all Python packages needed

#### Documentation Files:
- **`README.md`**: Main documentation and setup guide

#### Excel Files:
- **`sample_opcodes.xlsx`**: Sample opcode file provided for reference
- **`[dealer]_opcodes.xlsx`**: Your dealer-specific opcode files (you create these)
- **`closed_ros.xlsx`**: Contains extracted repair orders (generated, ignored by git)
- **`schedule_results.xlsx`**: Contains appointment scheduling results (generated, ignored by git)

#### Template Files:
- **`templates/email_template.txt`**: XML-based email notification template
- **`templates/text_template.txt`**: XML-based text notification template

#### Cache Files:
- **`appointment_cache.json`**: Tracks created appointments using RO numbers to prevent duplicates (generated, ignored by git)

## Support

### Before Asking for Help

1. **Read this README completely**
2. **Check the Troubleshooting section**
3. **Make sure you followed all setup steps**
4. **Try running the scripts again**

### When Asking for Help

Please provide:
1. **Your operating system** (Windows, Mac, Linux)
2. **The exact error message** (copy and paste it)
3. **What step you were on** when the error occurred
4. **Whether you've successfully run the scripts before**

### Maintenance

#### Updating Opcodes
- To change which opcodes to track, edit the appropriate Excel files
- Make sure to save as `.xlsx` format

#### Adding New Dealers
- Add new entries to `dealer_info.py`
- Create corresponding opcode Excel files
- Make sure UUIDs are correct

#### Updating Credentials
- Edit the `.env` file if credentials change
- Restart the scripts after making changes

---

**Remember**: This tool connects to myKaarma's live system. Always test with a small set of data first to make sure everything works as expected before processing large amounts of data. 
