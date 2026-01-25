"""
KPI Sync Azure Function
Automated KPI processing for Melbourne Mobile Physio

Triggers:
- Timer: Runs daily at 6 AM AEST (Mon-Fri)
- HTTP: Manual trigger for on-demand processing

Environment Variables Required:
- TENANT_ID: Azure AD tenant ID
- CLIENT_ID: App registration client ID
- CLIENT_SECRET: App registration client secret
- DRIVE_ID: SharePoint drive ID
- SHAREPOINT_SITE_ID: SharePoint site ID (optional, has default)
"""

import azure.functions as func
import logging
import json
import os
import re
from datetime import datetime, timedelta, date
from io import BytesIO
import requests
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Border, Side
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.utils import get_column_letter

# Initialize the Function App
app = func.FunctionApp()

# Cache for resolved file IDs
_file_id_cache = {}
_cache_expiry = None

# =============================================================================
# CONFIGURATION
# =============================================================================

# SharePoint IDs - use environment variables with fallbacks
SHAREPOINT_SITE_ID = os.environ.get(
    'SHAREPOINT_SITE_ID', 
    'melbournemobilephysio.sharepoint.com,d1b97ab0-f5fb-43a4-95bc-f395059bbbdb,17d2e0fa-4a08-455e-a327-ae65010bb70a'
)
DRIVE_ID = os.environ.get('DRIVE_ID', '')

# File paths in SharePoint
CONFIG_FILE_PATH = '/Excel files/KPI Files/KPI/config/KPI_Config_Tables_v4.xlsx'
TEAM_LEADER_FILE_PATH = '/Excel files/KPI Files/KPI/Team_Leader_2026.xlsx'

# Template file paths
TEMPLATE_PATHS = {
    'Physio': '/Excel files/KPI Files/KPI/templates/Template_Physio.xlsx',
    'OT': '/Excel files/KPI Files/KPI/templates/Template_OT.xlsx'
}

CONFIG_SHEETS = {
    'therapists': 'Config_Therapists',
    'teams': 'Config_Teams',
    'thresholds_physio': 'Config_Thresholds_Physio',
    'thresholds_ot': 'Config_Thresholds_OT',
    'kpis': 'Config_KPIs',
    'colours': 'Config_Colours',
    'competency_history': 'Config_Competency_History'
}

# Default colours (will be overridden by config)
DEFAULT_COLORS = {
    'red': 'FFE47373',
    'amber': 'FFFFB74D', 
    'green': 'FF81C784',
    'grey': 'FFC0C0C0',
    'white': 'FFFFFFFF'
}

# Month columns for Jan-Dec structure
MONTH_COLUMNS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'June', 'July', 'Aug', 'Sept', 'Oct', 'Nov', 'Dec']

# Month name to number mapping
MONTH_NAME_TO_NUM = {
    'Jan': 1, 'Feb': 2, 'Mar': 3, 'Apr': 4, 'May': 5, 'June': 6,
    'July': 7, 'Aug': 8, 'Sept': 9, 'Oct': 10, 'Nov': 11, 'Dec': 12
}

# Table configurations for Team Leader structure
PHYSIO_NORTH_TABLES = {
    'Billings_North': 'BillingsKPI',
    'Ceased_North': 'Ceased Services',
    'Documentation_North': 'Documentation',
    'Admin_North': 'Admin',
    'Attitude_North': 'Attitude'
}

PHYSIO_SOUTH_TABLES = {
    'Billings_South': 'BillingsKPI',
    'Ceased_South': 'Ceased Services',
    'Documentation_South': 'Documentation',
    'Admin_South': 'Admin',
    'Attitude_South': 'Attitude'
}

OT_TABLES = {
    'Billings_OT': 'BillingsKPI',
    'Compliance_OT': 'Compliance',
    'ReferrerEng_OT': 'Referrer Engagement',
    'Capacity_OT': 'Capacity',
    'Attitude_OT': 'Attitude'
}

SHEET_CONFIG = {
    'KPI Dashboard North': {
        'team_name': 'Physio_North',
        'tables': PHYSIO_NORTH_TABLES
    },
    'KPI Dashboard South': {
        'team_name': 'Physio_South',
        'tables': PHYSIO_SOUTH_TABLES
    },
    'KPI Dashboard OT': {
        'team_name': 'OT',
        'tables': OT_TABLES
    }
}


# =============================================================================
# GRAPH API HELPERS
# =============================================================================

def get_access_token():
    """Get access token using app registration (client credentials)."""
    # Use app registration credentials (not Managed Identity)
    # Managed Identity is only used for Key Vault access, not Graph API
    tenant_id = os.environ.get('TENANT_ID') or os.environ.get('AZURE_TENANT_ID')
    client_id = os.environ.get('CLIENT_ID') or os.environ.get('AZURE_CLIENT_ID')
    client_secret = os.environ.get('CLIENT_SECRET') or os.environ.get('AZURE_CLIENT_SECRET')
    
    if not all([tenant_id, client_id, client_secret]):
        logging.error(f"Missing credentials - TENANT_ID: {'SET' if tenant_id else 'MISSING'}, CLIENT_ID: {'SET' if client_id else 'MISSING'}, CLIENT_SECRET: {'SET' if client_secret else 'MISSING'}")
        raise Exception("Missing authentication credentials - check environment variables")
    
    token_url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        'grant_type': 'client_credentials',
        'client_id': client_id,
        'client_secret': client_secret,
        'scope': 'https://graph.microsoft.com/.default'
    }
    response = requests.post(token_url, data=data)
    
    if response.status_code == 200:
        logging.info("Using app registration authentication")
        return response.json()['access_token']
    else:
        logging.error(f"Token request failed: {response.status_code} - {response.text}")
        raise Exception(f"Failed to obtain access token: {response.status_code}")


def graph_request(endpoint, token, method='GET', data=None, content_type='application/json'):
    """Make a request to Microsoft Graph API."""
    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': content_type
    }
    
    url = f"https://graph.microsoft.com/v1.0{endpoint}"
    
    if method == 'GET':
        response = requests.get(url, headers=headers)
    elif method == 'PUT':
        response = requests.put(url, headers=headers, data=data)
    elif method == 'POST':
        response = requests.post(url, headers=headers, 
                                json=data if content_type == 'application/json' else None, 
                                data=data if content_type != 'application/json' else None)
    elif method == 'PATCH':
        response = requests.patch(url, headers=headers, json=data)
    
    return response


def resolve_file_path(file_path, token):
    """Resolve a SharePoint file path to a file ID, with caching."""
    global _file_id_cache, _cache_expiry
    
    now = datetime.utcnow()
    if _cache_expiry and now < _cache_expiry and file_path in _file_id_cache:
        return _file_id_cache[file_path]
    
    if not _cache_expiry or now >= _cache_expiry:
        _file_id_cache = {}
        _cache_expiry = now + timedelta(hours=24)
    
    encoded_path = file_path.replace(' ', '%20').replace('&', '%26')
    endpoint = f"/drives/{DRIVE_ID}/root:{encoded_path}"
    
    response = graph_request(endpoint, token)
    
    if response.status_code == 404:
        return None  # File doesn't exist
    
    if response.status_code >= 400:
        raise Exception(f"Graph API error: {response.status_code} - {response.text}")
    
    file_info = response.json()
    file_id = file_info['id']
    
    _file_id_cache[file_path] = file_id
    return file_id


def file_exists(file_path, token):
    """Check if a file exists at the given path."""
    file_id = resolve_file_path(file_path, token)
    return file_id is not None


def download_excel_file(file_id, token):
    """Download an Excel file from SharePoint."""
    endpoint = f"/drives/{DRIVE_ID}/items/{file_id}/content"
    response = graph_request(endpoint, token)
    if response.status_code >= 400:
        raise Exception(f"Failed to download file: {response.status_code}")
    return BytesIO(response.content)


def upload_excel_file(file_path, file_content, token):
    """Upload an Excel file to SharePoint (creates or overwrites)."""
    encoded_path = file_path.replace(' ', '%20').replace('&', '%26')
    endpoint = f"/drives/{DRIVE_ID}/root:{encoded_path}:/content"
    
    if isinstance(file_content, BytesIO):
        file_content.seek(0)
        data = file_content.read()
    else:
        data = file_content
    
    response = graph_request(endpoint, token, method='PUT', data=data,
                            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    
    if response.status_code >= 400:
        raise Exception(f"Failed to upload file: {response.status_code} - {response.text}")
    
    return response


def create_folder(folder_path, token):
    """Create a folder in SharePoint if it doesn't exist."""
    encoded_path = folder_path.replace(' ', '%20').replace('&', '%26')
    endpoint = f"/drives/{DRIVE_ID}/root:{encoded_path}"
    response = graph_request(endpoint, token)
    
    if response.status_code == 200:
        return response.json()['id']
    
    parent_path = '/'.join(folder_path.rsplit('/', 1)[:-1]) or '/'
    folder_name = folder_path.rsplit('/', 1)[-1]
    
    encoded_parent = parent_path.replace(' ', '%20').replace('&', '%26')
    endpoint = f"/drives/{DRIVE_ID}/root:{encoded_parent}:/children"
    
    data = {
        'name': folder_name,
        'folder': {},
        '@microsoft.graph.conflictBehavior': 'fail'
    }
    
    response = graph_request(endpoint, token, method='POST', data=data)
    
    if response.status_code >= 400:
        raise Exception(f"Failed to create folder: {response.status_code}")
    
    return response.json()['id']


# =============================================================================
# CONFIG LOADING
# =============================================================================

def load_config(token):
    """Load all configuration from the master Excel file."""
    logging.info("Loading configuration...")
    
    file_id = resolve_file_path(CONFIG_FILE_PATH, token)
    if not file_id:
        raise Exception(f"Config file not found: {CONFIG_FILE_PATH}")
    
    file_content = download_excel_file(file_id, token)
    wb = load_workbook(file_content, data_only=True)
    config = {}
    
    # Load therapists
    if CONFIG_SHEETS['therapists'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['therapists']]
        config['therapists'] = []
        headers = [cell.value for cell in ws[1]]
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                therapist = dict(zip(headers, row))
                therapist['IsActive'] = str(therapist.get('IsActive', 'TRUE')).upper() == 'TRUE'
                is_leader = therapist.get('IsTeamLeader', False)
                if isinstance(is_leader, str):
                    therapist['IsTeamLeader'] = is_leader.upper() == 'TRUE'
                else:
                    therapist['IsTeamLeader'] = bool(is_leader)
                config['therapists'].append(therapist)
        
        logging.info(f"Loaded {len(config['therapists'])} therapists")
    
    # Load teams
    if CONFIG_SHEETS['teams'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['teams']]
        config['teams'] = {}
        headers = [cell.value for cell in ws[1]]
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                team = dict(zip(headers, row))
                config['teams'][team['TeamId']] = team
        
        logging.info(f"Loaded {len(config['teams'])} teams")
    
    # Load thresholds
    config['thresholds'] = {'Physio': {}, 'OT': {}}
    
    for team_type, sheet_key in [('Physio', 'thresholds_physio'), ('OT', 'thresholds_ot')]:
        if CONFIG_SHEETS[sheet_key] in wb.sheetnames:
            ws = wb[CONFIG_SHEETS[sheet_key]]
            
            for row in ws.iter_rows(min_row=2, max_row=5, values_only=True):
                if row[0] and row[0] in ['Grad', 'CA', 'Senior', 'Team Average']:
                    config['thresholds'][team_type][row[0]] = {
                        'red_below': row[1],
                        'green_min': row[2],
                        'green_max': row[3],
                        'blue_above': row[4]
                    }
            
            logging.info(f"Loaded {team_type} thresholds: {list(config['thresholds'][team_type].keys())}")
    
    # Load ceased service thresholds
    config['ceased_thresholds'] = {}
    if CONFIG_SHEETS['thresholds_physio'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['thresholds_physio']]
        for row in ws.iter_rows(min_row=8, max_row=9, values_only=True):
            if row[0] and 'Ceased %' in str(row[0]):
                config['ceased_thresholds'] = {
                    'blue_below': row[1],
                    'green_min': row[2],
                    'green_max': row[3],
                    'red_above': row[4]
                }
                logging.info(f"Loaded ceased thresholds: {config['ceased_thresholds']}")
                break
        
        if not config['ceased_thresholds']:
            logging.warning("No ceased thresholds found - using defaults")
            config['ceased_thresholds'] = {
                'blue_below': 0.025,
                'green_min': 0.025,
                'green_max': 0.04,
                'red_above': 0.04
            }
    else:
        config['ceased_thresholds'] = {
            'blue_below': 0.025,
            'green_min': 0.025,
            'green_max': 0.04,
            'red_above': 0.04
        }
    
    # Load 1-5 rating scale thresholds
    config['rating_thresholds'] = []
    if CONFIG_SHEETS['thresholds_physio'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['thresholds_physio']]
        for row in ws.iter_rows(min_row=13, max_row=17, values_only=True):
            if row[0] and isinstance(row[0], (int, float)) and 1 <= row[0] <= 5:
                config['rating_thresholds'].append({
                    'rating': row[0],
                    'min': row[1],
                    'max': row[2],
                    'label': row[3] if len(row) > 3 else ''
                })
        
        if config['rating_thresholds']:
            logging.info(f"Loaded {len(config['rating_thresholds'])} rating thresholds")
        else:
            logging.warning("No rating thresholds found - using defaults")
            config['rating_thresholds'] = [
                {'rating': 1, 'min': 1, 'max': 1.499, 'label': 'Unsatisfactory'},
                {'rating': 2, 'min': 1.5, 'max': 2.499, 'label': 'Needs Improvement'},
                {'rating': 3, 'min': 2.5, 'max': 3.499, 'label': 'In Progress'},
                {'rating': 4, 'min': 3.5, 'max': 4.499, 'label': 'Very Good'},
                {'rating': 5, 'min': 4.5, 'max': 5, 'label': 'Excellent'}
            ]
    
    # Load competency history
    config['competency_history'] = []
    if CONFIG_SHEETS['competency_history'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['competency_history']]
        headers = [cell.value for cell in ws[1]]
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                record = dict(zip(headers, row))
                effective_date = record.get('EffectiveDate')
                if effective_date:
                    if hasattr(effective_date, 'date'):
                        record['EffectiveDate'] = effective_date.date()
                    elif not isinstance(effective_date, date):
                        try:
                            record['EffectiveDate'] = datetime.strptime(str(effective_date), '%Y-%m-%d').date()
                        except:
                            logging.warning(f"Could not parse date for {record.get('Name')}: {effective_date}")
                            continue
                    config['competency_history'].append(record)
        
        logging.info(f"Loaded {len(config['competency_history'])} competency history records")
    
    # Load colours
    if CONFIG_SHEETS['colours'] in wb.sheetnames:
        ws = wb[CONFIG_SHEETS['colours']]
        config['colours'] = {}
        
        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[0]:
                config['colours'][row[0]] = row[1]
    else:
        config['colours'] = DEFAULT_COLORS
    
    wb.close()
    return config


# =============================================================================
# COMPETENCY HISTORY HELPERS
# =============================================================================

def extract_year_from_filename(file_path):
    """Extract year from filename like 'Team_Leader_2026.xlsx'."""
    filename = file_path.split('/')[-1]
    match = re.search(r'(\d{4})', filename)
    if match:
        year = int(match.group(1))
        if 2020 <= year <= 2100:
            logging.info(f"Extracted year {year} from filename: {filename}")
            return year
    current_year = datetime.now().year
    logging.warning(f"Could not extract year from '{filename}' - using current year {current_year}")
    return current_year


def get_competency_for_month(therapist_name, month_name, year, config):
    """Get the competency that was active for a therapist in a specific month."""
    month_num = MONTH_NAME_TO_NUM.get(month_name)
    if not month_num:
        logging.warning(f"Unknown month name: {month_name}")
        return None
    
    target_date = date(year, month_num, 15)
    
    history = config.get('competency_history', [])
    therapist_records = [
        r for r in history 
        if r.get('Name', '').strip().lower() == therapist_name.strip().lower()
    ]
    
    if therapist_records:
        therapist_records.sort(key=lambda r: r.get('EffectiveDate', date.min), reverse=True)
        
        for record in therapist_records:
            effective_date = record.get('EffectiveDate')
            if effective_date and effective_date <= target_date:
                return record.get('Competency')
    
    for therapist in config.get('therapists', []):
        if therapist.get('Name', '').strip().lower() == therapist_name.strip().lower():
            return therapist.get('Competency')
    
    return None


# =============================================================================
# KPI DATA LOADING
# =============================================================================

def read_table_data(ws, table_name, kpi_column_name):
    """Read data from one Excel table."""
    result = {}
    
    try:
        if table_name not in ws.tables:
            logging.warning(f"Table '{table_name}' not found in worksheet '{ws.title}'")
            return result
        
        table = ws.tables[table_name]
        table_range = table.ref
        
        from openpyxl.utils import range_boundaries
        min_col, min_row, max_col, max_row = range_boundaries(table_range)
        
        header_row = []
        for col in range(min_col, max_col + 1):
            cell_value = ws.cell(row=min_row, column=col).value
            header_row.append(cell_value)
        
        month_indices = {}
        for i, header in enumerate(header_row):
            if header in MONTH_COLUMNS:
                month_indices[header] = i
        
        for row_idx in range(min_row + 1, max_row + 1):
            name_cell = ws.cell(row=row_idx, column=min_col)
            therapist_name = name_cell.value
            
            if not therapist_name:
                continue
            
            therapist_name = str(therapist_name).strip()
            
            if therapist_name not in result:
                result[therapist_name] = {}
            
            for month in MONTH_COLUMNS:
                if month in month_indices:
                    col_offset = month_indices[month]
                    cell = ws.cell(row=row_idx, column=min_col + col_offset)
                    value = cell.value
                    result[therapist_name][month] = value
        
        logging.info(f"Read {len(result)} therapists from table '{table_name}'")
        
    except Exception as e:
        logging.error(f"Error reading table '{table_name}': {str(e)}")
    
    return result


def process_dashboard_sheet(ws, team_name, table_config):
    """Extract all KPI data from one dashboard sheet."""
    logging.info(f"Processing sheet '{ws.title}' for team '{team_name}'")
    
    kpi_data = {}
    for table_name, kpi_column_name in table_config.items():
        table_data = read_table_data(ws, table_name, kpi_column_name)
        kpi_data[kpi_column_name] = table_data
    
    therapist_data = {}
    
    all_therapists = set()
    for kpi_name, therapist_dict in kpi_data.items():
        all_therapists.update(therapist_dict.keys())
    
    for therapist_name in all_therapists:
        therapist_data[therapist_name] = {}
        
        for month in MONTH_COLUMNS:
            therapist_data[therapist_name][month] = {}
            
            for kpi_name, therapist_dict in kpi_data.items():
                if therapist_name in therapist_dict and month in therapist_dict[therapist_name]:
                    value = therapist_dict[therapist_name][month]
                    therapist_data[therapist_name][month][kpi_name] = value
                else:
                    therapist_data[therapist_name][month][kpi_name] = None
    
    logging.info(f"Processed {len(therapist_data)} therapists from '{ws.title}'")
    return therapist_data


def transform_to_monthly_records(therapist_data, team_name):
    """Transform therapist-centric data into monthly records."""
    records = []
    
    for therapist_name, months_data in therapist_data.items():
        for month, kpi_values in months_data.items():
            record = {
                'Name': therapist_name,
                'Month': month
            }
            record.update(kpi_values)
            records.append(record)
    
    logging.info(f"Transformed {len(records)} monthly records for team '{team_name}'")
    return records


def load_kpi_dashboard_data(wb):
    """Load KPI data from Team Leader Dashboard tables."""
    logging.info("Loading KPI data from Dashboard tables...")
    
    master_data = {}
    
    for sheet_name, config in SHEET_CONFIG.items():
        team_name = config['team_name']
        table_config = config['tables']
        
        if sheet_name not in wb.sheetnames:
            logging.warning(f"Sheet '{sheet_name}' not found in workbook")
            master_data[team_name] = []
            continue
        
        ws = wb[sheet_name]
        therapist_data = process_dashboard_sheet(ws, team_name, table_config)
        records = transform_to_monthly_records(therapist_data, team_name)
        master_data[team_name] = records
    
    total_records = sum(len(records) for records in master_data.values())
    logging.info(f"Loaded {total_records} total records from Dashboard tables")
    for team_name, records in master_data.items():
        logging.info(f"  {team_name}: {len(records)} records")
    
    return master_data


# =============================================================================
# MAIN PROCESSING FUNCTIONS
# =============================================================================

def process_kpi_sync(process_individual=True, process_team_leader=True, therapist_filter=None):
    """
    Main KPI sync function.
    
    Args:
        process_individual: Whether to update individual therapist sheets
        process_team_leader: Whether to sync/format Team Leader file
        therapist_filter: Optional name to filter to single therapist
        
    Returns:
        dict: Processing statistics
    """
    # Import supporting modules
    from individual_sheet_v2 import update_individual_sheet
    from team_table_sync import sync_all_team_tables
    from team_leader_formatting import format_all_team_leader_sheets
    
    stats = {
        'status': 'success',
        'individual': {'success': 0, 'failed': 0, 'skipped': 0},
        'team_leader': {'synced': False, 'formatted': False}
    }
    
    try:
        # Get authentication token
        logging.info("Getting access token...")
        token = get_access_token()
        logging.info("Token obtained successfully")
        
        # Load configuration
        logging.info("Loading configuration...")
        config = load_config(token)
        therapists = config.get('therapists', [])
        active_therapists = [t for t in therapists if t.get('IsActive', True)]
        
        # Apply filter if specified
        if therapist_filter:
            active_therapists = [t for t in active_therapists 
                                if therapist_filter.lower() in t.get('Name', '').lower()]
        
        logging.info(f"Config loaded: {len(active_therapists)} active therapists")
        
        # Load master data from Team Leader file
        logging.info("Loading master KPI data from Team Leader file...")
        file_id = resolve_file_path(TEAM_LEADER_FILE_PATH, token)
        if not file_id:
            raise Exception(f"Team Leader file not found: {TEAM_LEADER_FILE_PATH}")
        
        file_content = download_excel_file(file_id, token)
        wb = load_workbook(file_content, data_only=True)
        master_data = load_kpi_dashboard_data(wb)
        wb.close()
        
        total_records = sum(len(records) for records in master_data.values())
        logging.info(f"Master data loaded: {total_records} records")
        
        # Extract year from Team Leader filename for competency history
        year = extract_year_from_filename(TEAM_LEADER_FILE_PATH)
        logging.info(f"Year for competency history: {year}")
        
        # Process individual sheets
        if process_individual:
            logging.info("Processing individual therapist sheets...")
            
            for i, therapist in enumerate(active_therapists, 1):
                name = therapist.get('Name', 'Unknown')
                file_path = therapist.get('FilePath', '')
                
                logging.info(f"[{i}/{len(active_therapists)}] Processing {name}")
                
                if not file_path:
                    logging.warning(f"No FilePath for {name} - skipping")
                    stats['individual']['skipped'] += 1
                    continue
                
                try:
                    result = update_individual_sheet(therapist, config, master_data, token, year)
                    if result:
                        stats['individual']['success'] += 1
                    else:
                        stats['individual']['failed'] += 1
                except Exception as e:
                    logging.error(f"Error processing {name}: {e}")
                    stats['individual']['failed'] += 1
        
        # Process Team Leader file
        if process_team_leader:
            logging.info("Processing Team Leader file...")
            
            # Download fresh copy
            file_content = download_excel_file(file_id, token)
            wb = load_workbook(file_content)
            
            # Sync tables
            sync_stats = sync_all_team_tables(wb, config, token)
            stats['team_leader']['synced'] = True
            logging.info(f"Sync complete: {sync_stats['tables']} tables")
            
            # Apply formatting
            format_stats = format_all_team_leader_sheets(wb, config, TEAM_LEADER_FILE_PATH)
            stats['team_leader']['formatted'] = True
            logging.info(f"Formatting complete: {format_stats['rows_formatted']} rows")
            
            # Upload
            output = BytesIO()
            wb.save(output)
            output.seek(0)
            wb.close()
            
            upload_excel_file(TEAM_LEADER_FILE_PATH, output, token)
            logging.info("Team Leader file uploaded successfully")
        
        logging.info("KPI sync completed successfully")
        
    except Exception as e:
        logging.error(f"KPI sync failed: {str(e)}")
        stats['status'] = 'error'
        stats['error'] = str(e)
    
    return stats


# =============================================================================
# AZURE FUNCTION TRIGGERS
# =============================================================================

@app.timer_trigger(schedule="0 */30 * * * *", arg_name="mytimer", run_on_startup=False)
def kpi_sync_timer(mytimer: func.TimerRequest) -> None:
    """
    Timer-triggered KPI sync function.
    Runs every 30 minutes.
    
    Cron format: second minute hour day-of-month month day-of-week
    0 */30 * * * * = Every 30 minutes (at :00 and :30)
    """
    utc_timestamp = datetime.utcnow().isoformat()
    
    if mytimer.past_due:
        logging.info('Timer is past due!')
    
    logging.info(f'KPI sync timer trigger started at {utc_timestamp}')
    
    try:
        stats = process_kpi_sync(process_individual=True, process_team_leader=True)
        logging.info(f"KPI sync completed: {stats}")
    except Exception as e:
        logging.error(f"KPI sync timer failed: {str(e)}")
        raise


@app.route(route="kpi_sync", methods=["POST", "GET"], auth_level=func.AuthLevel.FUNCTION)
def kpi_sync_http(req: func.HttpRequest) -> func.HttpResponse:
    """
    HTTP-triggered KPI sync function for manual/on-demand processing.
    
    Query parameters or JSON body:
    - process_individual: true/false (default: true)
    - process_team_leader: true/false (default: true)
    - therapist: Optional therapist name filter
    
    Example:
        POST /api/kpi_sync
        {"therapist": "Chris", "process_team_leader": false}
    """
    logging.info('KPI sync HTTP trigger received')
    
    # Parse parameters from query string or JSON body
    process_individual = True
    process_team_leader = True
    therapist_filter = None
    
    # Try to get from query params first
    if req.params.get('process_individual'):
        process_individual = req.params.get('process_individual', '').lower() == 'true'
    if req.params.get('process_team_leader'):
        process_team_leader = req.params.get('process_team_leader', '').lower() == 'true'
    if req.params.get('therapist'):
        therapist_filter = req.params.get('therapist')
    
    # Try to get from JSON body
    try:
        req_body = req.get_json()
        if req_body:
            if 'process_individual' in req_body:
                process_individual = req_body['process_individual']
            if 'process_team_leader' in req_body:
                process_team_leader = req_body['process_team_leader']
            if 'therapist' in req_body:
                therapist_filter = req_body['therapist']
    except ValueError:
        pass  # No JSON body, use query params
    
    try:
        stats = process_kpi_sync(
            process_individual=process_individual,
            process_team_leader=process_team_leader,
            therapist_filter=therapist_filter
        )
        
        return func.HttpResponse(
            json.dumps(stats, indent=2),
            status_code=200,
            mimetype="application/json"
        )
        
    except Exception as e:
        logging.error(f"KPI sync HTTP failed: {str(e)}")
        return func.HttpResponse(
            json.dumps({"status": "error", "error": str(e)}),
            status_code=500,
            mimetype="application/json"
        )


@app.route(route="health", methods=["GET"], auth_level=func.AuthLevel.ANONYMOUS)
def health_check(req: func.HttpRequest) -> func.HttpResponse:
    """Simple health check endpoint."""
    return func.HttpResponse(
        json.dumps({
            "status": "healthy",
            "timestamp": datetime.utcnow().isoformat(),
            "version": "2.0.0"
        }),
        status_code=200,
        mimetype="application/json"
    )
