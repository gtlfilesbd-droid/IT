import pandas as pd
import json
from datetime import datetime
from collections import defaultdict

def read_excel_data(file_path, asset_type):
    """Read Excel file and extract asset data."""
    try:
        df = pd.read_excel(file_path)
        assets = []
        
        # Special handling for Server file which has different structure
        if asset_type == 'Server':
            # Server format: S/L, Device Name, Model, Serial, Status, Remarks
            header_row = None
            for idx, row in df.iterrows():
                if 'S/L' in str(row.values) or 'Device Name' in str(row.values):
                    header_row = idx
                    break
            
            if header_row is None:
                print(f"Warning: Could not find header in {file_path}")
                return assets
            
            # Process server data rows
            for idx in range(header_row + 1, len(df)):
                row = df.iloc[idx]
                
                # Get S/L number (column 0)
                sl_num = row.iloc[0] if len(row) > 0 else None
                
                # Skip rows without valid S/L
                if pd.isna(sl_num) or sl_num == '':
                    continue
                
                try:
                    serial_num = int(sl_num)
                except (ValueError, TypeError):
                    continue
                
                asset = {
                    'AssetType': asset_type,
                    'Serial': serial_num,
                }
                
                # Device Name (column 1)
                if len(row) > 1 and not pd.isna(row.iloc[1]):
                    asset['Name'] = str(row.iloc[1]).strip()
                
                # Model (column 2)
                if len(row) > 2 and not pd.isna(row.iloc[2]):
                    asset['Model'] = str(row.iloc[2]).strip()
                
                # Device Serial (column 3) - store as DeviceSerial, not as asset serial
                if len(row) > 3 and not pd.isna(row.iloc[3]):
                    asset['DeviceSerial'] = str(row.iloc[3]).strip()
                
                # Status (column 4)
                if len(row) > 4 and not pd.isna(row.iloc[4]):
                    asset['Status'] = str(row.iloc[4]).strip()
                
                # Remarks (column 5)
                if len(row) > 5 and not pd.isna(row.iloc[5]):
                    asset['Remarks'] = str(row.iloc[5]).strip()
                
                if 'Name' in asset and asset['Name']:
                    assets.append(asset)
            
            return assets
        
        # For non-Server assets, use the original logic
        # Find the row with column headers (contains "Serial")
        header_row = None
        serial_col = None
        for idx, row in df.iterrows():
            # Check if this row contains "Serial" which indicates the header
            for col_idx, val in enumerate(row.values):
                if 'Serial' in str(val):
                    header_row = idx
                    serial_col = col_idx
                    break
            if header_row is not None:
                break
        
        if header_row is None or serial_col is None:
            print(f"Warning: Could not find header in {file_path}")
            return assets
        
        # Determine if this file has a "Category" column (like Laptop.xlsx)
        # or starts directly with Serial (like PC.xlsx, Monitor.xlsx)
        has_category = (serial_col == 1)  # If Serial is in column 1, there's a Category column
        
        # Process data rows after the header
        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Get serial number from the appropriate column
            serial = row.iloc[serial_col] if len(row) > serial_col else None
            
            # Skip rows without a valid serial number
            if pd.isna(serial) or serial == '':
                continue
            
            # Skip section header rows (like "New Purchase Laptop Information")
            try:
                serial_num = int(serial)
            except (ValueError, TypeError):
                continue
            
            # Extract data from each column
            asset = {
                'AssetType': asset_type,
                'Serial': serial_num,
            }
            
            # Column mapping depends on whether there's a category column
            if has_category:
                # Laptop.xlsx format: 0=Category, 1=Serial, 2=Name, 3=Model, 4=RAM, 5=Processor, etc.
                col_offset = 1
            else:
                # PC.xlsx/Monitor.xlsx format: 0=Serial, 1=Name, 2=Model, 3=RAM, 4=Processor, etc.
                col_offset = 0
            
            # Name
            name_col = col_offset + 1
            if len(row) > name_col and not pd.isna(row.iloc[name_col]):
                asset['Name'] = str(row.iloc[name_col]).strip()
            
            # Model
            model_col = col_offset + 2
            if len(row) > model_col and not pd.isna(row.iloc[model_col]):
                asset['Model'] = str(row.iloc[model_col]).strip()
            
            # For Laptop and Desktop PC assets, we have more columns
            if asset_type in ['Laptop', 'Desktop PC', 'Server']:
                # RAM
                ram_col = col_offset + 3
                if len(row) > ram_col and not pd.isna(row.iloc[ram_col]):
                    try:
                        asset['RAM'] = int(row.iloc[ram_col])
                    except (ValueError, TypeError):
                        asset['RAM'] = 0
                
                # Processor
                proc_col = col_offset + 4
                if len(row) > proc_col and not pd.isna(row.iloc[proc_col]):
                    asset['Processor'] = str(row.iloc[proc_col]).strip()
                
                # Storage
                storage_col = col_offset + 5
                if len(row) > storage_col and not pd.isna(row.iloc[storage_col]):
                    asset['Storage'] = str(row.iloc[storage_col]).strip()
                
                # Gen
                gen_col = col_offset + 6
                if len(row) > gen_col and not pd.isna(row.iloc[gen_col]):
                    try:
                        asset['Gen'] = int(row.iloc[gen_col])
                    except (ValueError, TypeError):
                        asset['Gen'] = 0
                
                # GPU
                gpu_col = col_offset + 7
                if len(row) > gpu_col and not pd.isna(row.iloc[gpu_col]):
                    asset['GPU'] = str(row.iloc[gpu_col]).strip()
                
                # User
                user_col = col_offset + 8
                if len(row) > user_col and not pd.isna(row.iloc[user_col]):
                    asset['User'] = str(row.iloc[user_col]).strip()
                
                # Location
                loc_col = col_offset + 9
                if len(row) > loc_col and not pd.isna(row.iloc[loc_col]):
                    asset['Location'] = str(row.iloc[loc_col]).strip()
                
                # Level
                level_col = col_offset + 10
                if len(row) > level_col and not pd.isna(row.iloc[level_col]):
                    asset['Level'] = str(row.iloc[level_col]).strip()
                
                # Status
                status_col = col_offset + 11
                if len(row) > status_col and not pd.isna(row.iloc[status_col]):
                    asset['Status'] = str(row.iloc[status_col]).strip()
                
                # Remarks
                remarks_col = col_offset + 12
                if len(row) > remarks_col and not pd.isna(row.iloc[remarks_col]):
                    asset['Remarks'] = str(row.iloc[remarks_col]).strip()
            
            elif asset_type == 'Monitor':
                # Monitor format: Serial, Name, Model, User, Location, Value Ratio, Status, Remarks
                # User
                user_col = col_offset + 3
                if len(row) > user_col and not pd.isna(row.iloc[user_col]):
                    asset['User'] = str(row.iloc[user_col]).strip()
                
                # Location
                loc_col = col_offset + 4
                if len(row) > loc_col and not pd.isna(row.iloc[loc_col]):
                    asset['Location'] = str(row.iloc[loc_col]).strip()
                
                # Level (Value Ratio)
                level_col = col_offset + 5
                if len(row) > level_col and not pd.isna(row.iloc[level_col]):
                    asset['Level'] = str(row.iloc[level_col]).strip()
                
                # Status
                status_col = col_offset + 6
                if len(row) > status_col and not pd.isna(row.iloc[status_col]):
                    asset['Status'] = str(row.iloc[status_col]).strip()
                
                # Remarks
                remarks_col = col_offset + 7
                if len(row) > remarks_col and not pd.isna(row.iloc[remarks_col]):
                    asset['Remarks'] = str(row.iloc[remarks_col]).strip()
            
            elif asset_type == 'Printer/Scanner':
                # Printer format: Similar to Monitor
                # User
                user_col = col_offset + 3
                if len(row) > user_col and not pd.isna(row.iloc[user_col]):
                    asset['User'] = str(row.iloc[user_col]).strip()
                
                # Location
                loc_col = col_offset + 4
                if len(row) > loc_col and not pd.isna(row.iloc[loc_col]):
                    asset['Location'] = str(row.iloc[loc_col]).strip()
                
                # Level
                level_col = col_offset + 5
                if len(row) > level_col and not pd.isna(row.iloc[level_col]):
                    asset['Level'] = str(row.iloc[level_col]).strip()
                
                # Status
                status_col = col_offset + 6
                if len(row) > status_col and not pd.isna(row.iloc[status_col]):
                    asset['Status'] = str(row.iloc[status_col]).strip()
                
                # Remarks
                remarks_col = col_offset + 7
                if len(row) > remarks_col and not pd.isna(row.iloc[remarks_col]):
                    asset['Remarks'] = str(row.iloc[remarks_col]).strip()
            
            # Only add if we have at least a name
            if 'Name' in asset and asset['Name']:
                assets.append(asset)
        
        return assets
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        import traceback
        traceback.print_exc()
        return []

def calculate_asset_value(asset):
    """Calculate current value of an asset based on type, specs, and condition."""
    base_value = 0
    
    # Base values by type and specs
    if asset['AssetType'] == 'Laptop':
        # Base laptop value
        base_value = 30000
        
        # Processor bonus
        processor = asset.get('Processor', '').lower()
        if 'i7' in processor:
            base_value += 15000
        elif 'i5' in processor:
            base_value += 8000
        elif 'i3' in processor:
            base_value += 3000
        elif 'ryzen 7' in processor:
            base_value += 15000
        elif 'ryzen 5' in processor:
            base_value += 8000
        
        # RAM bonus
        ram = asset.get('RAM', 0)
        base_value += ram * 1500
        
        # Generation bonus
        gen = asset.get('Gen', 0)
        if gen >= 11:
            base_value += (gen - 10) * 3000
        
        # GPU bonus
        gpu = asset.get('GPU', '')
        if '4GB' in str(gpu):
            base_value += 8000
        elif '2GB' in str(gpu):
            base_value += 4000
        
    elif asset['AssetType'] == 'Desktop PC':
        # Base PC value
        base_value = 40000
        
        # Processor bonus
        processor = asset.get('Processor', '').lower()
        if 'i9' in processor:
            base_value += 50000
        elif 'i7' in processor:
            base_value += 20000
        elif 'i5' in processor:
            base_value += 10000
        elif 'i3' in processor:
            base_value += 5000
        
        # RAM bonus
        ram = asset.get('RAM', 0)
        base_value += ram * 2000
        
        # Generation bonus
        gen = asset.get('Gen', 0)
        if gen >= 11:
            base_value += (gen - 10) * 4000
        
        # GPU bonus
        gpu = asset.get('GPU', '')
        if '4GB' in str(gpu):
            base_value += 10000
        elif '2GB' in str(gpu):
            base_value += 5000
        
    elif asset['AssetType'] == 'Monitor':
        # Base monitor value
        base_value = 8000
        
        # Size and brand adjustments
        model = asset.get('Model', '').lower()
        if 'dell' in model or 'hp' in model or 'lg' in model:
            base_value += 2000
        
        # Check for larger sizes in model name
        if '24' in model or '27' in model:
            base_value += 3000
        
    elif asset['AssetType'] == 'Printer/Scanner':
        # Base printer value
        base_value = 15000
        
        model = asset.get('Model', '').lower()
        name = asset.get('Name', '').lower()
        
        if 'hp' in name or 'hp' in model:
            base_value += 5000
        if 'laserjet' in model:
            base_value += 10000
        
    elif asset['AssetType'] == 'Server':
        # Base server/network equipment value
        base_value = 15000
        
        model = asset.get('Model', '').lower()
        name = asset.get('Name', '').lower()
        
        # Higher value for specific equipment types
        if 'router' in name or 'mikrotik' in name:
            base_value = 25000
        elif 'switch' in name:
            if '24 port' in name or '16 port' in name:
                base_value = 20000
            elif '8 port' in name:
                base_value = 12000
        elif 'dvr' in name or 'hikvision' in name:
            if '16 port' in name:
                base_value = 18000
            elif '8 port' in name:
                base_value = 12000
        elif 'ups' in name:
            base_value = 30000
        elif 'kvm' in name:
            base_value = 8000
        elif 'patch panel' in name or 'cable' in name:
            base_value = 3000
        elif 'pbax' in name or 'controller' in name:
            base_value = 15000
        elif 'acs' in name or 'converter' in name:
            base_value = 10000
    
    # Depreciation based on condition/remarks
    # All assets are used, so minimum 30% depreciation
    remarks = asset.get('Remarks', '').lower()
    status = asset.get('Status', '').lower()
    
    if 'excellent' in remarks:
        depreciation = 0.70  # 30% depreciation for excellent used condition
    elif 'good' in remarks or 'working' in status:
        depreciation = 0.60  # 40% depreciation for good used condition
    elif 'fair' in remarks or 'moderate' in remarks:
        depreciation = 0.50  # 50% depreciation for fair/moderate used condition
    else:
        depreciation = 0.60  # 40% depreciation as default
    
    final_value = base_value * depreciation
    
    return round(final_value, 2)

def divide_into_groups(assets):
    """Divide assets into 3 equal groups by value using greedy algorithm."""
    # Calculate values for all assets
    for asset in assets:
        asset['CurrentValue'] = calculate_asset_value(asset)
    
    # Sort assets by value (descending)
    sorted_assets = sorted(assets, key=lambda x: x['CurrentValue'], reverse=True)
    
    # Initialize groups
    groups = {'A': [], 'B': [], 'C': []}
    group_values = {'A': 0, 'B': 0, 'C': 0}
    group_counts = {'A': defaultdict(int), 'B': defaultdict(int), 'C': defaultdict(int)}
    
    # Assign each asset to the group with lowest total value
    for asset in sorted_assets:
        # Find group with minimum value
        min_group = min(group_values, key=group_values.get)
        
        # Add asset to that group
        groups[min_group].append(asset)
        group_values[min_group] += asset['CurrentValue']
        group_counts[min_group][asset['AssetType']] += 1
        
        # Add allocation remark
        asset['AllocationRemark'] = f"Assigned to balance total value (Group {min_group}: ৳{group_values[min_group]:,.0f})"
    
    return groups, group_values, group_counts

def generate_html(groups, group_values, group_counts, all_assets):
    """Generate HTML report."""
    
    total_value = sum(group_values.values())
    target_value = total_value / 3
    
    # Count asset types
    type_counts = defaultdict(int)
    for asset in all_assets:
        type_counts[asset['AssetType']] += 1
    
    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Genesis Technologies Ltd - Complete Asset Division (Equal Distribution)</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }}
        
        .container {{
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }}
        
        .header p {{
            font-size: 1.1em;
            opacity: 0.9;
        }}
        
        .notice {{
            background: #d4edda;
            border-left: 4px solid #28a745;
            padding: 15px 20px;
            margin: 20px 40px;
            border-radius: 5px;
        }}
        
        .notice strong {{
            color: #155724;
        }}
        
        .summary {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 40px;
            background: #f8f9fa;
        }}
        
        .summary-card {{
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            text-align: center;
            border-top: 4px solid #3498db;
        }}
        
        .summary-card h3 {{
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.1em;
        }}
        
        .summary-card .value {{
            font-size: 1.8em;
            font-weight: bold;
            color: #3498db;
            margin: 10px 0;
        }}
        
        .summary-card .label {{
            color: #7f8c8d;
            font-size: 0.85em;
        }}
        
        .groups-container {{
            padding: 40px;
        }}
        
        .group {{
            margin-bottom: 50px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        
        .group-a {{
            border: 3px solid #e74c3c;
        }}
        
        .group-b {{
            border: 3px solid #3498db;
        }}
        
        .group-c {{
            border: 3px solid #2ecc71;
        }}
        
        .group-header {{
            padding: 20px 30px;
            color: white;
            font-size: 1.5em;
            font-weight: bold;
        }}
        
        .group-a .group-header {{
            background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%);
        }}
        
        .group-b .group-header {{
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
        }}
        
        .group-c .group-header {{
            background: linear-gradient(135deg, #2ecc71 0%, #27ae60 100%);
        }}
        
        .group-stats {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 15px;
            padding: 20px 30px;
            background: #f8f9fa;
        }}
        
        .stat {{
            text-align: center;
        }}
        
        .stat-label {{
            color: #7f8c8d;
            font-size: 0.9em;
            margin-bottom: 5px;
        }}
        
        .stat-value {{
            color: #2c3e50;
            font-size: 1.3em;
            font-weight: bold;
        }}
        
        .asset-table {{
            width: 100%;
            border-collapse: collapse;
        }}
        
        .asset-table th {{
            background: #34495e;
            color: white;
            padding: 12px;
            text-align: left;
            font-weight: 600;
        }}
        
        .asset-table td {{
            padding: 10px 12px;
            border-bottom: 1px solid #ecf0f1;
        }}
        
        .asset-table tr:hover {{
            background: #f8f9fa;
        }}
        
        .asset-type-section {{
            margin: 20px 30px;
        }}
        
        .asset-type-header {{
            background: #34495e;
            color: white;
            padding: 15px 20px;
            margin-top: 20px;
            border-radius: 5px;
            font-size: 1.2em;
            font-weight: bold;
        }}
        
        .value-badge {{
            background: #3498db;
            color: white;
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.9em;
        }}
        
        .level-badge {{
            padding: 4px 8px;
            border-radius: 4px;
            font-size: 0.85em;
        }}
        
        @media print {{
            body {{
                background: white;
                padding: 0;
            }}
            
            .container {{
                box-shadow: none;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ Genesis Technologies Ltd</h1>
            <p>Complete Asset Division Report - Equal Distribution Among Three Groups</p>
            <p style="font-size: 0.9em; margin-top: 10px;">Generated on {datetime.now().strftime('%B %d, %Y at %I:%M %p')}</p>
        </div>
        
        <div class="notice">
            <strong>✅ Updated Asset Division:</strong> This report reflects only the assets currently available for sharing. Assets have been equally distributed among three groups based on total value.
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Total Assets</h3>
                <div class="value">{len(all_assets)}</div>
                <div class="label">Available for Sharing</div>
            </div>
            <div class="summary-card">
                <h3>Laptops</h3>
                <div class="value">{type_counts.get('Laptop', 0)}</div>
                <div class="label">Laptop Computers</div>
            </div>
            <div class="summary-card">
                <h3>Desktop PCs</h3>
                <div class="value">{type_counts.get('Desktop PC', 0)}</div>
                <div class="label">Desktop Computers</div>
            </div>
            <div class="summary-card">
                <h3>Monitors</h3>
                <div class="value">{type_counts.get('Monitor', 0)}</div>
                <div class="label">Display Monitors</div>
            </div>
            <div class="summary-card">
                <h3>Printers</h3>
                <div class="value">{type_counts.get('Printer/Scanner', 0)}</div>
                <div class="label">Printers & Scanners</div>
            </div>
            <div class="summary-card">
                <h3>Server Devices</h3>
                <div class="value">{type_counts.get('Server', 0)}</div>
                <div class="label">Network & Server Equipment</div>
            </div>
            <div class="summary-card">
                <h3>Total Value</h3>
                <div class="value">৳{total_value:,.0f}</div>
                <div class="label">Combined Worth (BDT)</div>
            </div>
            <div class="summary-card">
                <h3>Target per Group</h3>
                <div class="value">৳{target_value:,.0f}</div>
                <div class="label">Equal Distribution Target</div>
            </div>
        </div>

        <div class="groups-container">
"""
    
    # Generate sections for each group
    for group_name in ['A', 'B', 'C']:
        group_assets = groups[group_name]
        group_val = group_values[group_name]
        variance = ((group_val - target_value) / target_value) * 100
        
        # Count assets by type in this group
        type_counts_group = defaultdict(int)
        for asset in group_assets:
            type_counts_group[asset['AssetType']] += 1
        
        html += f"""
            <div class="group group-{group_name.lower()}">
                <div class="group-header">
                    Group {group_name} - Complete Asset Allocation
                </div>
                
                <div class="group-stats">
                    <div class="stat">
                        <div class="stat-label">Total Assets</div>
                        <div class="stat-value">{len(group_assets)}</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Total Value</div>
                        <div class="stat-value">৳{group_val:,.0f}</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Variance</div>
                        <div class="stat-value">{variance:+.2f}%</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Avg per Asset</div>
                        <div class="stat-value">৳{group_val/len(group_assets):,.0f}</div>
                    </div>
"""
        
        # Add type counts
        for asset_type in ['Laptop', 'Desktop PC', 'Monitor', 'Printer/Scanner', 'Server']:
            count = type_counts_group.get(asset_type, 0)
            if count > 0:
                display_name = asset_type.replace('Desktop PC', 'Desktop PC').replace('Printer/Scanner', 'Printer')
                html += f"""
                    <div class="stat">
                        <div class="stat-label">{display_name}</div>
                        <div class="stat-value">{count}</div>
                    </div>
"""
        
        html += """
                </div>
                
                <div class="asset-type-section">
"""
        
        # Group assets by type
        assets_by_type = defaultdict(list)
        for asset in group_assets:
            assets_by_type[asset['AssetType']].append(asset)
        
        # Display each asset type
        for asset_type in ['Laptop', 'Desktop PC', 'Monitor', 'Printer/Scanner', 'Server']:
            if asset_type not in assets_by_type:
                continue
            
            type_assets = sorted(assets_by_type[asset_type], key=lambda x: x['CurrentValue'], reverse=True)
            
            html += f"""
                    <div class="asset-type-header">
                        {asset_type}s ({len(type_assets)}) - Total Value: ৳{sum(a['CurrentValue'] for a in type_assets):,.0f}
                    </div>
                    
                    <table class="asset-table">
                        <thead>
                            <tr>
                                <th>Serial</th>
                                <th>Name</th>
                                <th>Model</th>
                                <th>Specs</th>
                                <th>User/Location</th>
                                <th>Status</th>
                                <th>Value</th>
                            </tr>
                        </thead>
                        <tbody>
"""
            
            for asset in type_assets:
                # Build specs string
                specs = []
                if asset.get('Processor'):
                    specs.append(f"CPU: {asset['Processor']}")
                if asset.get('RAM'):
                    specs.append(f"RAM: {asset['RAM']}GB")
                if asset.get('Storage'):
                    specs.append(f"Storage: {asset['Storage']}")
                if asset.get('GPU'):
                    specs.append(f"GPU: {asset['GPU']}")
                if asset.get('Gen'):
                    specs.append(f"Gen: {asset['Gen']}")
                
                specs_str = '<br>'.join(specs) if specs else 'N/A'
                
                user_loc = f"{asset.get('User', 'N/A')}<br>{asset.get('Location', 'N/A')}"
                status_str = f"{asset.get('Status', 'N/A')}<br><small>{asset.get('Remarks', '')}</small>"
                
                html += f"""
                            <tr>
                                <td><strong>{asset.get('Serial', 'N/A')}</strong></td>
                                <td>{asset.get('Name', 'N/A')}</td>
                                <td>{asset.get('Model', 'N/A')}</td>
                                <td><small>{specs_str}</small></td>
                                <td><small>{user_loc}</small></td>
                                <td><small>{status_str}</small></td>
                                <td><span class="value-badge">৳{asset['CurrentValue']:,.0f}</span></td>
                            </tr>
"""
            
            html += """
                        </tbody>
                    </table>
"""
        
        html += """
                </div>
            </div>
"""
    
    html += """
        </div>
    </div>
</body>
</html>
"""
    
    return html

def main():
    print("Starting asset processing...")
    
    # File mappings
    files = {
        'Laptop.xlsx': 'Laptop',
        'PC.xlsx': 'Desktop PC',
        'Monitor.xlsx': 'Monitor',
        'Printer Scaneer.xlsx': 'Printer/Scanner',
        'Server.xlsx': 'Server'
    }
    
    all_assets = []
    
    # Read all Excel files
    for filename, asset_type in files.items():
        print(f"Reading {filename}...")
        assets = read_excel_data(filename, asset_type)
        print(f"  Found {len(assets)} {asset_type}(s)")
        all_assets.extend(assets)
    
    print(f"\nTotal assets loaded: {len(all_assets)}")
    
    # Divide into groups
    print("\nDividing assets into 3 groups...")
    groups, group_values, group_counts = divide_into_groups(all_assets)
    
    # Display results
    print("\nGroup Distribution:")
    for group_name in ['A', 'B', 'C']:
        print(f"\nGroup {group_name}:")
        print(f"  Total Assets: {len(groups[group_name])}")
        print(f"  Total Value: BDT {group_values[group_name]:,.2f}")
        print(f"  Asset Types: {dict(group_counts[group_name])}")
    
    # Generate JSON files
    print("\nGenerating JSON files...")
    
    # all_assets_data.json
    output_data = {
        'groups': {}
    }
    for group_name in ['A', 'B', 'C']:
        output_data['groups'][group_name] = {
            'assets': groups[group_name],
            'total_value': group_values[group_name],
            'count': len(groups[group_name])
        }
    
    with open('all_assets_data.json', 'w', encoding='utf-8') as f:
        json.dump(output_data, f, indent=2, ensure_ascii=False)
    
    # all_assets_final.json
    with open('all_assets_final.json', 'w', encoding='utf-8') as f:
        json.dump(all_assets, f, indent=2, ensure_ascii=False)
    
    # Individual type JSON files
    assets_by_type = defaultdict(list)
    for asset in all_assets:
        assets_by_type[asset['AssetType']].append(asset)
    
    type_to_file = {
        'Desktop PC': 'PC_data.json',
        'Monitor': 'Monitor_data.json',
        'Printer/Scanner': 'Printer Scaneer_data.json',
        'Server': 'Server_data.json'
    }
    
    for asset_type, filename in type_to_file.items():
        if asset_type in assets_by_type:
            with open(filename, 'w', encoding='utf-8') as f:
                json.dump(assets_by_type[asset_type], f, indent=2, ensure_ascii=False)
    
    # Generate HTML
    print("\nGenerating HTML report...")
    html_content = generate_html(groups, group_values, group_counts, all_assets)
    
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("\nProcessing complete!")
    print("   - Generated index.html")
    print("   - Generated all_assets_data.json")
    print("   - Generated all_assets_final.json")
    print("   - Generated individual type JSON files")

if __name__ == '__main__':
    main()
