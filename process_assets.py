import pandas as pd
import json
from datetime import datetime
from collections import defaultdict

# Market Prices (Current New Market Price in BDT as of Jan 2026)
# Based on web research for Bangladesh market
MARKET_PRICES = {
    # === LAPTOPS ===
    'Laptop': {
        # New Laptops (i7)
        'DELL Inspiron 13 5310': 115000,  # i7 11th gen
        'Dell P106F': 110000,  # i7 11th gen
        'HP Probook 440 G4': 95000,  # i7 7th gen (older)
        
        # New Laptops (i5)
        'HP 15s-fq5317TU': 72500,  # i5 12th gen (updated price)
        'HP Probook 440 G8': 85000,  # i5 11th gen
        'Dell Inspiron 3511': 75000,  # i5 11th gen
        'Dell Inspiron 5410': 78000,  # i5 11th gen
        'Dell Inspiron 3493': 65000,  # i5 10th gen
        'HP 15s-fq2643TU': 70000,  # i5 11th gen
        'HP RTL8822CE': 72500,  # i5 12th gen (same model family)
        
        # New Laptops (i3)
        'LENOVO V14-G2-ITL': 48000,  # i3 11th gen
        
        # Reconditioned Laptops (i7)
        'HP 440-G4 recond': 60000,  # i7 7th gen used
        
        # Reconditioned Laptops (i5)
        'HP Elitbook 840 G3': 48000,  # i5 6th gen used (premium model)
        'HP 830 G5': 55000,  # i5 8th gen used
        'HP Elitbook': 45000,  # i5 generic elitebook used
        'HP HSN-112C': 50000,  # i5 8th gen used
        'Lenovo X230': 28000,  # i5 3rd gen very old but durable
        'HP Laptop generic': 45000,  # generic HP i5 used
        
        # Reconditioned Laptops (i3)
        'DELL Latitute 3350': 35000,  # i3 5th gen used
        'DELL inspiron 3542': 32000,  # i3 4th gen used
        'HP 440-G2': 30000,  # i3 5th gen used
        'ASUS generic': 28000,  # i3 generic ASUS used
    },
    
    # === DESKTOP PCs ===
    'PC': {
        'HP ProLiant DL380 Gen10 Plus': 650000,  # Server (midrange spec)
        'Desktop i9 Gen 13': 280000,  # High-end gaming/workstation
        'Desktop i9 Gen 11': 220000,  # Slightly older i9
        'Desktop i5 Gen 10': 85000,  # Mid-range desktop
        'Desktop i3 Gen 10': 55000,  # Entry desktop
        'Desktop i3 Gen 4': 35000,  # Very old desktop
        'HP ProDesk': 75000,  # i5 business desktop
        'Dell Optiplex 3020': 45000,  # i3 older business desktop
    },
    
    # === MONITORS ===
    'Monitor': {
        'HP M22f': 15500,  # 22" FHD IPS
        'ASUS VP229HE': 15500,  # 21.5" FHD
        'Dell 52240Lc': 12000,  # Older Dell model
        'Dell D1918Ho': 10000,  # 19" basic
        'Dell E170SC': 6000,  # Very old 17" model
        'Dell E1916HVF': 10000,  # 19" basic
        'Dell G2QR5D2': 11000,  # Standard Dell
        'HP LV1911': 9000,  # 18.5" basic
        'SAMSUNG S19C300B': 9500,  # 19" Samsung
    },
    
    # === PRINTERS & SCANNERS ===
    'Printer Scaneer': {
        'EPSON A3 C722B1': 80000,  # A3 color (L8180 equivalent)
        'EPSON L3250': 21000,  # A4 color inkjet multifunction
        'HP M26nw': 32000,  # Laser MFP
        'HP M404dw': 42000,  # Laser printer with duplex/wifi
        'HP M402dn': 22000,  # Laser printer with duplex/network
        'PANTUM P3010DW': 14000,  # Budget laser duplex wifi
        'Canon F166500': 16000,  # Canon multifunction
        'Canon F166400': 15000,  # Canon multifunction
        'HP Laser 107w': 11000,  # Basic laser wifi
        'SAMSUNG M2820ND': 18000,  # Laser network duplex
        'EPSON J371A scanner': 12000,  # Dedicated scanner
        'Canon EUROPA scanner': 10000,  # Scanner
    },
    
    # === SERVER ROOM / NETWORKING ===
    'Server Room': {
        'MikroTik RB2011UiAS-RM': 13000,  # Router
        'MikroTik CCR2004-16G-2S+': 85000,  # High-end switch
        'TP-Link TL-SG1024D': 7500,  # 24-port gigabit switch
        'Cisco CBS350-24T-4G': 33000,  # Managed 24-port switch
        'TP-Link T1500G-10PS': 18000,  # 8-port PoE switch
        'KVM Switch 8-port': 8000,  # KVM switch
        'Hikvision DS-720BHGH1-F2': 12000,  # 8-ch DVR
        'Hikvision DS-7216HQH1-K2': 23000,  # 16-ch DVR
        'Patch Panel 24-port': 3500,  # Patch panel
        'Cable Manager': 2000,  # Cable management
        'PBAX KX-Tes824BX': 25000,  # PABX controller
        'ACS Soyal AR-727CM-V3': 15000,  # Access control converter
        'UPS APC/SRCE6KUX1': 115000,  # UPS APC model specific
        'UPS Battery 16 PCS': 160000,  # UPS Battery 16 PCS
        'UPS generic': 160000,  # UPS system generic
        'Switch generic': 10000,  # Generic switch
    }
}

def get_market_price(asset):
    """Get market price for an asset based on name, model, and type."""
    asset_type = asset['AssetType']
    name = asset.get('Name', '').strip()
    model = asset.get('Model', '').strip()
    processor = asset.get('Processor', '').strip().lower()
    
    if asset_type not in MARKET_PRICES:
        return 50000  # Default fallback
    
    prices = MARKET_PRICES[asset_type]
    
    # Try exact model match first
    if model in prices:
        return prices[model]
    
    # Try name match
    if name in prices:
        return prices[name]
    
    # Try combined name + model
    combined = f"{name} {model}".strip()
    if combined in prices:
        return prices[combined]
    
    # Asset-specific matching logic
    if asset_type == 'Laptop':
        # Check if reconditioned
        is_recond = asset.get('PurchaseType') == 'reconditioned'
        
        # Match by brand and processor
        if 'DELL' in name.upper() or 'Dell' in name:
            if 'i7' in processor:
                return 110000 if not is_recond else 60000
            elif 'i5' in processor:
                return 75000 if not is_recond else 50000
            elif 'i3' in processor:
                return 50000 if not is_recond else 32000
        elif 'HP' in name.upper():
            if 'elite' in name.lower() or 'elite' in model.lower():
                return 90000 if not is_recond else 48000
            elif 'probook' in name.lower() or 'probook' in model.lower():
                if 'i7' in processor:
                    return 95000 if not is_recond else 60000
                elif 'i5' in processor:
                    return 85000 if not is_recond else 50000
            else:  # Regular HP
                if 'i7' in processor:
                    return 100000 if not is_recond else 60000
                elif 'i5' in processor:
                    return 72500 if not is_recond else 45000
                elif 'i3' in processor:
                    return 48000 if not is_recond else 30000
        elif 'LENOVO' in name.upper() or 'Lenovo' in name:
            if 'thinkpad' in name.lower() or 'x230' in model.lower():
                return 55000 if not is_recond else 28000
            elif 'i3' in processor:
                return 48000 if not is_recond else 35000
            elif 'i5' in processor:
                return 70000 if not is_recond else 45000
        elif 'ASUS' in name.upper():
            if 'i3' in processor:
                return 45000 if not is_recond else 28000
            elif 'i5' in processor:
                return 65000 if not is_recond else 40000
        
        # Generic fallback by processor
        if 'i7' in processor:
            return 105000 if not is_recond else 55000
        elif 'i5' in processor:
            return 70000 if not is_recond else 45000
        elif 'i3' in processor:
            return 48000 if not is_recond else 30000
        
        return 60000  # Generic laptop fallback
    
    elif asset_type == 'PC':
        if 'server' in name.lower() or 'proliant' in name.lower():
            return 650000
        elif 'xeon' in processor:
            return 650000
        elif 'i9' in processor:
            gen = asset.get('Gen', 0)
            if gen >= 13:
                return 280000
            elif gen >= 11:
                return 220000
            else:
                return 180000
        elif 'i5' in processor:
            return 85000
        elif 'i3' in processor:
            gen = asset.get('Gen', 0)
            if gen >= 10:
                return 55000
            else:
                return 35000
        elif 'dual' in processor or ('core' in processor and 'dual' in processor):
            return 30000  # Dual Core processors
        return 70000  # Generic desktop
    
    elif asset_type == 'Monitor':
        if 'm22f' in model.lower():
            return 15500
        elif 'asus' in name.lower():
            return 15500
        elif '24' in model or '22' in model or '21' in model:
            return 12000
        elif '19' in model or '18' in model:
            return 10000
        elif '17' in model:
            return 6000
        return 11000  # Generic monitor
    
    elif asset_type == 'Printer Scaneer':
        model_lower = model.lower()
        name_lower = name.lower()
        
        if 'a3' in name_lower or 'l8180' in model_lower or 'c722b1' in model_lower:
            return 80000
        elif 'l3250' in model_lower:
            return 21000
        elif 'm404' in model_lower:
            return 42000
        elif 'm402' in model_lower:
            return 22000
        elif 'm26' in model_lower:
            return 32000
        elif '107w' in model_lower:
            return 11000
        elif 'pantum' in name_lower or 'p3010' in model_lower:
            return 14000
        elif 'samsung' in name_lower:
            return 18000
        elif 'canon' in name_lower:
            if 'scanner' in name_lower or 'scaneer' in name_lower:
                return 10000
            return 16000
        elif 'epson' in name_lower:
            if 'scanner' in name_lower or 'scaneer' in name_lower:
                return 12000
            return 21000
        return 15000  # Generic printer
    
    elif asset_type == 'Server Room':
        model_lower = model.lower()
        name_lower = name.lower()
        
        # Check for specific UPS models first
        if 'apc' in model_lower or 'srce6kux1' in model_lower:
            return 115000
        elif 'battery' in name_lower and '16' in name_lower:
            return 160000
        elif 'mikrotik' in name_lower:
            if 'ccr2004' in model_lower or '16g' in model_lower:
                return 85000
            return 13000
        elif 'cisco' in name_lower:
            return 33000
        elif 'tp-link' in name_lower or 'tplink' in name_lower:
            if '24' in name_lower:
                return 7500
            elif '8' in name_lower:
                return 18000
            return 12000
        elif 'hikvision' in name_lower or 'dvr' in name_lower:
            if '16' in name_lower:
                return 23000
            elif '8' in name_lower:
                return 12000
            return 18000
        elif 'kvm' in name_lower:
            return 8000
        elif 'patch' in name_lower:
            return 3500
        elif 'cable' in name_lower:
            return 2000
        elif 'pbax' in name_lower or 'pabx' in name_lower:
            return 25000
        elif 'acs' in name_lower or 'soyal' in name_lower:
            return 15000
        elif 'ups' in name_lower:
            return 160000  # UPS with 16 PCS batteries
        return 10000  # Generic network device
    
    return 50000  # Final fallback

def calculate_asset_value(asset):
    """Calculate current value based on market price and remarks-based depreciation."""
    # Get market price
    market_price = get_market_price(asset)
    
    # Special case: PC serials 11, 13, 14, 15 get 90% reduction (10% of market price)
    if asset.get('AssetType') == 'PC':
        serial = asset.get('Serial')
        if serial in [11, 13, 14, 15]:
            final_value = market_price * 0.10  # 90% reduction
            return round(final_value, 2)
    
    # Check if reconditioned (laptops only)
    is_reconditioned_laptop = False
    if asset.get('AssetType') == 'Laptop':
        if asset.get('PurchaseType') == 'reconditioned':
            is_reconditioned_laptop = True
        # Check all string fields for "recondition" keyword
        for key, value in asset.items():
            if isinstance(value, str) and 'recondition' in value.lower():
                is_reconditioned_laptop = True
                break
    
    # Get remarks field
    remarks = asset.get('Remarks', '').lower()
    
    # Determine depreciation rate based on remarks and asset type
    if is_reconditioned_laptop:
        # Reconditioned laptops: Higher depreciation
        if 'excellent' in remarks:
            depreciation_rate = 0.40  # 60% reduction
            depreciation_label = '60%'
        elif 'good' in remarks:
            depreciation_rate = 0.30  # 70% reduction
            depreciation_label = '70%'
        elif 'moderate' in remarks or 'fair' in remarks:
            depreciation_rate = 0.20  # 80% reduction
            depreciation_label = '80%'
        else:
            depreciation_rate = 0.30  # Default to 70% reduction for reconditioned
            depreciation_label = '70%'
    else:
        # All other assets (new laptops, all PCs, monitors, printers, servers)
        if 'excellent' in remarks:
            depreciation_rate = 0.70  # 30% reduction
            depreciation_label = '30%'
        elif 'good' in remarks:
            depreciation_rate = 0.60  # 40% reduction
            depreciation_label = '40%'
        elif 'moderate' in remarks or 'fair' in remarks:
            depreciation_rate = 0.50  # 50% reduction
            depreciation_label = '50%'
        else:
            depreciation_rate = 0.70  # Default to 30% reduction
            depreciation_label = '30%'
    
    final_value = market_price * depreciation_rate
    
    return round(final_value, 2)

def read_excel_data(file_path, sheet_name):
    """Read Excel sheet and extract asset data."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
        assets = []
        
        # Find header row (contains "Serial" or "S/L")
        header_row = None
        for idx, row in df.iterrows():
            if 'Serial' in str(row.values) or 'S/L' in str(row.values):
                header_row = idx
                break
        
        if header_row is None:
            print(f"  Warning: Could not find header in {sheet_name}")
            return assets
        
        # Track current section for laptops (new vs reconditioned)
        current_section = 'new'  # default
        
        # Check rows before header for section markers (laptops only)
        if sheet_name == 'Laptop':
            for idx in range(0, header_row):
                row = df.iloc[idx]
                if len(row) > 0:
                    first_col = str(row.iloc[0]).lower() if not pd.isna(row.iloc[0]) else ''
                    if 'new purchase' in first_col:
                        current_section = 'new'
        
        # Determine column structure
        serial_col = None
        for col_idx, val in enumerate(df.iloc[header_row].values):
            if 'Serial' in str(val) or 'S/L' in str(val):
                serial_col = col_idx
                break
        
        if serial_col is None:
            return assets
        
        has_category = (serial_col == 1)  # If Serial is in column 1, there's a Category column
        
        # Process data rows
        for idx in range(header_row + 1, len(df)):
            row = df.iloc[idx]
            
            # Check for section headers in laptops
            if sheet_name == 'Laptop' and has_category and len(row) > 0:
                first_col = str(row.iloc[0]).lower() if not pd.isna(row.iloc[0]) else ''
                if 'recondition' in first_col or 'used' in first_col:
                    current_section = 'reconditioned'
                    continue
                elif 'new purchase' in first_col:
                    current_section = 'new'
                    continue
            
            # Get serial number
            serial = row.iloc[serial_col] if len(row) > serial_col else None
            
            if pd.isna(serial) or serial == '':
                continue
            
            # Skip non-numeric serials (headers)
            try:
                serial_num = int(serial)
            except (ValueError, TypeError):
                continue
            
            # Build asset dictionary
            asset = {
                'AssetType': sheet_name,
                'Serial': serial_num,
            }
            
            # Mark purchase type for laptops
            if sheet_name == 'Laptop':
                asset['PurchaseType'] = current_section
            
            # Column mapping
            col_offset = 1 if has_category else 0
            
            # Name
            name_col = col_offset + 1
            if len(row) > name_col and not pd.isna(row.iloc[name_col]):
                asset['Name'] = str(row.iloc[name_col]).strip()
            
            # Model
            model_col = col_offset + 2
            if len(row) > model_col and not pd.isna(row.iloc[model_col]):
                asset['Model'] = str(row.iloc[model_col]).strip()
            
            # Additional fields based on sheet type
            if sheet_name in ['Laptop', 'PC']:
                # RAM
                ram_col = col_offset + 3
                if len(row) > ram_col and not pd.isna(row.iloc[ram_col]):
                    try:
                        asset['RAM'] = int(row.iloc[ram_col])
                    except (ValueError, TypeError):
                        pass
                
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
                        pass
                
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
            
            elif sheet_name == 'Monitor':
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
            
            elif sheet_name == 'Printer Scaneer':
                # Function
                func_col = col_offset + 4
                if len(row) > func_col and not pd.isna(row.iloc[func_col]):
                    asset['Function'] = str(row.iloc[func_col]).strip()
                
                # Location
                loc_col = col_offset + 5
                if len(row) > loc_col and not pd.isna(row.iloc[loc_col]):
                    asset['Location'] = str(row.iloc[loc_col]).strip()
                
                # Level
                level_col = col_offset + 6
                if len(row) > level_col and not pd.isna(row.iloc[level_col]):
                    asset['Level'] = str(row.iloc[level_col]).strip()
                
                # Status
                status_col = col_offset + 7
                if len(row) > status_col and not pd.isna(row.iloc[status_col]):
                    asset['Status'] = str(row.iloc[status_col]).strip()
                
                # Remarks
                remarks_col = col_offset + 8
                if len(row) > remarks_col and not pd.isna(row.iloc[remarks_col]):
                    asset['Remarks'] = str(row.iloc[remarks_col]).strip()
            
            elif sheet_name == 'Server Room':
                # Status
                status_col = 4
                if len(row) > status_col and not pd.isna(row.iloc[status_col]):
                    asset['Status'] = str(row.iloc[status_col]).strip()
                
                # Remarks
                remarks_col = 5
                if len(row) > remarks_col and not pd.isna(row.iloc[remarks_col]):
                    asset['Remarks'] = str(row.iloc[remarks_col]).strip()
            
            # Only add if we have at least a name
            if 'Name' in asset and asset['Name']:
                assets.append(asset)
        
        return assets
    except Exception as e:
        print(f"Error reading {sheet_name}: {e}")
        import traceback
        traceback.print_exc()
        return []

def divide_into_groups(assets):
    """Divide assets into 3 equal groups by value using greedy algorithm."""
    # Calculate values for all assets
    for asset in assets:
        asset['MarketPrice'] = get_market_price(asset)
        asset['CurrentValue'] = calculate_asset_value(asset)
        
        # Determine depreciation label and remark category based on remarks and type
        is_reconditioned_laptop = False
        if asset.get('AssetType') == 'Laptop' and asset.get('PurchaseType') == 'reconditioned':
            is_reconditioned_laptop = True
        
        remarks = asset.get('Remarks', '').lower()
        
        # Determine remark category for display
        if 'excellent' in remarks:
            remark_category = 'Excellent'
        elif 'good' in remarks:
            remark_category = 'Good'
        elif 'moderate' in remarks or 'fair' in remarks:
            remark_category = 'Moderate'
        else:
            remark_category = 'Excellent'  # Default
        
        # Special case: PC serials 11, 13, 14, 15 get 90% reduction
        if asset.get('AssetType') == 'PC' and asset.get('Serial') in [11, 13, 14, 15]:
            asset['DepreciationRate'] = '90%'
            asset['RemarkCategory'] = 'Special'
        elif is_reconditioned_laptop:
            if 'excellent' in remarks:
                asset['DepreciationRate'] = '60%'
            elif 'good' in remarks:
                asset['DepreciationRate'] = '70%'
            elif 'moderate' in remarks or 'fair' in remarks:
                asset['DepreciationRate'] = '80%'
            else:
                asset['DepreciationRate'] = '70%'
                remark_category = 'Good'  # Default for reconditioned
        else:
            if 'excellent' in remarks:
                asset['DepreciationRate'] = '30%'
            elif 'good' in remarks:
                asset['DepreciationRate'] = '40%'
            elif 'moderate' in remarks or 'fair' in remarks:
                asset['DepreciationRate'] = '50%'
            else:
                asset['DepreciationRate'] = '30%'
        
        # Store remark category for display
        asset['RemarkCategory'] = remark_category
    
    # Sort all assets by value (descending) for standard greedy algorithm
    sorted_assets = sorted(assets, key=lambda x: x['CurrentValue'], reverse=True)
    
    # Initialize groups
    groups = {'A': [], 'B': [], 'C': []}
    group_values = {'A': 0, 'B': 0, 'C': 0}
    group_counts = {'A': defaultdict(int), 'B': defaultdict(int), 'C': defaultdict(int)}
    
    # Standard greedy algorithm - assign each asset to group with lowest total value
    for asset in sorted_assets:
        # Find group with minimum value
        min_group = min(group_values, key=group_values.get)
        
        # Add asset to that group
        groups[min_group].append(asset)
        group_values[min_group] += asset['CurrentValue']
        group_counts[min_group][asset['AssetType']] += 1
        
        # Add allocation remark
        asset['AllocationRemark'] = f"Assigned to balance total value (Group {min_group}: BDT {group_values[min_group]:,.0f})"
    
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
    <title>Genesis Technologies Ltd - Complete Asset Division (Market Price Based)</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 0;
            min-height: 100vh;
        }}
        
        .tab-navigation {{
            max-width: 1600px;
            margin: 20px auto 0 auto;
            padding: 0 20px;
            display: flex;
            gap: 5px;
            background: transparent;
        }}
        
        .tab-link {{
            color: rgba(255, 255, 255, 0.8);
            text-decoration: none;
            padding: 15px 30px;
            border-radius: 10px 10px 0 0;
            font-weight: 600;
            font-size: 1.05em;
            transition: all 0.3s ease;
            background: rgba(255, 255, 255, 0.1);
            border: 2px solid transparent;
            border-bottom: none;
            position: relative;
            top: 2px;
        }}
        
        .tab-link:hover {{
            background: rgba(255, 255, 255, 0.2);
            color: white;
        }}
        
        .tab-link.active {{
            background: white;
            color: #2c3e50;
            border: 2px solid white;
            border-bottom: none;
            box-shadow: 0 -2px 10px rgba(0,0,0,0.1);
        }}
        
        .container {{
            max-width: 1600px;
            margin: 0 auto 20px auto;
            background: white;
            border-radius: 0 15px 15px 15px;
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
        
        .new-badge {{
            background: #2ecc71;
            color: white;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 0.8em;
            margin-left: 8px;
        }}
        
        .recond-badge {{
            background: #e67e22;
            color: white;
            padding: 3px 8px;
            border-radius: 3px;
            font-size: 0.8em;
            margin-left: 8px;
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
    <nav class="tab-navigation">
        <a href="index.html" class="tab-link active">Asset Distribution</a>
        <a href="methodology.html" class="tab-link">Methodology & Reports</a>
    </nav>

    <div class="container">
        <div class="header">
            <h1>Genesis Technologies Ltd</h1>
            <p>Complete Asset Division Report - Market Price Based Valuation</p>
        </div>
        
        <div class="notice">
            <strong>Updated Pricing Methodology:</strong> All asset values are based on current market prices (January 2026) with remarks-based depreciation. <strong>NEW ASSETS:</strong> Excellent: 70% value (30% reduction), Good: 60% value (40% reduction), Moderate: 50% value (50% reduction). <strong>RECONDITIONED LAPTOPS:</strong> Excellent: 40% value (60% reduction), Good: 30% value (70% reduction), Moderate: 20% value (80% reduction).
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Total Assets</h3>
                <div class="value">{len(all_assets)}</div>
                <div class="label">All Asset Types</div>
            </div>
            <div class="summary-card">
                <h3>Laptops</h3>
                <div class="value">{type_counts.get('Laptop', 0)}</div>
                <div class="label">Laptop Computers</div>
            </div>
            <div class="summary-card">
                <h3>Desktop PCs</h3>
                <div class="value">{type_counts.get('PC', 0)}</div>
                <div class="label">Desktop & Servers</div>
            </div>
            <div class="summary-card">
                <h3>Monitors</h3>
                <div class="value">{type_counts.get('Monitor', 0)}</div>
                <div class="label">Display Monitors</div>
            </div>
            <div class="summary-card">
                <h3>Printers</h3>
                <div class="value">{type_counts.get('Printer Scaneer', 0)}</div>
                <div class="label">Printers & Scanners</div>
            </div>
            <div class="summary-card">
                <h3>Server Devices</h3>
                <div class="value">{type_counts.get('Server Room', 0)}</div>
                <div class="label">Network & Server Equipment</div>
            </div>
            <div class="summary-card">
                <h3>Total Value</h3>
                <div class="value">BDT {total_value:,.0f}</div>
                <div class="label">Combined Worth</div>
            </div>
            <div class="summary-card">
                <h3>Target per Group</h3>
                <div class="value">BDT {target_value:,.0f}</div>
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
                        <div class="stat-value">BDT {group_val:,.0f}</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Variance</div>
                        <div class="stat-value">{variance:+.2f}%</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Avg per Asset</div>
                        <div class="stat-value">BDT {group_val/len(group_assets):,.0f}</div>
                    </div>
"""
        
        # Add type counts
        for asset_type in ['Laptop', 'PC', 'Monitor', 'Printer Scaneer', 'Server Room']:
            count = type_counts_group.get(asset_type, 0)
            if count > 0:
                display_name = asset_type.replace('Printer Scaneer', 'Printer').replace('Server Room', 'Server')
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
        for asset_type in ['Laptop', 'PC', 'Monitor', 'Printer Scaneer', 'Server Room']:
            if asset_type not in assets_by_type:
                continue
            
            type_assets = sorted(assets_by_type[asset_type], key=lambda x: x['CurrentValue'], reverse=True)
            
            html += f"""
                    <div class="asset-type-header">
                        {asset_type}s ({len(type_assets)}) - Total Value: BDT {sum(a['CurrentValue'] for a in type_assets):,.0f}
                    </div>
                    
                    <table class="asset-table">
                        <thead>
                            <tr>
                                <th>Serial</th>
                                <th>Name & Type</th>
                                <th>Model</th>
                                <th>User</th>
                                <th>Specs</th>
                                <th>Market Price</th>
                                <th>Depreciation</th>
                                <th>Current Value</th>
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
                if asset.get('Location'):
                    specs.append(f"Loc: {asset['Location']}")
                
                specs_str = '<br>'.join(specs[:4]) if specs else 'N/A'  # Limit to 4 lines
                
                # Get user info
                user_info = asset.get('User', 'N/A')
                
                # Determine if new or reconditioned
                is_recond = asset.get('PurchaseType') == 'reconditioned' or asset['DepreciationRate'] in ['60%', '70%', '80%']
                status_badge = '<span class="recond-badge">RECOND</span>' if is_recond else '<span class="new-badge">NEW</span>'
                
                # Format depreciation with remark
                depreciation_display = f"{asset['DepreciationRate']}<br><small>({asset.get('RemarkCategory', 'N/A')})</small>"
                
                html += f"""
                            <tr>
                                <td><strong>{asset.get('Serial', 'N/A')}</strong></td>
                                <td>{asset.get('Name', 'N/A')} {status_badge}</td>
                                <td>{asset.get('Model', 'N/A')}</td>
                                <td><small>{user_info}</small></td>
                                <td><small>{specs_str}</small></td>
                                <td><small>BDT {asset['MarketPrice']:,.0f}</small></td>
                                <td><small>{depreciation_display}</small></td>
                                <td><span class="value-badge">BDT {asset['CurrentValue']:,.0f}</span></td>
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
    print("Starting asset processing with market-based pricing...")
    
    # Read all sheets from combined Excel file
    file_path = 'GTL IT Equipment Information 22 Jan 2026.xlsx'
    sheet_names = ['Laptop', 'PC', 'Monitor', 'Printer Scaneer', 'Server Room']
    
    all_assets = []
    
    # Read all sheets
    for sheet_name in sheet_names:
        print(f"Reading {sheet_name}...")
        assets = read_excel_data(file_path, sheet_name)
        print(f"  Found {len(assets)} {sheet_name}(s)")
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
        'groups': {},
        'pricing_methodology': {
            'description': 'Market-based pricing with depreciation',
            'new_asset_depreciation': '30%',
            'reconditioned_asset_depreciation': '60%',
            'market_price_date': 'January 2026'
        }
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
    
    # Generate HTML
    print("\nGenerating HTML report...")
    html_content = generate_html(groups, group_values, group_counts, all_assets)
    
    with open('index.html', 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    print("\nProcessing complete!")
    print("   - Generated index.html")
    print("   - Generated all_assets_data.json")
    print("   - Generated all_assets_final.json")
    print(f"\nTotal Value: BDT {sum(group_values.values()):,.2f}")
    print(f"Group A: BDT {group_values['A']:,.2f}")
    print(f"Group B: BDT {group_values['B']:,.2f}")
    print(f"Group C: BDT {group_values['C']:,.2f}")

if __name__ == '__main__':
    main()
