import pandas as pd
import json
from datetime import datetime
import copy

# Read all asset files
print("Reading all asset files...")

# Read Monitors
monitors = []
df_mon = pd.read_excel('Monitor.xlsx')
for idx, row in df_mon.iterrows():
    if idx < 2:
        continue
    if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
        try:
            serial = int(row.iloc[0])
            monitor = {
                'AssetType': 'Monitor',
                'Serial': serial,
                'Name': str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A',
                'Model': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A',
                'User': str(row.iloc[3]) if pd.notna(row.iloc[3]) else 'N/A',
                'Location': str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'N/A',
                'ValueRatio': str(row.iloc[5]) if pd.notna(row.iloc[5]) else '2️⃣',
                'Status': str(row.iloc[6]) if pd.notna(row.iloc[6]) else 'Working',
                'Remarks': str(row.iloc[7]) if pd.notna(row.iloc[7]) else 'Good'
            }
            monitors.append(monitor)
        except:
            continue

# Read Desktop PCs (keep RAM values as they are in Excel)
pcs = []
df_pc = pd.read_excel('PC.xlsx')
for idx, row in df_pc.iterrows():
    if idx < 2:
        continue
    if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
        try:
            serial = int(row.iloc[0])
            ram_val = row.iloc[3]
            # Keep RAM as-is (17, 18, 19 are in the file)
            ram = int(ram_val) if pd.notna(ram_val) and not pd.isna(ram_val) else 0
            
            pc = {
                'AssetType': 'Desktop PC',
                'Serial': serial,
                'Name': str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A',
                'Model': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A',
                'RAM': ram,
                'Processor': str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'N/A',
                'Storage': str(row.iloc[5]) if pd.notna(row.iloc[5]) else 'N/A',
                'Gen': int(row.iloc[6]) if pd.notna(row.iloc[6]) else 0,
                'GPU': str(row.iloc[7]) if pd.notna(row.iloc[7]) else 'Integrated',
                'User': str(row.iloc[8]) if pd.notna(row.iloc[8]) else 'N/A',
                'Location': str(row.iloc[9]) if pd.notna(row.iloc[9]) else 'N/A',
                'Level': str(row.iloc[10]) if pd.notna(row.iloc[10]) else '3️⃣',
                'Status': str(row.iloc[11]) if pd.notna(row.iloc[11]) else 'Working',
                'Remarks': str(row.iloc[12]) if pd.notna(row.iloc[12]) else 'Good'
            }
            pcs.append(pc)
        except Exception as e:
            continue

# Read Printers/Scanners
printers = []
df_printer = pd.read_excel('Printer Scaneer.xlsx')
is_scanner_section = False
for idx, row in df_printer.iterrows():
    if idx < 2:
        continue
    if 'SCANNER' in str(row.iloc[0]).upper():
        is_scanner_section = True
        continue
    if pd.notna(row.iloc[1]) and isinstance(row.iloc[1], (int, float)):
        try:
            serial = int(row.iloc[1])
            category = str(row.iloc[0]) if pd.notna(row.iloc[0]) else 'A4'
            printer = {
                'AssetType': 'Scanner' if is_scanner_section else 'Printer',
                'Category': category,
                'Serial': serial,
                'DeviceName': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A',
                'Model': str(row.iloc[3]) if pd.notna(row.iloc[3]) else 'N/A',
                'DeviceSerial': str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'N/A',
                'Function': str(row.iloc[5]) if pd.notna(row.iloc[5]) else 'N/A',
                'Location': str(row.iloc[6]) if pd.notna(row.iloc[6]) else 'N/A',
                'ValueRatio': str(row.iloc[7]) if pd.notna(row.iloc[7]) else '3️⃣',
                'Status': str(row.iloc[8]) if pd.notna(row.iloc[8]) else 'Working',
                'Remarks': str(row.iloc[9]) if pd.notna(row.iloc[9]) else 'Good'
            }
            printers.append(printer)
        except:
            continue

# Read Server Room Devices
servers = []
df_server = pd.read_excel('Server.xlsx')
for idx, row in df_server.iterrows():
    if idx < 2:
        continue
    if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
        try:
            sl = int(row.iloc[0])
            server = {
                'AssetType': 'Server Device',
                'SL': sl,
                'DeviceName': str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A',
                'Model': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A',
                'Serial': str(row.iloc[3]) if pd.notna(row.iloc[3]) else 'N/A',
                'Status': str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'Working',
                'Remarks': str(row.iloc[5]) if pd.notna(row.iloc[5]) else 'Good'
            }
            servers.append(server)
        except:
            continue

# Read Laptops
laptops = []
purchase_type = 'New'
df_laptop = pd.read_excel('Laptop.xlsx')
for idx, row in df_laptop.iterrows():
    if idx == 0:
        continue
    if idx == 1:
        continue
    
    first_col = str(row.iloc[0])
    if pd.notna(first_col):
        first_col = first_col.strip()
        if 'New Purchase' in first_col:
            purchase_type = 'New'
            continue
        elif 'Recondition' in first_col:
            purchase_type = 'Reconditioned'
            continue
    
    if pd.notna(row.iloc[1]):
        try:
            serial = int(row.iloc[1])
            laptop = {
                'AssetType': 'Laptop',
                'Serial': serial,
                'Name': str(row.iloc[2]) if pd.notna(row.iloc[2]) else 'N/A',
                'Model': str(row.iloc[3]) if pd.notna(row.iloc[3]) else 'N/A',
                'RAM': int(row.iloc[4]) if pd.notna(row.iloc[4]) else 0,
                'Processor': str(row.iloc[5]) if pd.notna(row.iloc[5]) else 'N/A',
                'Storage': str(row.iloc[6]) if pd.notna(row.iloc[6]) else 'N/A',
                'Gen': int(row.iloc[7]) if pd.notna(row.iloc[7]) else 0,
                'GPU': str(row.iloc[8]) if pd.notna(row.iloc[8]) else 'Integrated',
                'User': str(row.iloc[9]) if pd.notna(row.iloc[9]) else 'Unassigned',
                'Location': str(row.iloc[10]) if pd.notna(row.iloc[10]) else 'N/A',
                'Level': str(row.iloc[11]) if pd.notna(row.iloc[11]) else '3️⃣',
                'Status': str(row.iloc[12]) if pd.notna(row.iloc[12]) else 'Working',
                'Remarks': str(row.iloc[13]) if pd.notna(row.iloc[13]) else 'Good',
                'PurchaseType': purchase_type
            }
            laptops.append(laptop)
        except:
            continue

print(f"Parsed: {len(monitors)} monitors, {len(pcs)} PCs, {len(printers)} printers, {len(servers)} servers, {len(laptops)} laptops")

# Calculate values with CORRECT HP Server valuation
def calculate_monitor_value(monitor):
    model = monitor['Model'].upper()
    value_ratio = monitor['ValueRatio']
    remarks = monitor['Remarks'].lower()
    
    if 'M22' in model or '22' in model:
        base_price = 8000
    elif 'VP229' in model or '19' in model:
        base_price = 6000
    elif 'D1918' in model or 'E1916' in model:
        base_price = 5000
    elif 'E170' in model or '17' in model:
        base_price = 4000
    else:
        base_price = 6000
    
    if '1️⃣' in value_ratio:
        base_price *= 1.2
    elif '3️⃣' in value_ratio:
        base_price *= 0.8
    
    if 'excellent' in remarks:
        multiplier = 1.0
    elif 'good' in remarks:
        multiplier = 0.75
    else:
        multiplier = 0.60
    
    base_price *= 0.70  # Depreciation
    return round(base_price * multiplier, 2)

def calculate_pc_value(pc):
    """Calculate PC value - SPECIAL HANDLING FOR HP SERVER"""
    processor = pc['Processor'].lower()
    gen = pc['Gen']
    ram = pc['RAM']
    remarks = pc['Remarks'].lower()
    level = pc['Level']
    name = pc['Name'].lower()
    model = pc['Model'].lower()
    
    # HP SERVER - MUCH HIGHER VALUATION
    if 'server' in name or 'proliant' in model:
        # Enterprise server - even if old, very valuable
        base_price = 120000  # Base value for enterprise server
        if ram >= 32:
            base_price += 20000
        if 'xeon' in processor:
            base_price += 30000
        # Apply depreciation for Gen 4 (old but still valuable)
        if gen == 4:
            base_price *= 0.60  # 40% depreciation for old server
        else:
            base_price *= 0.70
    else:
        # Regular desktop
        if 'i9' in processor:
            if gen >= 13:
                base_price = 80000
            elif gen >= 11:
                base_price = 60000
            else:
                base_price = 40000
        elif 'i5' in processor:
            if gen >= 11:
                base_price = 40000
            elif gen >= 8:
                base_price = 30000
            else:
                base_price = 20000
        elif 'i3' in processor:
            if gen >= 10:
                base_price = 25000
            else:
                base_price = 15000
        else:
            base_price = 20000
        
        # RAM adjustment (including 17, 18, 19GB)
        if ram >= 32:
            base_price += 10000
        elif ram >= 16:
            base_price += 5000
        elif ram >= 8:
            base_price += 2000
    
    # GPU adjustment
    if pc['GPU'] != 'Integrated' and 'nan' not in pc['GPU'].lower():
        if 'gb' in pc['GPU'].lower() or 'GB' in pc['GPU']:
            base_price += 8000
    
    # Level adjustment
    if '1️⃣' in level:
        base_price *= 1.15
    elif '3️⃣' in level:
        base_price *= 0.85
    
    # Condition adjustment
    if 'excellent' in remarks:
        multiplier = 1.0
    elif 'good' in remarks:
        multiplier = 0.80
    elif 'moderate' in remarks:
        multiplier = 0.60
    else:
        multiplier = 0.50
    
    # Depreciation for old PCs (not servers)
    if 'server' not in name and 'proliant' not in model:
        if gen > 0:
            years_old = max(0, 13 - gen)
            for year in range(years_old):
                base_price *= 0.75
        else:
            base_price *= 0.50
    
    return round(base_price * multiplier, 2)

def calculate_printer_value(printer):
    category = printer['Category'].upper()
    device_name = printer['DeviceName'].upper()
    value_ratio = printer['ValueRatio']
    remarks = printer['Remarks'].lower()
    
    if category == 'A3':
        base_price = 20000
    elif 'SCANNER' in device_name or printer['AssetType'] == 'Scanner':
        base_price = 8000
    else:
        base_price = 12000
    
    if '1️⃣' in value_ratio:
        base_price *= 1.3
    elif '2️⃣' in value_ratio:
        base_price *= 1.0
    elif '3️⃣' in value_ratio:
        base_price *= 0.9
    elif '4️⃣' in value_ratio:
        base_price *= 0.8
    elif '5️⃣' in value_ratio:
        base_price *= 0.7
    elif '6️⃣' in value_ratio:
        base_price *= 0.6
    
    if 'excellent' in remarks:
        multiplier = 1.0
    else:
        multiplier = 0.75
    
    base_price *= 0.60  # Depreciation
    return round(base_price * multiplier, 2)

def calculate_server_value(server):
    device_name = server['DeviceName'].upper()
    model = server['Model'].upper()
    remarks = server['Remarks'].lower()
    
    if 'ROUTER' in device_name:
        base_price = 15000
    elif 'SWITCH' in device_name:
        if '16 PORT' in device_name or '16G' in model:
            base_price = 10000
        elif '24 PORT' in device_name or '24' in model:
            base_price = 15000
        elif '8 PORT' in device_name or '8' in model:
            base_price = 6000
        else:
            base_price = 8000
    elif 'DVR' in device_name:
        if '16 PORT' in device_name or '16' in model:
            base_price = 12000
        elif '8 PORT' in device_name or '8' in model:
            base_price = 8000
        else:
            base_price = 10000
    elif 'KVM' in device_name:
        base_price = 5000
    elif 'UPS' in device_name:
        if 'BATTERY' in device_name:
            base_price = 8000
        else:
            base_price = 20000
    elif 'CONVERTER' in device_name or 'CONVETER' in device_name:
        base_price = 8000
    elif 'PANEL' in device_name:
        base_price = 3000
    elif 'CABLE' in device_name or 'MANAGER' in device_name:
        base_price = 2000
    elif 'CONTROLLER' in device_name or 'PBAX' in device_name:
        base_price = 10000
    else:
        base_price = 5000
    
    if 'excellent' in remarks:
        multiplier = 1.0
    else:
        multiplier = 0.75
    
    base_price *= 0.65  # Depreciation
    return round(base_price * multiplier, 2)

def calculate_laptop_value_bdt(laptop):
    processor = laptop['Processor'].lower()
    gen = laptop['Gen']
    ram = laptop['RAM']
    purchase_type = laptop['PurchaseType']
    remarks = laptop['Remarks'].lower()
    level = laptop['Level']
    
    if 'i7' in processor:
        if gen >= 11:
            base_price = 75000
        elif gen >= 8:
            base_price = 60000
        elif gen >= 6:
            base_price = 45000
        else:
            base_price = 30000
    elif 'i5' in processor:
        if gen >= 11:
            base_price = 60000
        elif gen >= 8:
            base_price = 45000
        elif gen >= 6:
            base_price = 35000
        else:
            base_price = 25000
    elif 'i3' in processor:
        if gen >= 11:
            base_price = 40000
        elif gen >= 8:
            base_price = 30000
        elif gen >= 5:
            base_price = 22000
        else:
            base_price = 15000
    else:
        base_price = 25000
    
    if ram >= 8:
        base_price += 5000
    
    if laptop['GPU'] != 'Integrated' and 'nan' not in laptop['GPU'].lower():
        if 'gb' in laptop['GPU'].lower() or 'GB' in laptop['GPU']:
            base_price += 8000
        else:
            base_price += 3000
    
    if purchase_type == 'Reconditioned':
        base_price *= 0.40
    
    current_gen = 13
    years_old = max(0, current_gen - gen)
    
    if purchase_type == 'New':
        if years_old >= 5:
            annual_depreciation = 0.35
        elif years_old >= 3:
            annual_depreciation = 0.30
        else:
            annual_depreciation = 0.25
    else:
        if years_old >= 5:
            annual_depreciation = 0.30
        else:
            annual_depreciation = 0.25
    
    for year in range(years_old):
        base_price *= (1 - annual_depreciation)
    
    if gen <= 5:
        base_price *= 0.70
    elif gen <= 7:
        base_price *= 0.85
    
    condition_multiplier = {
        'excellent': 1.0,
        'good': 0.80,
        'moderate': 0.60,
        'poor': 0.40
    }
    
    for condition, multiplier in condition_multiplier.items():
        if condition in remarks:
            base_price *= multiplier
            break
    
    if '1️⃣' in level:
        base_price *= 1.15
    elif '3️⃣' in level:
        base_price *= 0.85
    
    if base_price < 5000:
        base_price = 5000
    
    return round(base_price, 2)

# Calculate all values
all_assets = []

for monitor in monitors:
    monitor['CurrentValue'] = calculate_monitor_value(monitor)
    all_assets.append(monitor)

for pc in pcs:
    pc['CurrentValue'] = calculate_pc_value(pc)
    all_assets.append(pc)

for printer in printers:
    printer['CurrentValue'] = calculate_printer_value(printer)
    all_assets.append(printer)

for server in servers:
    server['CurrentValue'] = calculate_server_value(server)
    all_assets.append(server)

for laptop in laptops:
    laptop['CurrentValue'] = calculate_laptop_value_bdt(laptop)
    all_assets.append(laptop)

print(f"\nTotal assets: {len(all_assets)}")
print(f"Total value: {sum(a['CurrentValue'] for a in all_assets):,.0f} BDT")

# Find HP Server
hp_server = [a for a in all_assets if 'server' in a.get('Name', '').lower() or 'proliant' in a.get('Model', '').lower()]
if hp_server:
    print(f"\nHP Server value: {hp_server[0]['CurrentValue']:,.0f} BDT")

# EQUAL DISTRIBUTION ALGORITHM
print("\n" + "="*60)
print("IMPLEMENTING EQUAL DISTRIBUTION")
print("="*60)

def divide_assets_equally(all_assets):
    """Divide assets equally with ±1% variance target"""
    
    total_value = sum(a['CurrentValue'] for a in all_assets)
    target_value = total_value / 3
    
    groups = {
        'A': {'assets': [], 'total_value': 0, 'asset_counts': {}},
        'B': {'assets': [], 'total_value': 0, 'asset_counts': {}},
        'C': {'assets': [], 'total_value': 0, 'asset_counts': {}}
    }
    
    # Sort by value descending
    sorted_assets = sorted(all_assets, key=lambda x: x['CurrentValue'], reverse=True)
    
    # Identify high-value items (>10% of target)
    high_value_threshold = target_value * 0.10
    high_value_items = [a for a in sorted_assets if a['CurrentValue'] > high_value_threshold]
    regular_items = [a for a in sorted_assets if a['CurrentValue'] <= high_value_threshold]
    
    print(f"High-value items (>10% of target): {len(high_value_items)}")
    print(f"Regular items: {len(regular_items)}")
    
    # Step 1: Distribute high-value items round-robin
    for i, asset in enumerate(high_value_items):
        group_key = ['A', 'B', 'C'][i % 3]
        groups[group_key]['assets'].append(asset)
        groups[group_key]['total_value'] += asset['CurrentValue']
        asset_type = asset['AssetType']
        if asset_type not in groups[group_key]['asset_counts']:
            groups[group_key]['asset_counts'][asset_type] = 0
        groups[group_key]['asset_counts'][asset_type] += 1
    
    # Step 2: Distribute regular items using greedy algorithm
    for asset in regular_items:
        # Find group with minimum total value
        min_group = min(groups.keys(), key=lambda g: groups[g]['total_value'])
        groups[min_group]['assets'].append(asset)
        groups[min_group]['total_value'] += asset['CurrentValue']
        asset_type = asset['AssetType']
        if asset_type not in groups[min_group]['asset_counts']:
            groups[min_group]['asset_counts'][asset_type] = 0
        groups[min_group]['asset_counts'][asset_type] += 1
    
    # Step 3: Fine-tuning - swap items if variance > 1%
    max_iterations = 100
    for iteration in range(max_iterations):
        values = [groups[g]['total_value'] for g in groups]
        max_val = max(values)
        min_val = min(values)
        variance_pct = ((max_val - min_val) / target_value) * 100
        
        if variance_pct <= 1.0:
            break
        
        # Try to swap items
        max_group = max(groups.keys(), key=lambda g: groups[g]['total_value'])
        min_group = min(groups.keys(), key=lambda g: groups[g]['total_value'])
        
        # Find items to swap
        max_group_assets = sorted(groups[max_group]['assets'], key=lambda x: x['CurrentValue'], reverse=True)
        min_group_assets = sorted(groups[min_group]['assets'], key=lambda x: x['CurrentValue'], reverse=True)
        
        swapped = False
        for max_asset in max_group_assets[:5]:  # Try top 5 from max group
            for min_asset in min_group_assets[:5]:  # Try top 5 from min group
                # Calculate new totals
                new_max_total = groups[max_group]['total_value'] - max_asset['CurrentValue'] + min_asset['CurrentValue']
                new_min_total = groups[min_group]['total_value'] - min_asset['CurrentValue'] + max_asset['CurrentValue']
                
                # Check if swap improves balance
                new_max = max(new_max_total, new_min_total, 
                             max(groups[g]['total_value'] for g in groups if g not in [max_group, min_group]))
                new_min = min(new_max_total, new_min_total,
                             min(groups[g]['total_value'] for g in groups if g not in [max_group, min_group]))
                new_variance = ((new_max - new_min) / target_value) * 100
                
                if new_variance < variance_pct:
                    # Perform swap
                    groups[max_group]['assets'].remove(max_asset)
                    groups[min_group]['assets'].remove(min_asset)
                    groups[max_group]['assets'].append(min_asset)
                    groups[min_group]['assets'].append(max_asset)
                    groups[max_group]['total_value'] = new_max_total
                    groups[min_group]['total_value'] = new_min_total
                    swapped = True
                    break
            if swapped:
                break
        
        if not swapped:
            break
    
    return groups

# Divide assets
groups = divide_assets_equally(all_assets)

# Generate detailed allocation remarks
def generate_allocation_remark(asset, group_name, groups, all_assets):
    """Generate detailed remark explaining why asset was assigned to this group"""
    total_value = sum(a['CurrentValue'] for a in all_assets)
    target_value = total_value / 3
    avg_value_per_asset = total_value / len(all_assets)
    
    group = groups[group_name]
    group_value = group['total_value']
    group_count = len(group['assets'])
    avg_count = len(all_assets) / 3
    
    remarks = []
    
    # Value contribution
    if asset['CurrentValue'] > avg_value_per_asset * 1.5:
        remarks.append("High-value item")
    elif asset['CurrentValue'] < avg_value_per_asset * 0.5:
        remarks.append("Lower-value item")
    else:
        remarks.append("Standard-value item")
    
    # Group balance reason
    if group_value < target_value * 0.99:
        remarks.append("assigned to balance total value (group was below target)")
    elif group_value > target_value * 1.01:
        remarks.append("assigned despite group being above target to maintain type balance")
    else:
        remarks.append("assigned to maintain balanced distribution")
    
    # Count compensation
    if group_count < avg_count - 1:
        remarks.append("compensates for receiving fewer items")
    elif group_count > avg_count + 1:
        remarks.append("group received more items but lower total value")
    
    # Quality factor
    remarks_lower = asset['Remarks'].lower()
    if 'excellent' in remarks_lower:
        remarks.append("excellent condition adds value")
    elif 'good' in remarks_lower:
        remarks.append("good condition")
    elif 'moderate' in remarks_lower:
        remarks.append("moderate condition")
    
    # Type distribution
    asset_type = asset['AssetType']
    type_count_in_group = group['asset_counts'].get(asset_type, 0)
    total_type_count = sum(1 for a in all_assets if a['AssetType'] == asset_type)
    avg_type_per_group = total_type_count / 3
    
    if type_count_in_group <= avg_type_per_group:
        remarks.append("balances asset type mix")
    
    # Specific reason
    if asset['CurrentValue'] > target_value * 0.10:
        remarks.append("high-value item distributed to ensure equal group totals")
    
    return ". ".join(remarks).capitalize() + "."

# Add remarks to all assets
for group_name in ['A', 'B', 'C']:
    for asset in groups[group_name]['assets']:
        asset['AllocationRemark'] = generate_allocation_remark(asset, group_name, groups, all_assets)

# Print summary
print("\n" + "="*60)
print("DIVISION SUMMARY")
print("="*60)

total_value = sum(a['CurrentValue'] for a in all_assets)
target_value = total_value / 3

for group_name in ['A', 'B', 'C']:
    group = groups[group_name]
    variance = ((group['total_value'] - target_value) / target_value * 100)
    variance_sign = "+" if variance > 0 else ""
    print(f"\nGroup {group_name}:")
    print(f"  Total Assets: {len(group['assets'])}")
    print(f"  Total Value: {group['total_value']:,.0f} BDT")
    print(f"  Variance: {variance_sign}{variance:.2f}%")
    print(f"  Asset Types: {dict(group['asset_counts'])}")

# Save data
with open('all_assets_final.json', 'w', encoding='utf-8') as f:
    json.dump({
        'groups': groups,
        'all_assets': all_assets,
        'total_value': total_value,
        'target_value': target_value
    }, f, indent=2, ensure_ascii=False, default=str)

print("\nData saved to all_assets_final.json")
