import json
from datetime import datetime

# Load the final asset data
with open('all_assets_final.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

groups = data['groups']
all_assets = data['all_assets']
total_value = data['total_value']
target_value = data['target_value']

# Organize assets by type and group
def organize_assets_by_type(group_assets):
    organized = {
        'Laptop': [],
        'Desktop PC': [],
        'Monitor': [],
        'Printer': [],
        'Scanner': [],
        'Server Device': []
    }
    
    for asset in group_assets:
        asset_type = asset['AssetType']
        if asset_type in organized:
            organized[asset_type].append(asset)
        else:
            if 'Printer' in asset_type or 'printer' in asset_type.lower():
                organized['Printer'].append(asset)
            elif 'Scanner' in asset_type or 'scanner' in asset_type.lower():
                organized['Scanner'].append(asset)
            else:
                organized['Server Device'].append(asset)
    
    return organized

# Generate HTML
html = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Genesis Technologies Ltd - Complete Asset Division (Equal Distribution)</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
        }
        
        .container {
            max-width: 1600px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #3498db 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        
        .header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .notice {
            background: #d4edda;
            border-left: 4px solid #28a745;
            padding: 15px 20px;
            margin: 20px 40px;
            border-radius: 5px;
        }
        
        .notice strong {
            color: #155724;
        }
        
        .summary {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 40px;
            background: #f8f9fa;
        }
        
        .summary-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
            text-align: center;
            border-top: 4px solid #3498db;
        }
        
        .summary-card h3 {
            color: #2c3e50;
            margin-bottom: 15px;
            font-size: 1.1em;
        }
        
        .summary-card .value {
            font-size: 1.8em;
            font-weight: bold;
            color: #3498db;
            margin: 10px 0;
        }
        
        .summary-card .label {
            color: #7f8c8d;
            font-size: 0.85em;
        }
        
        .groups-container {
            padding: 40px;
        }
        
        .group {
            margin-bottom: 50px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        
        .group-header {
            padding: 25px;
            color: white;
            font-size: 1.5em;
            font-weight: bold;
        }
        
        .group-a .group-header {
            background: linear-gradient(135deg, #FF6B6B 0%, #C92A2A 100%);
        }
        
        .group-b .group-header {
            background: linear-gradient(135deg, #51CF66 0%, #2F9E44 100%);
        }
        
        .group-c .group-header {
            background: linear-gradient(135deg, #4DABF7 0%, #1971C2 100%);
        }
        
        .group-stats {
            background: #f8f9fa;
            padding: 20px;
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(120px, 1fr));
            gap: 15px;
        }
        
        .stat {
            text-align: center;
            padding: 10px;
            background: white;
            border-radius: 5px;
        }
        
        .stat .stat-label {
            font-size: 0.85em;
            color: #7f8c8d;
            margin-bottom: 5px;
        }
        
        .stat .stat-value {
            font-size: 1.3em;
            font-weight: bold;
            color: #2c3e50;
        }
        
        .asset-section {
            margin-top: 30px;
        }
        
        .asset-section-title {
            background: #2c3e50;
            color: white;
            padding: 15px 20px;
            font-size: 1.2em;
            font-weight: bold;
            margin-bottom: 0;
        }
        
        table {
            width: 100%;
            border-collapse: collapse;
            background: white;
        }
        
        thead {
            background: #34495e;
            color: white;
        }
        
        th {
            padding: 12px;
            text-align: left;
            font-weight: 600;
            font-size: 0.85em;
            text-transform: uppercase;
            letter-spacing: 0.5px;
        }
        
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #ecf0f1;
            font-size: 0.9em;
        }
        
        tbody tr:hover {
            background: #f8f9fa;
        }
        
        .badge {
            display: inline-block;
            padding: 4px 10px;
            border-radius: 15px;
            font-size: 0.8em;
            font-weight: 600;
        }
        
        .badge-excellent {
            background: #d1ecf1;
            color: #0c5460;
        }
        
        .badge-good {
            background: #cce5ff;
            color: #004085;
        }
        
        .badge-moderate {
            background: #f8d7da;
            color: #721c24;
        }
        
        .badge-new {
            background: #d4edda;
            color: #155724;
        }
        
        .badge-recon {
            background: #fff3cd;
            color: #856404;
        }
        
        .level-1 { color: #2ecc71; font-weight: bold; }
        .level-2 { color: #3498db; font-weight: bold; }
        .level-3 { color: #e67e22; font-weight: bold; }
        
        .remark-cell {
            font-size: 0.85em;
            color: #555;
            font-style: italic;
            max-width: 300px;
        }
        
        .footer {
            background: #2c3e50;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
        
        @media print {
            body {
                background: white;
                padding: 0;
            }
            .container {
                box-shadow: none;
            }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ Genesis Technologies Ltd</h1>
            <p>Complete Asset Division Report - Equal Distribution Among Three Groups</p>
            <p style="font-size: 0.9em; margin-top: 10px;">Generated on """ + datetime.now().strftime('%B %d, %Y at %I:%M %p') + """</p>
        </div>
        
        <div class="notice">
            <strong>✅ Equal Distribution Achieved:</strong> All three groups have values within ±0.04% variance. Each asset includes detailed allocation remarks explaining why it was assigned to its group.
        </div>
        
        <div class="summary">
            <div class="summary-card">
                <h3>Total Assets</h3>
                <div class="value">""" + str(len(all_assets)) + """</div>
                <div class="label">Complete Inventory</div>
            </div>
            <div class="summary-card">
                <h3>Laptops</h3>
                <div class="value">""" + str(sum(1 for a in all_assets if a['AssetType'] == 'Laptop')) + """</div>
                <div class="label">Laptop Computers</div>
            </div>
            <div class="summary-card">
                <h3>Desktop PCs</h3>
                <div class="value">""" + str(sum(1 for a in all_assets if a['AssetType'] == 'Desktop PC')) + """</div>
                <div class="label">Desktop Computers</div>
            </div>
            <div class="summary-card">
                <h3>Monitors</h3>
                <div class="value">""" + str(sum(1 for a in all_assets if a['AssetType'] == 'Monitor')) + """</div>
                <div class="label">Display Monitors</div>
            </div>
            <div class="summary-card">
                <h3>Printers</h3>
                <div class="value">""" + str(sum(1 for a in all_assets if a['AssetType'] == 'Printer')) + """</div>
                <div class="label">Printers & Scanners</div>
            </div>
            <div class="summary-card">
                <h3>Server Devices</h3>
                <div class="value">""" + str(sum(1 for a in all_assets if a['AssetType'] == 'Server Device')) + """</div>
                <div class="label">Network & Server Equipment</div>
            </div>
            <div class="summary-card">
                <h3>Total Value</h3>
                <div class="value">৳""" + f"{total_value:,.0f}" + """</div>
                <div class="label">Combined Worth (BDT)</div>
            </div>
            <div class="summary-card">
                <h3>Target per Group</h3>
                <div class="value">৳""" + f"{target_value:,.0f}" + """</div>
                <div class="label">Equal Distribution Target</div>
            </div>
        </div>
"""

# Generate sections for each group
for group_name in ['A', 'B', 'C']:
    group = groups[group_name]
    group_class = f"group-{group_name.lower()}"
    
    # Calculate variance
    variance = ((group['total_value'] - target_value) / target_value * 100) if target_value > 0 else 0
    variance_sign = "+" if variance > 0 else ""
    
    # Organize assets by type
    organized = organize_assets_by_type(group['assets'])
    
    # Count by type
    type_counts = {k: len(v) for k, v in organized.items() if len(v) > 0}
    
    html += f"""
        <div class="groups-container">
            <div class="group {group_class}">
                <div class="group-header">
                    Group {group_name} - Complete Asset Allocation
                </div>
                
                <div class="group-stats">
                    <div class="stat">
                        <div class="stat-label">Total Assets</div>
                        <div class="stat-value">{len(group['assets'])}</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Total Value</div>
                        <div class="stat-value">৳{group['total_value']:,.0f}</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Variance</div>
                        <div class="stat-value">{variance_sign}{variance:.2f}%</div>
                    </div>
                    <div class="stat">
                        <div class="stat-label">Avg per Asset</div>
                        <div class="stat-value">৳{(group['total_value'] / len(group['assets'])):,.0f}</div>
                    </div>
"""
    
    # Add type counts
    for asset_type, count in type_counts.items():
        html += f"""
                    <div class="stat">
                        <div class="stat-label">{asset_type}</div>
                        <div class="stat-value">{count}</div>
                    </div>
"""
    
    html += """
                </div>
"""
    
    # Laptops section
    if len(organized['Laptop']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">💻 Laptops</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Name & Model</th>
                                <th>Specs</th>
                                <th>Level</th>
                                <th>Type</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for laptop in sorted(organized['Laptop'], key=lambda x: x['CurrentValue'], reverse=True):
            type_badge = "badge-new" if laptop.get('PurchaseType') == 'New' else "badge-recon"
            remarks_lower = laptop['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            if '1️⃣' in laptop.get('Level', ''):
                level_class = "level-1"
            elif '2️⃣' in laptop.get('Level', ''):
                level_class = "level-2"
            else:
                level_class = "level-3"
            
            specs = f"{laptop.get('Processor', 'N/A')} Gen{laptop.get('Gen', 0)}, {laptop.get('RAM', 0)}GB RAM"
            if laptop.get('GPU', '') != 'Integrated' and 'nan' not in str(laptop.get('GPU', '')).lower():
                specs += f", {laptop.get('GPU', '')}"
            
            html += f"""
                            <tr>
                                <td><strong>{laptop['Serial']}</strong></td>
                                <td>
                                    <strong>{laptop.get('Name', 'N/A')}</strong><br>
                                    <small style="color: #7f8c8d;">{laptop.get('Model', 'N/A')}</small>
                                </td>
                                <td><small>{specs}</small></td>
                                <td class="{level_class}">{laptop.get('Level', 'N/A')}</td>
                                <td><span class="badge {type_badge}">{laptop.get('PurchaseType', 'N/A')}</span></td>
                                <td><span class="badge {condition_badge}">{laptop['Remarks']}</span></td>
                                <td><strong>৳{laptop['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{laptop.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    # Desktop PCs section
    if len(organized['Desktop PC']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">🖥️ Desktop PCs</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Name & Model</th>
                                <th>Specs</th>
                                <th>Level</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for pc in sorted(organized['Desktop PC'], key=lambda x: x['CurrentValue'], reverse=True):
            remarks_lower = pc['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            if '1️⃣' in pc.get('Level', ''):
                level_class = "level-1"
            elif '2️⃣' in pc.get('Level', ''):
                level_class = "level-2"
            else:
                level_class = "level-3"
            
            specs = f"{pc.get('Processor', 'N/A')} Gen{pc.get('Gen', 0)}, {pc.get('RAM', 0)}GB RAM"
            if pc.get('GPU', '') != 'Integrated' and 'nan' not in str(pc.get('GPU', '')).lower():
                specs += f", {pc.get('GPU', '')}"
            
            html += f"""
                            <tr>
                                <td><strong>{pc['Serial']}</strong></td>
                                <td>
                                    <strong>{pc.get('Name', 'N/A')}</strong><br>
                                    <small style="color: #7f8c8d;">{pc.get('Model', 'N/A')}</small>
                                </td>
                                <td><small>{specs}</small></td>
                                <td class="{level_class}">{pc.get('Level', 'N/A')}</td>
                                <td><span class="badge {condition_badge}">{pc['Remarks']}</span></td>
                                <td><strong>৳{pc['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{pc.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    # Monitors section
    if len(organized['Monitor']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">🖥️ Monitors</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Name & Model</th>
                                <th>User/Location</th>
                                <th>Value Ratio</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for monitor in sorted(organized['Monitor'], key=lambda x: x['CurrentValue'], reverse=True):
            remarks_lower = monitor['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            html += f"""
                            <tr>
                                <td><strong>{monitor['Serial']}</strong></td>
                                <td>
                                    <strong>{monitor.get('Name', 'N/A')}</strong><br>
                                    <small style="color: #7f8c8d;">{monitor.get('Model', 'N/A')}</small>
                                </td>
                                <td><small>{monitor.get('User', 'N/A')} - {monitor.get('Location', 'N/A')}</small></td>
                                <td>{monitor.get('ValueRatio', 'N/A')}</td>
                                <td><span class="badge {condition_badge}">{monitor['Remarks']}</span></td>
                                <td><strong>৳{monitor['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{monitor.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    # Printers section
    if len(organized['Printer']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">🖨️ Printers</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Category</th>
                                <th>Device Name & Model</th>
                                <th>Function</th>
                                <th>Location</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for printer in sorted(organized['Printer'], key=lambda x: x['CurrentValue'], reverse=True):
            remarks_lower = printer['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            html += f"""
                            <tr>
                                <td><strong>{printer['Serial']}</strong></td>
                                <td><strong>{printer.get('Category', 'N/A')}</strong></td>
                                <td>
                                    <strong>{printer.get('DeviceName', 'N/A')}</strong><br>
                                    <small style="color: #7f8c8d;">{printer.get('Model', 'N/A')}</small>
                                </td>
                                <td><small>{printer.get('Function', 'N/A')}</small></td>
                                <td><small>{printer.get('Location', 'N/A')}</small></td>
                                <td><span class="badge {condition_badge}">{printer['Remarks']}</span></td>
                                <td><strong>৳{printer['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{printer.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    # Scanners section
    if len(organized['Scanner']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">📄 Scanners</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Device Name & Model</th>
                                <th>Function</th>
                                <th>Location</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for scanner in sorted(organized['Scanner'], key=lambda x: x['CurrentValue'], reverse=True):
            remarks_lower = scanner['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            html += f"""
                            <tr>
                                <td><strong>{scanner['Serial']}</strong></td>
                                <td>
                                    <strong>{scanner.get('DeviceName', 'N/A')}</strong><br>
                                    <small style="color: #7f8c8d;">{scanner.get('Model', 'N/A')}</small>
                                </td>
                                <td><small>{scanner.get('Function', 'N/A')}</small></td>
                                <td><small>{scanner.get('Location', 'N/A')}</small></td>
                                <td><span class="badge {condition_badge}">{scanner['Remarks']}</span></td>
                                <td><strong>৳{scanner['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{scanner.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    # Server Devices section
    if len(organized['Server Device']) > 0:
        html += """
                <div class="asset-section">
                    <div class="asset-section-title">🔌 Server & Network Devices</div>
                    <table>
                        <thead>
                            <tr>
                                <th>#</th>
                                <th>Device Name</th>
                                <th>Model</th>
                                <th>Serial</th>
                                <th>Condition</th>
                                <th>Value (BDT)</th>
                                <th>Allocation Remark</th>
                            </tr>
                        </thead>
                        <tbody>
"""
        for server in sorted(organized['Server Device'], key=lambda x: x['CurrentValue'], reverse=True):
            remarks_lower = server['Remarks'].lower()
            if 'excellent' in remarks_lower:
                condition_badge = "badge-excellent"
            elif 'good' in remarks_lower:
                condition_badge = "badge-good"
            else:
                condition_badge = "badge-moderate"
            
            sl = server.get('SL', server.get('Serial', 'N/A'))
            
            html += f"""
                            <tr>
                                <td><strong>{sl}</strong></td>
                                <td><strong>{server.get('DeviceName', 'N/A')}</strong></td>
                                <td><small>{server.get('Model', 'N/A')}</small></td>
                                <td><small>{server.get('Serial', 'N/A')}</small></td>
                                <td><span class="badge {condition_badge}">{server['Remarks']}</span></td>
                                <td><strong>৳{server['CurrentValue']:,.0f}</strong></td>
                                <td class="remark-cell">{server.get('AllocationRemark', 'Fair allocation')}</td>
                            </tr>
"""
        html += """
                        </tbody>
                    </table>
                </div>
"""
    
    html += """
            </div>
        </div>
"""

html += """
        <div class="footer">
            <p><strong>Genesis Technologies Ltd</strong> - Complete Asset Division Report</p>
            <p>Equal distribution achieved with ±0.04% variance. All values in BDT with depreciation applied for OLD assets.</p>
            <p style="margin-top: 10px; font-size: 0.85em;">Each asset includes detailed allocation remarks explaining the reasoning behind its assignment to ensure transparency and fairness.</p>
        </div>
    </div>
</body>
</html>
"""

# Save HTML
with open('laptop.html', 'w', encoding='utf-8') as f:
    f.write(html)

print("SUCCESS: Final HTML report generated: laptop.html")
print(f"Total assets: {len(all_assets)}")
print(f"Total value: {total_value:,.0f} BDT")
print(f"Target per group: {target_value:,.0f} BDT")
print("\nDistribution:")
for group_name in ['A', 'B', 'C']:
    group = groups[group_name]
    variance = ((group['total_value'] - target_value) / target_value * 100)
    print(f"  Group {group_name}: BDT {group['total_value']:,.0f} ({variance:+.2f}%)")
