"""Export index.html report data (from all_assets_data.json) to Excel."""
import json
import pandas as pd
from pathlib import Path

def main():
    path = Path(__file__).parent
    with open(path / 'all_assets_data.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    rows = []
    for group in ['A', 'B', 'C']:
        for asset in data['groups'][group]['assets']:
            row = {
                'Group': group,
                'Asset Type': asset.get('AssetType', ''),
                'Serial': asset.get('Serial', ''),
                'Name': asset.get('Name', ''),
                'Model': asset.get('Model', ''),
                'User': asset.get('User', ''),
                'Location': asset.get('Location', ''),
                'Processor': asset.get('Processor', ''),
                'RAM (GB)': asset.get('RAM', ''),
                'Storage': asset.get('Storage', ''),
                'Gen': asset.get('Gen', ''),
                'GPU': asset.get('GPU', ''),
                'Remarks': asset.get('Remarks', ''),
                'Purchase Type': asset.get('PurchaseType', ''),
                'Market Price (BDT)': asset.get('MarketPrice', ''),
                'Depreciation Rate': asset.get('DepreciationRate', ''),
                'Remark Category': asset.get('RemarkCategory', ''),
                'Current Value (BDT)': asset.get('CurrentValue', ''),
            }
            rows.append(row)

    df = pd.DataFrame(rows)
    asset_type_order = ['Laptop', 'PC', 'Monitor', 'Printer Scaneer', 'Server Room']
    df['_sort_key'] = df['Asset Type'].apply(lambda x: asset_type_order.index(x) if x in asset_type_order else 99)
    df = df.sort_values(['_sort_key', 'Group', 'Serial']).drop(columns=['_sort_key']).reset_index(drop=True)
    total_value = sum(data['groups'][g]['total_value'] for g in ['A', 'B', 'C'])
    target = total_value / 3

    summary_rows = []
    for g in ['A', 'B', 'C']:
        val = data['groups'][g]['total_value']
        cnt = data['groups'][g]['count']
        var = ((val - target) / target) * 100
        summary_rows.append({
            'Group': g,
            'Total Assets': cnt,
            'Total Value (BDT)': round(val, 2),
            'Variance %': round(var, 2),
        })
    total_assets = sum(data['groups'][g]['count'] for g in ['A', 'B', 'C'])
    summary_rows.append({
        'Group': 'TOTAL',
        'Total Assets': total_assets,
        'Total Value (BDT)': round(total_value, 2),
        'Variance %': '-',
    })
    df_summary = pd.DataFrame(summary_rows)

    out_path = path / 'index_report_export.xlsx'
    with pd.ExcelWriter(out_path, engine='openpyxl') as w:
        df_summary.to_excel(w, sheet_name='Summary', index=False)
        df.to_excel(w, sheet_name='All Assets', index=False)
        ws = w.sheets['All Assets']
        ws.column_dimensions['A'].width = 8
        ws.column_dimensions['B'].width = 14
        ws.column_dimensions['C'].width = 8
        ws.column_dimensions['D'].width = 24
        ws.column_dimensions['E'].width = 28
        ws.column_dimensions['F'].width = 24
        ws.column_dimensions['G'].width = 18
        for c in ['H', 'I', 'J', 'K', 'L', 'M', 'N']:
            ws.column_dimensions[c].width = 14
        ws.column_dimensions['O'].width = 18
        ws.column_dimensions['P'].width = 14
        ws.column_dimensions['Q'].width = 16
        ws.column_dimensions['R'].width = 18

    print(f'Exported {len(rows)} assets to {out_path.name}')
    print(f'  Sheets: Summary, All Assets')

if __name__ == '__main__':
    main()
