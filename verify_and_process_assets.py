import pandas as pd
import json

# Read and verify all Excel files carefully
print("="*60)
print("VERIFYING ALL ASSET DATA")
print("="*60)

# Read PC data
print("\nReading PC.xlsx...")
df_pc = pd.read_excel('PC.xlsx')
print(f"PC file shape: {df_pc.shape}")
print("\nPC Data (rows 2-15):")
for idx in range(2, min(15, len(df_pc))):
    row = df_pc.iloc[idx]
    if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], (int, float)):
        serial = int(row.iloc[0])
        name = str(row.iloc[1]) if pd.notna(row.iloc[1]) else 'N/A'
        ram = row.iloc[3]
        processor = str(row.iloc[4]) if pd.notna(row.iloc[4]) else 'N/A'
        print(f"  Serial {serial}: {name} - RAM: {ram}, Processor: {processor}")

# Read Monitor data
print("\nReading Monitor.xlsx...")
df_mon = pd.read_excel('Monitor.xlsx')
print(f"Monitor file shape: {df_mon.shape}")

# Read Printer data
print("\nReading Printer Scaneer.xlsx...")
df_printer = pd.read_excel('Printer Scaneer.xlsx')
print(f"Printer file shape: {df_printer.shape}")

# Read Server data
print("\nReading Server.xlsx...")
df_server = pd.read_excel('Server.xlsx')
print(f"Server file shape: {df_server.shape}")

# Read Laptop data
print("\nReading Laptop.xlsx...")
df_laptop = pd.read_excel('Laptop.xlsx')
print(f"Laptop file shape: {df_laptop.shape}")

print("\n" + "="*60)
print("Data verification complete. Processing assets...")
print("="*60)
