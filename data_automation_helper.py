
"""
Power BI Python Helper - Automates data transformation
"""

import pandas as pd
import os
from datetime import datetime

# Configuration
DATA_DIR = r"c:\Users\Vengala Bhaskar\Downloads\Python Folder\powerbi_data"

def load_data():
    """Load all CSV files"""
    print("Loading data...")
    
    orders = pd.read_csv(os.path.join(DATA_DIR, "Orders_Sample.csv"))
    people = pd.read_csv(os.path.join(DATA_DIR, "People_Sample.csv"))
    returns = pd.read_csv(os.path.join(DATA_DIR, "Returns_Sample.csv"))
    
    return orders, people, returns

def generate_statistics(orders, people, returns):
    """Generate statistical insights"""
    
    print("\n" + "=" * 80)
    print("DATA STATISTICS")
    print("=" * 80)
    
    print(f"\nORDERS TABLE:")
    print(f"  Total Records: {len(orders):,}")
    print(f"  Total Sales: ${orders['Sales'].sum():,.2f}")
    print(f"  Total Profit: ${orders['Profit'].sum():,.2f}")
    print(f"  Profit Margin: {(orders['Profit'].sum() / orders['Sales'].sum() * 100):.2f}%")
    print(f"  Average Order Value: ${orders['Sales'].mean():,.2f}")
    print(f"  Date Range: {orders['Order Date'].min()} to {orders['Order Date'].max()}")
    
    print(f"\nPEOPLE TABLE:")
    print(f"  Total Regional Managers: {len(people)}")
    print(f"  Regions: {', '.join(people['Region'].unique())}")
    
    print(f"\nRETURNS TABLE:")
    print(f"  Total Returns: {len(returns):,}")
    print(f"  Return Rate: {(len(returns) / len(orders['Order ID'].unique()) * 100):.2f}%")
    
    print("\n" + "=" * 80)

def export_insights():
    """Export insights to file"""
    
    orders, people, returns = load_data()
    
    insights_path = os.path.join(DATA_DIR, "Data_Insights.txt")
    
    with open(insights_path, 'w') as f:
        f.write("DATA INSIGHTS AND STATISTICS\n")
        f.write("=" * 80 + "\n\n")
        
        f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        
        f.write("ORDERS ANALYSIS\n")
        f.write("-" * 80 + "\n")
        f.write(f"Total Orders: {len(orders)}\n")
        f.write(f"Total Sales: ${orders['Sales'].sum():,.2f}\n")
        f.write(f"Total Profit: ${orders['Profit'].sum():,.2f}\n")
        f.write(f"Profit Margin: {(orders['Profit'].sum() / orders['Sales'].sum() * 100):.2f}%\n")
        f.write(f"Average Sales per Order: ${orders['Sales'].mean():,.2f}\n")
        f.write(f"Total Quantity: {int(orders['Quantity'].sum())}\n")
        f.write(f"Average Discount: {(orders['Discount'].mean() * 100):.2f}%\n\n")
        
        f.write("SALES BY CATEGORY\n")
        f.write("-" * 80 + "\n")
        for category, group in orders.groupby('Category'):
            f.write(f"{category}: ${group['Sales'].sum():,.2f} ({len(group)} orders)\n")
        
        f.write("\nSALES BY REGION\n")
        f.write("-" * 80 + "\n")
        for region, group in orders.groupby('Region'):
            f.write(f"{region}: ${group['Sales'].sum():,.2f} ({len(group)} orders)\n")
        
        f.write("\nSALES BY SEGMENT\n")
        f.write("-" * 80 + "\n")
        for segment, group in orders.groupby('Segment'):
            f.write(f"{segment}: ${group['Sales'].sum():,.2f} ({len(group)} orders)\n")
    
    print(f"\nInsights exported to: {insights_path}")

if __name__ == "__main__":
    orders, people, returns = load_data()
    generate_statistics(orders, people, returns)
    export_insights()
