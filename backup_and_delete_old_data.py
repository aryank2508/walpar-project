"""
Backup and delete orders before April 2024
WARNING: This will permanently delete data!
"""
import os
import django
import json
from datetime import datetime

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'purchase_order_project.settings')
django.setup()

from purchase_order_app.models import Order, ProductDetails, PackagingDetails, ArtworkApproval, CommercialDetails, ClientDispatch

# Cutoff date: April 1, 2024
CUTOFF_DATE = datetime(2024, 4, 1).date()

print("=" * 80)
print("DATA CLEANUP: Remove orders before April 2024")
print("=" * 80)

# Get orders to delete
orders_to_delete = Order.objects.filter(po_date__lt=CUTOFF_DATE)
count_to_delete = orders_to_delete.count()

print(f"\nOrders to delete: {count_to_delete}")

if count_to_delete == 0:
    print("No orders to delete. Exiting.")
    exit(0)

# Create backup
print("\n1. Creating backup...")
backup_data = []

for order in orders_to_delete:
    order_data = {
        'order_id': order.id,
        'order_form': order.order_form,
        'reference_format_no': order.reference_format_no,
        'po_reference': order.po_reference,
        'po_date': str(order.po_date) if order.po_date else None,
        'year': order.year,
        'order_type': order.order_type,
    }
    
    # Add related data
    if hasattr(order, 'product_details'):
        order_data['product'] = {
            'brand_name': order.product_details.brand_name,
            'generic_name': order.product_details.generic_name,
        }
    
    if hasattr(order, 'client_dispatch'):
        order_data['client'] = {
            'company_name': order.client_dispatch.company_name,
        }
    
    backup_data.append(order_data)

# Save backup
backup_file = f'backup_pre_april_2024_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
with open(backup_file, 'w', encoding='utf-8') as f:
    json.dump(backup_data, f, indent=2, ensure_ascii=False)

print(f"   Backup saved to: {backup_file}")

# Ask for confirmation
print(f"\n2. Ready to delete {count_to_delete} orders")
print("   This action CANNOT be undone!")
print("   Type 'DELETE' to proceed, or anything else to cancel: ")

confirmation = input().strip()

if confirmation == 'DELETE':
    print("\n3. Deleting orders...")
    deleted_count, deleted_details = orders_to_delete.delete()
    print(f"   Deleted {deleted_count} records:")
    for model, count in deleted_details.items():
        print(f"   - {model}: {count}")
    
    print("\n4. DONE!")
    print(f"   Backup file: {backup_file}")
    print(f"   Deleted orders: {count_to_delete}")
else:
    print("\n   Cancelled. No data was deleted.")

print("\n" + "=" * 80)
