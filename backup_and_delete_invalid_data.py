"""
Backup and delete orders outside valid date range
Valid range: April 1, 2024 to Today (Feb 3, 2026)
WARNING: This will permanently delete data!
"""
import os
import django
import json
from datetime import datetime

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'purchase_order_project.settings')
django.setup()

from purchase_order_app.models import Order

# Valid date range
CUTOFF_START = datetime(2024, 4, 1).date()  # April 1, 2024
CUTOFF_END = datetime.now().date()  # Today

print("=" * 80)
print("DATA CLEANUP: Remove orders outside valid date range")
print(f"Valid range: {CUTOFF_START} to {CUTOFF_END}")
print("=" * 80)

# Get orders to delete (before April 2024 OR after today)
orders_before = Order.objects.filter(po_date__lt=CUTOFF_START)
orders_after = Order.objects.filter(po_date__gt=CUTOFF_END)

count_before = orders_before.count()
count_after = orders_after.count()
total_to_delete = count_before + count_after

print(f"\nOrders before April 2024: {count_before}")
print(f"Orders after today (future): {count_after}")
print(f"Total to delete: {total_to_delete}")

if total_to_delete == 0:
    print("No orders to delete. Exiting.")
    exit(0)

# Create backup
print("\n1. Creating backup...")
backup_data = {
    'backup_date': datetime.now().isoformat(),
    'valid_range': {
        'start': str(CUTOFF_START),
        'end': str(CUTOFF_END)
    },
    'orders_before_april_2024': [],
    'orders_future_dated': []
}

# Backup old orders
for order in orders_before:
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
    
    backup_data['orders_before_april_2024'].append(order_data)

# Backup future orders
for order in orders_after:
    order_data = {
        'order_id': order.id,
        'order_form': order.order_form,
        'reference_format_no': order.reference_format_no,
        'po_reference': order.po_reference,
        'po_date': str(order.po_date) if order.po_date else None,
        'year': order.year,
        'order_type': order.order_type,
    }
    
    if hasattr(order, 'product_details'):
        order_data['product'] = {
            'brand_name': order.product_details.brand_name,
            'generic_name': order.product_details.generic_name,
        }
    
    if hasattr(order, 'client_dispatch'):
        order_data['client'] = {
            'company_name': order.client_dispatch.company_name,
        }
    
    backup_data['orders_future_dated'].append(order_data)

# Save backup
backup_file = f'backup_invalid_data_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
with open(backup_file, 'w', encoding='utf-8') as f:
    json.dump(backup_data, f, indent=2, ensure_ascii=False)

print(f"   Backup saved to: {backup_file}")
print(f"   - Old orders backed up: {count_before}")
print(f"   - Future orders backed up: {count_after}")

# Ask for confirmation
print(f"\n2. Ready to delete {total_to_delete} orders")
print("   This action CANNOT be undone!")
print("   Type 'DELETE' to proceed, or anything else to cancel: ")

confirmation = input().strip()

if confirmation == 'DELETE':
    print("\n3. Deleting orders...")
    
    # Delete old orders
    if count_before > 0:
        deleted_count, deleted_details = orders_before.delete()
        print(f"   Old orders deleted: {deleted_count} records")
    
    # Delete future orders
    if count_after > 0:
        deleted_count, deleted_details = orders_after.delete()
        print(f"   Future orders deleted: {deleted_count} records")
    
    # Verify deletion
    remaining_invalid = Order.objects.filter(po_date__lt=CUTOFF_START).count() + Order.objects.filter(po_date__gt=CUTOFF_END).count()
    remaining_valid = Order.objects.filter(po_date__gte=CUTOFF_START, po_date__lte=CUTOFF_END).count()
    
    print("\n4. DONE!")
    print(f"   Backup file: {backup_file}")
    print(f"   Total deleted: {total_to_delete}")
    print(f"   Remaining invalid orders: {remaining_invalid}")
    print(f"   Remaining valid orders: {remaining_valid}")
else:
    print("\n   Cancelled. No data was deleted.")

print("\n" + "=" * 80)
