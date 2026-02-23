import os
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'purchase_order_project.settings')
import django
django.setup()
from purchase_order_app.models import Order
from datetime import date

# Check po_date distribution
from django.db.models import Min, Max

print("=== PO DATE ANALYSIS ===")
earliest = Order.objects.exclude(po_date__isnull=True).aggregate(Min('po_date'))
latest = Order.objects.exclude(po_date__isnull=True).aggregate(Max('po_date'))
print(f"Earliest po_date: {earliest}")
print(f"Latest po_date: {latest}")

# Orders by month
from django.db.models.functions import TruncMonth
from django.db.models import Count
monthly = Order.objects.exclude(po_date__isnull=True).annotate(
    month=TruncMonth('po_date')
).values('month').annotate(count=Count('id')).order_by('month')[:30]
print("\nOrders by month (first 30):")
for item in monthly:
    print(f"  {item['month'].strftime('%Y-%m')}: {item['count']}")

# Check year field vs po_date year mismatch
print("\n=== ORDERS WITH 2023 DATE ===")
orders_2023 = Order.objects.filter(po_date__year=2023).count()
print(f"Orders with po_date in 2023: {orders_2023}")

orders_2023_q1 = Order.objects.filter(po_date__gte=date(2023,4,1), po_date__lt=date(2024,4,1)).count()
print(f"Orders from Apr 2023 to Mar 2024: {orders_2023_q1}")

print("\nFirst 5 orders with 2023 dates:")
for o in Order.objects.filter(po_date__year=2023).select_related('commercial_details')[:5]:
    rate = getattr(o.commercial_details, 'rate', None) if hasattr(o, 'commercial_details') else None
    print(f"  id={o.id}, year={o.year}, po_date={o.po_date}, rate={rate}")
