import os
os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'purchase_order_project.settings')
import django
django.setup()
from purchase_order_app.models import Order, CommercialDetails

# Check DB data quality for earnings
print("=== DB DATA ANALYSIS FOR EARNINGS ===\n")

total = Order.objects.count()
with_commercial = Order.objects.filter(commercial_details__isnull=False).count()
with_rate = Order.objects.filter(commercial_details__rate__isnull=False).exclude(commercial_details__rate=0).count()
with_po_date = Order.objects.exclude(po_date__isnull=True).count()
with_all = Order.objects.exclude(po_date__isnull=True).filter(
    commercial_details__isnull=False,
    commercial_details__rate__isnull=False
).exclude(commercial_details__rate=0).count()

print(f"Total orders: {total}")
print(f"Orders with commercial_details: {with_commercial}")
print(f"Orders with rate != 0 and not null: {with_rate}")
print(f"Orders with po_date: {with_po_date}")
print(f"Orders with po_date AND rate: {with_all}")

print("\nSample commercial_details rate values:")
from decimal import Decimal
samples = CommercialDetails.objects.exclude(rate__isnull=True).exclude(rate=0)[:20]
for s in samples:
    print(f"  Order {s.order_id}: rate={s.rate}")

# Check orders with rate from April 2023
from datetime import date
april_2023 = date(2023, 4, 1)
orders_from_2023 = Order.objects.exclude(po_date__isnull=True).filter(
    po_date__gte=april_2023,
    commercial_details__isnull=False
).exclude(commercial_details__rate=0).exclude(commercial_details__rate__isnull=True).count()
print(f"\nOrders from April 2023 with valid rate: {orders_from_2023}")

# Check by year
from django.db.models import Count
orders_by_year = Order.objects.exclude(po_date__isnull=True).values('year').annotate(count=Count('id')).order_by('year')
print("\nOrders by year in DB:")
for item in orders_by_year:
    print(f"  {item['year']}: {item['count']} orders")
