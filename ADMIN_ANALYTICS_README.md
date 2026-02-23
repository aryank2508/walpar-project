# Admin Analytics Dashboard

A comprehensive analytics dashboard for the Purchase Order Management System that provides monthly and yearly statistics.

## Access

The analytics dashboard is accessible at:
- **URL**: `/admin/analytics/`
- **Direct Link**: `http://localhost:8000/admin/analytics/`

**Note**: You must be logged in as a staff/admin user to access this page.

## Features

### 📊 Summary Statistics
- **Total Orders**: Count of all orders in the system
- **Years Covered**: Number of unique years with order data
- **This Month**: Orders placed in the current month with growth percentage vs. previous month
- **Average per Month**: Average number of orders per month

### 📅 Charts and Visualizations

1. **Orders by Year (Bar Chart)**
   - Shows the distribution of orders across different years
   - Displays count for each year

2. **Orders by Month - Last 12 Months (Line Chart)**
   - Trend analysis showing order volume over the past year
   - Helps identify seasonal patterns

3. **Top Order Types (Doughnut Chart)**
   - Visual breakdown of orders by type
   - Shows top 10 order types

### 🔍 Filtering Options

The dashboard includes filters to analyze specific data ranges:

- **Filter by Year**: Select a specific year to view data for that year only
- **Date From**: Start date for custom date range filtering
- **Date To**: End date for custom date range filtering
- **Clear Filters**: Reset all filters to show all data

## Usage

1. **Access the Dashboard**:
   - Log in to Django admin (`/admin/`)
   - Navigate to `/admin/analytics/` or click the analytics link if available

2. **Apply Filters** (Optional):
   - Select a year from the dropdown to filter by year
   - Or specify a custom date range using "Date From" and "Date To"
   - Click "Apply Filters" to update the charts

3. **View Statistics**:
   - Review the summary cards at the top for quick insights
   - Scroll down to see detailed charts
   - Charts are interactive (hover for detailed information)

4. **Clear Filters**:
   - Click "Clear" button to reset filters and view all data

## Technical Details

### View Function
- **File**: `purchase_order_app/admin_views.py`
- **Function**: `admin_analytics_dashboard()`
- **Decorator**: `@staff_member_required` (requires admin/staff login)

### Template
- **File**: `purchase_order_app/templates/purchase_order_app/admin_analytics.html`
- **Extends**: Django admin base template
- **Styling**: `purchase_order_app/static/purchase_order_app/css/admin_analytics.css`

### Charts Library
- Uses **Chart.js** (loaded via CDN)
- Chart types: Bar, Line, Doughnut

### Data Aggregation
- Uses Django ORM aggregation functions (`Count`, `TruncMonth`, `TruncYear`)
- Optimized queries with proper filtering
- Handles edge cases (null dates, empty datasets)

## URL Configuration

The analytics dashboard is registered in `purchase_order_project/urls.py`:

```python
path('admin/analytics/', admin_analytics_dashboard, name='admin_analytics'),
```

## Permissions

Only staff users (users with `is_staff=True`) can access the analytics dashboard. This is enforced by the `@staff_member_required` decorator.

## Future Enhancements

Potential additions:
- Export analytics data to Excel/PDF
- More granular time period options (weekly, quarterly)
- Revenue/Commercial analysis charts
- Client/Company-wise statistics
- Product-wise analysis
- Comparative year-over-year analysis

