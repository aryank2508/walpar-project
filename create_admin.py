
import os
import django

os.environ.setdefault('DJANGO_SETTINGS_MODULE', 'purchase_order_project.settings')
django.setup()

from django.contrib.auth import get_user_model
User = get_user_model()

username = 'admin'
password = 'admin123'
email = 'admin@example.com'

if not User.objects.filter(username=username).exists():
    User.objects.create_superuser(username, email, password)
    print(f"Successfully created superuser:\nUsername: {username}\nPassword: {password}")
else:
    print(f"Superuser '{username}' already exists. If you don't know the password, I can reset it for you.")
    # Optional: Reset password just in case they forgot it
    # user = User.objects.get(username=username)
    # user.set_password(password)
    # user.save()
    # print(f"Reset password for '{username}' to '{password}'")
