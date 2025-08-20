from django.apps import AppConfig


class AccountsConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'accounts'

    def ready(self):
        # Ensure profile exists for new users (signals could be added later)
        pass

    def ready(self):
        # Ensure profile exists for new users (signals could be added later)
        pass