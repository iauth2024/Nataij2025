from django.apps import AppConfig

class ResultsConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'results'

    def ready(self):
        # Remove or comment out the call to load_excel_data()
        # from .views import load_excel_data
        # load_excel_data()  # This was causing the issue
        pass