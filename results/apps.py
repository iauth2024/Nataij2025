from django.apps import AppConfig

class ResultsConfig(AppConfig):
    name = 'results'
    def ready(self):
        from .views import load_excel_data
        load_excel_data()