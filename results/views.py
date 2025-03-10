import os
import pandas as pd
from django.shortcuts import render
from django.conf import settings
import logging

logger = logging.getLogger(__name__)

EXCEL_FILE_PATH = os.getenv('EXCEL_FILE_PATH', os.path.join(settings.BASE_DIR, 'results', 'Nateeja.xlsx'))

def load_excel_data():
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            logger.error("Excel file not found at %s", EXCEL_FILE_PATH)
            raise FileNotFoundError("Excel file not found.")
        excel_file = pd.ExcelFile(EXCEL_FILE_PATH)
        data = {}
        for sheet in excel_file.sheet_names:
            df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet)
            df.columns = [str(col) for col in df.columns]
            df = df.loc[:, ~df.columns.str.contains(r"^Unnamed", na=False)]
            rename_dict = {col: "درجہ (Rank)" for col in df.columns if "درجہ" in col and "1" in col}
            df.rename(columns=rename_dict, inplace=True)
            data[sheet] = df
        logger.info("Excel data loaded successfully with %d sheets", len(data))
        return data
    except Exception as e:
        logger.exception("Failed to load Excel data: %s", str(e))
        raise

def search_excel(request):
    if request.method == 'POST':
        try:
            excel_data = load_excel_data()
            search_value_raw = request.POST.get('search_value', '').strip()
            if not search_value_raw:
                raise ValueError("Please enter an admission number.")
            search_value = int(search_value_raw)
            logger.info("Searching for admission number: %d", search_value)

            for sheet_name, df in excel_data.items():
                if 'داخلہ نمبر' in df.columns:
                    result = df[df['داخلہ نمبر'] == search_value]
                    if not result.empty:
                        result_dict = result.iloc[0].to_dict()
                        # Rest of your processing logic...
                        return render(request, 'search_results.html', context)

            message = f"کوئی طالب علم داخلہ نمبر '{search_value}' کے ساتھ کسی شیٹ میں نہیں ملا۔"
        except FileNotFoundError:
            message = "Excel فائل مقررہ راستے پر نہیں ملی۔"
        except ValueError:
            message = "براہ کرم ایک درست داخلہ نمبر درج کریں۔"
        except Exception as e:
            logger.exception("Error processing request: %s", str(e))
            message = f"ایک خرابی پیش آئی: {str(e)}"

        context = {'message': message}
        return render(request, 'search_results.html', context)

    return render(request, 'search_form.html')