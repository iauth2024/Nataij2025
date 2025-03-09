import os
import pandas as pd
from django.shortcuts import render
import logging

logger = logging.getLogger(__name__)

EXCEL_FILE_PATH = "C:/nataij/results/Nateeja.xlsx"
EXCEL_DATA = None

def load_excel_data():
    """Loads Excel file data into memory."""
    global EXCEL_DATA
    try:
        if not os.path.exists(EXCEL_FILE_PATH):
            logger.error("Excel file not found at %s", EXCEL_FILE_PATH)
            raise FileNotFoundError("Excel file not found at the specified path.")
        
        excel_file = pd.ExcelFile(EXCEL_FILE_PATH)
        EXCEL_DATA = {
            sheet: pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet) for sheet in excel_file.sheet_names
        }
        logger.info("Excel data loaded successfully with %d sheets", len(EXCEL_DATA))
    
    except Exception as e:
        logger.exception("Failed to load Excel data: %s", str(e))
        EXCEL_DATA = None  # Avoid breaking the app
        raise

# Load data initially if not already loaded
try:
    load_excel_data()
except Exception as e:
    logger.warning("Excel file not loaded at startup: %s", str(e))

def search_excel(request):
    """Handles the search functionality for admission numbers in Excel."""
    if request.method == 'POST':
        try:
            search_value_raw = request.POST.get('search_value', '').strip()
            if not search_value_raw:
                raise ValueError("Please enter an admission number.")
            
            search_value = int(search_value_raw)
            logger.info("Searching for admission number: %d", search_value)

            # Reload data if it was not loaded earlier
            if EXCEL_DATA is None:
                logger.warning("Excel data is None. Reloading now...")
                load_excel_data()
                if EXCEL_DATA is None:
                    raise FileNotFoundError("Excel data not available.")

            for sheet_name, df in EXCEL_DATA.items():
                if df.empty:
                    logger.warning("Skipping empty sheet: %s", sheet_name)
                    continue
                if 'داخلہ نمبر' not in df.columns:
                    logger.warning("Skipping sheet %s as it lacks 'داخلہ نمبر' column.", sheet_name)
                    continue

                # Work on a copy to avoid modifying cached data
                df_copy = df.copy()
                
                # Remove unnamed columns
                df_copy = df_copy.loc[:, ~df_copy.columns.str.contains("^Unnamed")]

                # Rename rank column if necessary
                rename_dict = {col: "درجہ (Rank)" for col in df_copy.columns if "درجہ" in col and "1" in col}
                df_copy = df_copy.rename(columns=rename_dict)

                # Search for the student
                result = df_copy[df_copy['داخلہ نمبر'] == search_value]
                if not result.empty:
                    result_dict = result.iloc[0].to_dict()

                    # Remove unnecessary columns
                    result_dict.pop("جائزہ اوسط", None)
                    result_dict.pop("اوسط نمبر", None)
                    result_dict.pop("اوسط نمبر.1", None)

                    # Format numerical values
                    for key, value in result_dict.items():
                        if isinstance(value, (int, float)) and key not in ["داخلہ نمبر", "رول نمبر"]:
                            if key == "کل اوسط":
                                result_dict[key] = "{:.2f}".format(value)
                            elif pd.notna(value):
                                result_dict[key] = int(value)

                    if "درجہ.1" in result_dict:
                        result_dict["درجہ (Rank)"] = result_dict.pop("درجہ.1")

                    # Organize sections
                    top_section_keys = ["ہال ٹکٹ نمبر", "داخلہ نمبر", "شعبہ", "درجہ", "نمبر شمار", "نام طالب علم"]
                    bottom_section_keys = ["کل نمبرات", "فیصد", "درجۂ کامیابی", "پوزیشن", "امتیازی پوزیشن", "درجہ (Rank)", "کل اوسط"]
                    visible_columns = top_section_keys + bottom_section_keys

                    columns = df_copy.columns.tolist()
                    top_section_data = {key: result_dict.get(key, 'N/A') for key in top_section_keys if key in columns}
                    bottom_section_data = {key: result_dict.get(key, 'N/A') for key in bottom_section_keys if key in columns}
                    middle_section_data = {
                        key: value for key, value in result_dict.items()
                        if key not in visible_columns and key in columns and pd.notna(value)
                    }

                    # Render the results
                    context = {
                        'top_section_data': top_section_data,
                        'middle_section_data': middle_section_data,
                        'bottom_section_data': bottom_section_data,
                        'class_name': sheet_name,
                        'search_value': search_value
                    }
                    logger.info("Rendering results for %d", search_value)
                    return render(request, 'search_results.html', context)

            message = f"کوئی طالب علم داخلہ نمبر '{search_value}' کے ساتھ کسی شیٹ میں نہیں ملا۔"

        except FileNotFoundError:
            message = "Excel فائل مقررہ راستے پر نہیں ملی۔"
        except ValueError:
            message = "براہ کرم ایک درست داخلہ نمبر درج کریں۔"
        except Exception as e:
            logger.exception("Error processing request: %s", str(e))
            message = f"ایک خرابی پیش آئی: {str(e)}"

        # Render message in case of errors
        context = {'message': message}
        return render(request, 'search_results.html', context)

    return render(request, 'search_form.html')
