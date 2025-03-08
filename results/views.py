import os
import pandas as pd
from django.shortcuts import render
from django.conf import settings

def search_excel(request):
    # Use a relative path based on BASE_DIR for portability
    file_path = os.path.join(settings.BASE_DIR, 'results', 'Nateeja.xlsx')

    if request.method == 'POST':
        try:
            # Get search value and convert to integer
            search_value_raw = request.POST.get('search_value', '').strip()
            if not search_value_raw:
                raise ValueError("Please enter an admission number.")
            search_value = int(search_value_raw)

            # Check if Excel file exists
            if not os.path.exists(file_path):
                raise FileNotFoundError("Excel file not found at the specified path.")

            # Read Excel file with multiple sheets
            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names

            # Iterate through each sheet
            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                # Remove unnamed columns
                df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
                
                # Rename columns with "درجہ" and "1" to "درجہ (Rank)"
                rename_dict = {col: "درجہ (Rank)" for col in df.columns if "درجہ" in col and "1" in col}
                df.rename(columns=rename_dict, inplace=True)

                # Search for student by admission number
                if 'داخلہ نمبر' in df.columns:
                    result = df[df['داخلہ نمبر'] == search_value]
                    if not result.empty:
                        result_dict = result.iloc[0].to_dict()

                        # Remove specific average fields, keep "کل اوسط"
                        result_dict.pop("جائزہ اوسط", None)
                        result_dict.pop("اوسط نمبر", None)
                        result_dict.pop("اوسط نمبر.1", None)

                        # Format numerical values
                        for key, value in result_dict.items():
                            if isinstance(value, (int, float)) and key not in ["داخلہ نمبر", "رول نمبر"]:
                                if key == "کل اوسط":  # Two decimal places for "کل اوسط"
                                    result_dict[key] = "{:.2f}".format(value)
                                elif pd.notna(value):  # Convert other numbers to integers if not NaN
                                    result_dict[key] = int(value)

                        # Handle "درجہ.1" renaming
                        if "درجہ.1" in result_dict:
                            result_dict["درجہ (Rank)"] = result_dict.pop("درجہ.1")

                        # Define sections
                        top_section_keys = ["ہال ٹکٹ نمبر", "داخلہ نمبر", "شعبہ", "درجہ", "نمبر شمار", "نام طالب علم"]
                        bottom_section_keys = ["کل نمبرات", "فیصد", "درجۂ کامیابی", "پوزیشن", "امتیازی پوزیشن", "درجہ (Rank)", "کل اوسط"]
                        visible_columns = top_section_keys + bottom_section_keys

                        columns = df.columns.tolist()
                        top_section_data = {key: result_dict.get(key, 'N/A') for key in top_section_keys if key in columns}
                        bottom_section_data = {key: result_dict.get(key, 'N/A') for key in bottom_section_keys if key in columns}
                        middle_section_data = {key: value for key, value in result_dict.items() 
                                             if key not in visible_columns and key in columns and pd.notna(value)}

                        context = {
                            'top_section_data': top_section_data,
                            'middle_section_data': middle_section_data,
                            'bottom_section_data': bottom_section_data,
                            'class_name': sheet_name,  # Sheet name as class name
                            'search_value': search_value
                        }
                        return render(request, 'search_results.html', context)

            message = f"کوئی طالب علم داخلہ نمبر '{search_value}' کے ساتھ کسی شیٹ میں نہیں ملا۔"

        except FileNotFoundError as e:
            message = "Excel فائل مقررہ راستے پر نہیں ملی۔"
        except ValueError:
            message = "براہ کرم ایک درست داخلہ نمبر درج کریں۔"
        except Exception as e:
            message = f"ایک خرابی پیش آئی: {str(e)}"

        context = {'message': message}
        return render(request, 'search_results.html', context)

    # Render search form for GET requests
    return render(request, 'search_form.html')