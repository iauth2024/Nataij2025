import os
import pandas as pd
from django.shortcuts import render

def search_excel(request):
    file_path = 'C:\\nataij\\results\\Nateeja.xlsx'

    if request.method == 'POST':
        try:
            search_value = int(request.POST.get('search_value'))

            if not os.path.exists(file_path):
                raise FileNotFoundError("Excel file not found at the specified path.")

            excel_file = pd.ExcelFile(file_path)
            sheet_names = excel_file.sheet_names

            for sheet_name in sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet_name)
                df = df.loc[:, ~df.columns.str.contains("^Unnamed")]
                
                rename_dict = {col: "درجہ (Rank)" for col in df.columns if "درجہ" in col and "1" in col}
                df.rename(columns=rename_dict, inplace=True)

                if 'داخلہ نمبر' in df.columns:
                    result = df[df['داخلہ نمبر'] == search_value]
                    if not result.empty:
                        result_dict = result.iloc[0].to_dict()

                        # Remove only specific average fields, keep "کل اوسط"
                        result_dict.pop("جائزہ اوسط", None)
                        result_dict.pop("اوسط نمبر", None)  # Remove "اوسط نمبر"
                        result_dict.pop("اوسط نمبر.1", None)  # Remove "اوسط نمبر.1"

                        # Format numerical values
                        for key, value in result_dict.items():
                            if isinstance(value, (int, float)) and key not in ["داخلہ نمبر", "رول نمبر"]:
                                if key == "کل اوسط":  # Format "کل اوسط" with two decimal places
                                    result_dict[key] = "{:.2f}".format(value)
                                else:  # Convert other numbers to integers
                                    result_dict[key] = int(value)

                        if "درجہ.1" in result_dict:
                            result_dict["درجہ (Rank)"] = result_dict.pop("درجہ.1")

                        top_section_keys = ["ہال ٹکٹ نمبر", "داخلہ نمبر", "شعبہ", "درجہ", "نمبر شمار", "نام طالب علم"]
                        # Add "کل اوسط" to bottom section
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
                            'class_name': sheet_name,
                            'search_value': search_value
                        }
                        return render(request, 'excel_search/search_results.html', context)

            message = f"No student found with admission number '{search_value}' in any sheet."

        except FileNotFoundError as e:
            message = str(e)
        except ValueError:
            message = "Please enter a valid admission number."
        except Exception as e:
            message = f"An error occurred: {str(e)}"

        context = {'message': message}
        return render(request, 'excel_search/search_results.html', context)

    return render(request, 'excel_search/search_form.html')