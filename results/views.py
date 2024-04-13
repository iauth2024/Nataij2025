from django.shortcuts import render
import pandas as pd


def search_excel(request):
    if request.method == 'POST':
        search_value = int(request.POST.get('search_value'))
        sheet_name = request.POST.get('sheet_name')  # Retrieve selected sheet name
        file_path = 'c:\\Nateeja.xlsx'

        try:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            columns = df.columns.tolist()  # Get column names
            result = df[df['داخلہ نمبر'] == search_value]  # Assuming 'داخلہ نمبر' is the search column
            if not result.empty:
                # Convert DataFrame to dictionary for passing to template
                result_dict = result.to_dict(orient='records')[0]

                # Format the percentage value to one decimal place with '%' sign
                result_dict['فیصد'] = "{:.1f}%".format(result_dict['فیصد'])

                # Split the data into top, middle, and bottom sections
                top_section_data = {}
                middle_section_data = {}
                bottom_section_data = {}

                for key, value in result_dict.items():
                    if key in ["نمبر شمار", "داخلہ نمبر", "رول نمبر", "نام", "جماعت"]:
                        top_section_data[key] = value
                    elif key in ["کل نمبرات", "فیصد", "درجۂ کامیابی", "پوزیشن","امتیازی پوزیشن", "امتیازی پوزیش"]:
                        bottom_section_data[key] = value
                    else:
                        middle_section_data[key] = value

                context = {
                    'columns': columns,
                    'top_section_data': top_section_data,
                    'middle_section_data': middle_section_data,
                    'bottom_section_data': bottom_section_data
                }
                return render(request, 'excel_search/search_results.html', context)
            else:
                message = f"No results found in sheet '{sheet_name}' for search value '{search_value}'."
        except ValueError:
            message = f"Worksheet named '{sheet_name}' not found."

        context = {
            'message': message
        }
        return render(request, 'excel_search/search_results.html', context)

    sheet_names = [
        'Ifta',
        'Daura Hadees',
        'Mouqoof Alaih (Alif)',
        'Mouqoof Alaih (Baa)',
        'Panjum',
        'Charum',
        'Suwwam Alif',
        'Suwwam Baa',
        'Duwwam  Alif',
        'Duwwam Baa',
        'Duwwam Jeem',
        'Duwwam Madarsa wa School',
        'Awwal Alif',
        'Awwal Aalimiyat Baa',
        'Awwal Aalimiyat jeem',
        'Awwal Madarsa Series',
        'Awwal School series',
        'Idadiyah Madarsa Series',
        'Idadiyah School Series'
    ]
    return render(request, 'excel_search/search_form.html', {'sheet_names': sheet_names})
