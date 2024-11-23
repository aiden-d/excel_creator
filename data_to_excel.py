import xlsxwriter
import json
import os


def number_to_excel_char_index(number):
    return chr(number + 65)

def number_to_excel_numerical_index(number):
    return number + 1

def coordinates_to_excel(coordinates):
    return number_to_excel_char_index(coordinates[1]) + str(number_to_excel_numerical_index(coordinates[0]))


def get_data():
    script_path = os.path.abspath(__file__)
    script_dir = os.path.dirname(script_path)
    with open(f"{script_dir}/data.json", 'r') as file:
        data = json.load(file)
    return data

def write_data_to_excel(data, locations, name='output.xlsx'):
    workbook = xlsxwriter.Workbook(name)
    worksheet = workbook.add_worksheet()
    start_point = (2, 1)

    location_col_map = {}
    row = start_point[0]
    title_col = start_point[1]
    sub_title_col = start_point[1] + 1
    min_location_col = sub_title_col + 1
    max_location_col = min_location_col + len(locations) - 1

    total_gross_row = -1
    total_cogs_row = -1
    total_labour_row = -1
    gross_profit_row = -1

    def format_cell(row, col):
        json = {}
        json['num_format'] = '0.00'
        if row == start_point[0]:
            json['bottom'] = 1
            json['bold'] = True
        if col == min_location_col:
            json['left'] = 1
        if col == max_location_col:
            json['right'] = 1
        if row == total_gross_row or row == total_cogs_row or row == total_labour_row or row == total_gross_row + 1 or row == total_cogs_row + 1 or row == total_labour_row + 1:
            json['top'] = 1
            json['bottom'] = 1
            json['bold'] = True
        if col == title_col:
            json['bold'] = True
        if row == total_cogs_row + 1 or row == total_labour_row + 1 or row == gross_profit_row + 1:
            json['num_format'] = '0.0%'

        return workbook.add_format(json)
    
    def write(row, col, value):
        worksheet.write(row, col, value, format_cell(row, col))

    for i in range(start_point[1], start_point[1] + 2):
        write(start_point[0],  i, "")

    locations = sorted(locations)
    for i, location in enumerate(locations):
        col = sub_title_col + 1 + i
        write(row, col, location.capitalize())
        
        location_col_map[location] = col
    longest_location = len(max(locations, key=len))
    worksheet.set_column(start_point[1],  max_location_col + 1, longest_location)
    pipeline = ["income", "cogs", "labour"]
    

    left_bound = start_point[1]
    right_bound = sub_title_col + len(locations) + 1

    def skip_lines(n, row):
        for i in range(n):
            row += 1
            for col in range(left_bound, right_bound + 1):
                write(row, col, "")
            
        return row


    for pipeline_stage in pipeline:
        row = skip_lines(2, row)
        write(row, title_col, pipeline_stage.capitalize())
        sub_titles = set()
        for location, values in data.items():
            if pipeline_stage not in values:
                continue
            for sub_title, value in values[pipeline_stage].items():
                if sub_title not in sub_titles:
                    sub_titles.add(sub_title)
        sub_titles = sorted(list(sub_titles))

        for sub_title in sub_titles:
            write(row, sub_title_col, sub_title.capitalize())
            for location, values in data.items():
                if pipeline_stage not in values or sub_title not in values[pipeline_stage] or location not in location_col_map:
                    continue
                write(row, location_col_map[location], values[pipeline_stage][sub_title])

            write(row, max_location_col + 1, f"=SUM({coordinates_to_excel((row, min_location_col))}:{coordinates_to_excel((row, max_location_col))})")
            row += 1
        
        if pipeline_stage == "income":
            total_gross_row = row
        elif pipeline_stage == "cogs":
                total_cogs_row = row
        elif pipeline_stage == "labour":
            total_labour_row = row
        write(row, title_col, "Total gross" if pipeline_stage == "income" else "Total")
        write(row, title_col + 1, "")
        

        for i, location in enumerate(locations):
            col = location_col_map[location]
            write(row, col, f"=SUM({coordinates_to_excel((row - len(sub_titles), col))}:{coordinates_to_excel((row -1, col))})")       
        write(row, max_location_col + 1, f"=SUM({coordinates_to_excel((row, min_location_col))}:{coordinates_to_excel((row, max_location_col))})")
        
        if pipeline_stage == "income":
            
            row += 1
            write(row, title_col, "Total Net")
            write(row, title_col + 1, "")
            for i, location in enumerate(locations):
                col = location_col_map[location]
                write(row, col, f"={coordinates_to_excel((row -1, col))}/1.2")
            write(row, max_location_col + 1, f"=SUM({coordinates_to_excel((row, min_location_col))}:{coordinates_to_excel((row, max_location_col))})")
        else:
            
            row += 1
            write(row, title_col, "%")
            write(row, title_col + 1, "")
            for i, location in enumerate(locations):
                col = location_col_map[location]
                write(row, col, f"={coordinates_to_excel((row -1, col))}/{coordinates_to_excel((total_gross_row, col))}")
            write(row, max_location_col + 1, f"={coordinates_to_excel((row -1, max_location_col +1))}/{coordinates_to_excel((total_gross_row, max_location_col +1))}")
        
    row = skip_lines(2, row)
    gross_profit_row = row
    write(row, title_col, "Gross Profit")
    write(row, title_col + 1, "")
    for i, location in enumerate(locations):
        col = location_col_map[location]
        write(row, col, f"={coordinates_to_excel((total_gross_row, col))}-{coordinates_to_excel((total_cogs_row, col))}-{coordinates_to_excel((total_labour_row, col))}")
    write(row, max_location_col + 1, f"=SUM({coordinates_to_excel((row, min_location_col))}:{coordinates_to_excel((row, max_location_col))})")

    row += 1
    write(row, title_col, "%")
    write(row, title_col + 1, "")
    for i, location in enumerate(locations):
        col = location_col_map[location]
        write(row, col, f"={coordinates_to_excel((row -1, col))}/{coordinates_to_excel((total_gross_row, col))}")
    write(row, max_location_col + 1, f"={coordinates_to_excel((row -1, max_location_col +1))}/{coordinates_to_excel((total_gross_row, max_location_col +1))}")
    
    workbook.close()


    

def create_excel_file_v2(locations=["park street", "whiteladies road", "gloucester road", "gloucester road (TRUE)", "north street"]):    
    json = get_data()
    locations = json['locations']
    data = json['data']
    # For the cogs data colapse the xero label
    for location in data.keys():
        cogs_data = data[location]["cogs"]["xero"]
        data[location]["cogs"] = cogs_data
    write_data_to_excel(data, locations)
    

create_excel_file_v2()
