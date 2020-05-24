# Function that extract the data for a given row before outputing some CSV.
def generate_csv_row(project_id, csi_division, csi_sub_division, row_data):
    project_id = project_id
    csi_number = csi_sub_division[0:8]  # first 6 digits of CIS SUB
    csi_division = csi_division
    csi_subdivision = csi_sub_division
    category = csi_division
    subcategory = csi_subdivision
    item_code = row_data['RSMeans 12-digit code']
    activity = None
    ddc_qty = row_data['QUANT']
    ddc_unit = row_data['UNIT']
    ddc_material_cost = row_data['TOTAL MAT. $:']
    ddc_labor_cost = row_data['TOTAL LABOR $:']
    # extra space for the column name. Typo?
    ddc_equipment_cost = row_data['TOTAL  EQUIP $:']
    ddc_markup = None  # recompute in Python
    ddc_unit_cost = None  # recompute in Python
    ddc_avg_unit_price = None  # recompute in Python
    ddc_extended_total_cost = None  # recompute in Python
    bid1_qty = row_data['QUANT.1']
    bid1_unit = row_data['UNIT COST.1']
    bid1_material_cost = row_data['TOTAL MAT. $:.1']
    bid1_labor_cost = row_data['TOTAL LABOR $:.1']
    bid1_equipment_cost = row_data['TOTAL  EQUIP $:.1']
    bid1_unit_cost = None  # recompute in Python
    bid1_ext_total_cost = None  # recompute in Python
    bid1_variance = None  # -- recompute in Python
    bid2_qty = row_data['QUANT.2']
    bid2_unit = row_data['UNIT COST.2']
    bid2_material_cost = row_data['TOTAL MAT. $:.2']
    bid2_labor_cost = row_data['TOTAL LABOR $:.2']
    bid2_equipment_cost = row_data['TOTAL  EQUIP $:.2']
    bid2_unit_cost = None  # recompute in Python
    bid2_ext_total_cost = None  # recompute in Python
    bid2_variance = None  # -- recompute in Python
    bid3_qty = row_data['QUANT.3']
    bid3_unit = row_data['UNIT COST.3']
    bid3_material_cost = row_data['TOTAL MAT. $:.3']
    bid3_labor_cost = row_data['TOTAL LABOR $:.3']
    bid3_equipment_cost = row_data['TOTAL  EQUIP $:.3']
    bid3_unit_cost = None  # recompute in Python
    bid3_ext_total_cost = None  # recompute in Python
    bid3_variance = None  # -- recompute in Python

    # Now that we have all the values, we recompute some of them.
    # TODO: add all of them
    #ddc_avg_unit_price = (bid1_unit_cost + bid2_unit_cost + bid3_unit_cost) * 1/3
    csv_row = (project_id, csi_number, csi_division, csi_subdivision, category, subcategory, item_code, activity, ddc_qty, ddc_unit, ddc_material_cost, ddc_labor_cost, ddc_equipment_cost, ddc_markup, ddc_unit_cost, ddc_avg_unit_price, ddc_extended_total_cost, bid1_qty, bid1_unit, bid1_material_cost, bid1_labor_cost, bid1_equipment_cost,
               bid1_unit_cost, bid1_ext_total_cost, bid1_variance, bid2_qty, bid2_unit, bid2_material_cost, bid2_labor_cost, bid2_equipment_cost, bid2_unit_cost, bid2_ext_total_cost, bid2_variance, bid3_qty, bid3_unit, bid3_material_cost, bid3_labor_cost, bid3_equipment_cost, bid3_unit_cost, bid3_ext_total_cost, bid3_variance)
    return csv_row


def process_excel_file_as_pd(data, project_id):
    # This is where we parse the EXCEL file.
    # Because of the nested structure, we need to keep track of CSI and CSI_SUB.
    # We know we expect 48 divisions.
    # We know that "insert row above" marks the end of a division.

    output_rows = []
    current_CSI_DIVISION = None
    current_CSI_SUB_DIVISION = None
    i = 1
    while i < len(data):
        row = data.iloc[i]
        csi_div = row['CSI DIVISION:']
        csi_sub_div = row['CSI SUB DIVISION:']

        if csi_div.startswith('DIVISION'):
            current_CSI_DIVISION = csi_div
            i = i + 1
            continue  # rows with DIVISION info don't contain data.

        if csi_sub_div != '':
            current_CSI_SUB_DIVISION = csi_sub_div
            i = i + 1
            continue  # rows with SUB DIVISION info don't contain data.

        if csi_div == 'Insert row above':
            if current_CSI_DIVISION is not None and current_CSI_DIVISION.startswith('DIVISION 48'):
                break
            current_CSI_DIVISION = None
            current_CSI_SUB_DIVISION = None
            i = i + 2
            # rows with "Insert row above" don't contain data and the next one is blank.
            continue

        description_of_work = row['DESCRIPTION OF WORK']
        if description_of_work is not '':
            # A row with a non-empty `DESCRIPTION OF WORK` contains valid data ==> CSV.
            csv_row = generate_csv_row(
                project_id, current_CSI_DIVISION, current_CSI_SUB_DIVISION, row)
            output_rows.append(csv_row)

        i = i + 1  # We go to the next row.
    return output_rows
