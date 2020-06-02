import re

CSI_NUMBER_PATTERN = re.compile('[^a-zA-Z]+')


def get_csi_number(str):
    if str is None or str == "":
        return ""
    # first 6-8 digits, e.g.
    # - `23 05 93 Testing, Adjusting, and Balancing for HVAC`
    # - `26 41 13.13 Lightning Protection for Buildings`
    match = CSI_NUMBER_PATTERN.match(str)
    return match.group().strip()


def _float(s):
    return float(s) if s else 0.0


def generate_csv_row(project_id, markup, csi_division, csi_sub_division, row_data):
    """Function that extracts the data for a given row before outputing some CSV."""
    project_id = project_id
    csi_number = get_csi_number(csi_sub_division)
    csi_division = csi_division
    csi_subdivision = csi_sub_division
    category = csi_division
    subcategory = csi_subdivision
    item_code = row_data['RSMeans 12-digit code']
    activity = row_data['DESCRIPTION OF WORK']
    ddc_qty = _float(row_data['QUANT'])
    ddc_unit = row_data['UNIT']
    ddc_material_cost = _float(row_data['TOTAL MAT. $:'])
    ddc_labor_cost = _float(row_data['TOTAL LABOR $:'])
    # extra space for the column name. Typo?
    ddc_equipment_cost = _float(row_data['TOTAL  EQUIP $:'])
    ddc_markup = (ddc_material_cost + ddc_labor_cost +
                  ddc_equipment_cost) * markup
    ddc_unit_cost = (ddc_material_cost + ddc_labor_cost +
                     ddc_equipment_cost + ddc_markup) / ddc_qty if ddc_qty > 0 else ''
    ddc_avg_unit_price = None  # MUST BE COMPUTED AT THE END.
    ddc_extended_total_cost = ddc_qty * ddc_unit_cost if ddc_qty > 0 else ''
    bid1_qty = _float(row_data['QUANT.1'])
    bid1_unit = row_data['UNIT COST.1']
    bid1_material_cost = _float(row_data['TOTAL MAT. $:.1'])
    bid1_labor_cost = _float(row_data['TOTAL LABOR $:.1'])
    bid1_equipment_cost = _float(row_data['TOTAL  EQUIP $:.1'])
    bid1_unit_cost = (bid1_material_cost + bid1_labor_cost +
                      bid1_equipment_cost) / bid1_qty if bid1_qty > 0 else ''
    bid1_ext_total_cost = bid1_qty * bid1_unit_cost if bid1_qty > 0 else ''
    bid1_variance = ''
    try:
        bid1_variance = (bid1_ext_total_cost -
                         ddc_extended_total_cost) / ddc_extended_total_cost
    except Exception:
        pass

    bid2_qty = _float(row_data['QUANT.2'])
    bid2_unit = row_data['UNIT COST.2']
    bid2_material_cost = _float(row_data['TOTAL MAT. $:.2'])
    bid2_labor_cost = _float(row_data['TOTAL LABOR $:.2'])
    bid2_equipment_cost = _float(row_data['TOTAL  EQUIP $:.2'])
    bid2_unit_cost = (bid2_material_cost + bid2_labor_cost +
                      bid2_equipment_cost) / bid2_qty if bid2_qty > 0 else ''
    bid2_ext_total_cost = bid2_qty * bid2_unit_cost if bid2_qty else ''
    bid2_variance = ''
    try:
        bid2_variance = (bid2_ext_total_cost -
                         ddc_extended_total_cost) / ddc_extended_total_cost
    except Exception:
        pass

    bid3_qty = _float(row_data['QUANT.3'])
    bid3_unit = row_data['UNIT COST.3']
    bid3_material_cost = _float(row_data['TOTAL MAT. $:.3'])
    bid3_labor_cost = _float(row_data['TOTAL LABOR $:.3'])
    bid3_equipment_cost = _float(row_data['TOTAL  EQUIP $:.3'])
    bid3_unit_cost = (bid3_material_cost + bid3_labor_cost +
                      bid3_equipment_cost) / bid3_qty if bid3_qty > 0 else ''
    bid3_ext_total_cost = bid3_qty * bid3_unit_cost if bid3_unit_cost != '' else ''
    bid3_variance = ''
    try:
        bid3_variance = (bid3_ext_total_cost -
                         ddc_extended_total_cost) / ddc_extended_total_cost
    except Exception:
        pass

    try:
        ddc_avg_unit_price = 1.0/3 * \
            (bid1_unit_cost + bid2_unit_cost + bid3_unit_cost)
    except Exception:
        ddc_avg_unit_price = ''

    csv_row = (project_id, csi_number, csi_division, csi_subdivision, category, subcategory, item_code, activity, ddc_qty, ddc_unit, ddc_material_cost, ddc_labor_cost, ddc_equipment_cost, ddc_markup, ddc_unit_cost, ddc_avg_unit_price, ddc_extended_total_cost, bid1_qty, bid1_unit, bid1_material_cost, bid1_labor_cost, bid1_equipment_cost,
               bid1_unit_cost, bid1_ext_total_cost, bid1_variance, bid2_qty, bid2_unit, bid2_material_cost, bid2_labor_cost, bid2_equipment_cost, bid2_unit_cost, bid2_ext_total_cost, bid2_variance, bid3_qty, bid3_unit, bid3_material_cost, bid3_labor_cost, bid3_equipment_cost, bid3_unit_cost, bid3_ext_total_cost, bid3_variance)
    return csv_row


def process_excel_file_as_pd(data, project_id):
    # This is where we parse the EXCEL file.
    # Because of the nested structure, we need to keep track of CSI and CSI_SUB.
    # We know we expect 48 divisions.
    # We know that "insert row above" marks the end of a division.

    markup = float(data.iloc[0]['MARK-UP'])
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
            # some SUB DIVISION contain data.

        if csi_div == 'Insert row above':
            if current_CSI_DIVISION is not None and current_CSI_DIVISION.startswith('DIVISION 48'):
                break
            current_CSI_DIVISION = None
            current_CSI_SUB_DIVISION = None
            i = i + 2
            # rows with "Insert row above" don't contain data and the next one is blank.
            continue

        description_of_work = row['DESCRIPTION OF WORK']
        quant = row['QUANT']
        if description_of_work is not '' or (quant not in ['', 'SUB TOTAL']):
            # A row with a non-empty `DESCRIPTION OF WORK` or a `QUANT` value not 'SUB TOTAL' contains valid data ==> CSV.
            csv_row = generate_csv_row(
                project_id, markup, current_CSI_DIVISION, current_CSI_SUB_DIVISION, row)
            output_rows.append(csv_row)

        i = i + 1  # We go to the next row.
    return output_rows
