def prepare_form (document, vip_name_title, pax, country, flight, date, time):
    vip_country_table = document.tables[0]
    row = vip_country_table.rows[1]
    vip_name_title_cell = row.cells[0]
    country_cell = row.cells[2]

    flight_info_table = document.tables[1]
    row = flight_info_table.rows[1]
    date_cell = row.cells[2]
    time_cell = row.cells[4]
    flight_cell = row.cells[6]

    if pax != "0":
        vip_name_title_cell.text = vip_name_title +" + " + pax + " Pax."
    else:
        vip_name_title_cell.text = vip_name_title

    country_cell.text = country
    country_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    date_cell.text = date
    date_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    time_cell.text = time
    time_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER

    flight_cell.text = flight
    flight_cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER


def add_drivers(document, limo_df):
    limo_table = document.tables[3]

    if len(limo_df) > 15:
        print("There are more than 15 limousines in the list. Which is not allowed GVA \nPlease check the list and try again.")
        exit()

    for i in range(0, len(limo_df)):
        row = limo_table.rows[i+1]
        limo_cell = row.cells[0]
        driver_cell = row.cells[1]
        limo_cell.text = limo_df.iloc[i, 0]
        driver_cell.text = limo_df.iloc[i, 1]

        
