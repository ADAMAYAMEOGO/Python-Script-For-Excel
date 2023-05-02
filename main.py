from scripts import  add_string_to_column,excel_to_csv,read_column_from_excel,split_series_to_excel_files


if __name__ == "__main__":
    input_file = 'buls_whatsapp_contact_part_2.xlsx'
    output_file = 'buls_whatsapp_contact_part_2.csv'
    data= read_column_from_excel("buls_whatsapp_contact_part_2.xlsx", "WhatsApp Number(with country code)")
    data_add= add_string_to_column(data,"+")
    excel_to_csv(input_file, output_file)
