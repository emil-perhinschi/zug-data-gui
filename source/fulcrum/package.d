module fulcrum;

struct FulcrumData {
    string[] columns;
}

void load_xlsx_file(string file_path);
void identify_file_type(string file_path);
void list_columns(FulcrumData);