import std.stdio;
import gtk.Box;
import gtk.Label;
import gtk.ComboBoxText;
import gtk.Main;
import gtk.MainWindow;
import gtk.Widget;
import gtk.Button;
import gtk.TreeView;
import gtk.ListStore;
import gtk.TreeViewColumn;
import gtk.CellRendererText;
import gtk.TreeIter;
import gtk.FileChooserDialog;

import gdk.Event;

void main(string[] args) {
    
    Main.init(args);
    ZDGWindow w = new ZDGWindow("Zug data GUI", args);
    Main.run();
}

class ZDGWindow : MainWindow {
    int width  = 640;
    int height = 480;
    Box main_container;
    VboxFormContainer vbox_form;
    string[] headers;

    Button export_button;

    this(string title, string[] args) {
        super(title);

        this.setSizeRequest(this.width, this.height);
        this.addOnDestroy(delegate void (Widget w) => Main.quit());

        main_container = new Box(Orientation.VERTICAL, 5);
        this.add(main_container);

        Button button = new Button("Choose file", &choose_file_to_pivot);
        main_container.packStart(button, false, false, 5);

        export_button = new Button("Export to file", &choose_file_and_export);
        main_container.packEnd(export_button, false, false, 5);
        this.showAll();

    }

    void display_headers_form(string file_path, string[] headers) {
        if (this.vbox_form) {
            this.main_container.remove(this.vbox_form);
        }
        this.vbox_form = new VboxFormContainer(file_path, headers);
        main_container.packStart(vbox_form, false, false, 5);
        this.showAll();
    }

    void choose_file_to_pivot(Button button) {
        import std.file: exists;
        import gtk.MessageDialog;

        string[] buttons_labels = ["Openx", "Cancelx"];
        ResponseType[] response_types = [ResponseType.OK, ResponseType.CANCEL];
        auto file_chooser = new FileChooserDialog( 
            "Pick a file", 
            this, // parent window
            FileChooserAction.OPEN, 
            buttons_labels,
            response_types
        );
        file_chooser.setSelectMultiple(false);
        ResponseType response = cast(ResponseType) file_chooser.run();
        if ( response == ResponseType.OK ) {
            string file_path = file_chooser.getFilename();
            if (file_path && file_path.exists() ) {
                this.headers = get_xlsx_headers(file_path);
                this.display_headers_form(file_path, headers);
            } else {
                MessageDialog error_dialog = new MessageDialog(
                    this,
                    GtkDialogFlags.MODAL,
                    MessageType.ERROR,
                    ButtonsType.OK,
                    "No file selected or could not find the file."
                );
                int reponse = error_dialog.run();
                error_dialog.destroy();
                writeln("failed to find file ", file_path, " response: ", response);
            }
        } else {
            writeln("canceled ... ");
        }

        file_chooser.hide();
    }

    void choose_file_and_export(Button button) {
        import std.file: exists;
        import gtk.MessageDialog;
        import gtk.FileChooserDialog;

        string[] buttons_labels = ["Save", "Cancel"];
        ResponseType[] response_types = [ResponseType.OK, ResponseType.CANCEL];
        auto file_chooser = new FileChooserDialog( 
            "Pick a file to export to", 
            this, // parent window
            FileChooserAction.SAVE, 
            buttons_labels,
            response_types
        );
        file_chooser.setSelectMultiple(false);
        ResponseType response = cast(ResponseType) file_chooser.run();

        if ( response == ResponseType.OK ) {
            writeln("export to file ResponseType.OK");
            string file_path = file_chooser.getFilename();
            auto data = this.get_pivoted_data();
            export_to_csv(file_path, data);            
        } else {
           writeln("export to file ResponseType NOT OK");
        }

        file_chooser.hide();
    }

    string[][] get_pivoted_data() {
        import std.algorithm: map, canFind, sort, SwapStrategy;
        import std.typecons: Tuple;
        import std.array: array, join;

        string row_header = this.vbox_form.columns_in_file[this.vbox_form.combo_row.getActive()];
        string data_header = this.vbox_form.data_columns[this.vbox_form.combo_data.getActive()];
        string[] columns_headers = this.vbox_form.combo_columns.get_active().map!(a => this.vbox_form.columns_to_merge[a]).array();
        string[][string][string] lookup_table;
        int[string] lookup_first_column;
        int[string] lookup_headers;
        string[][] data = read_xlsx_file(this.vbox_form.file_path);
        
        string[] headers = this.vbox_form.columns_in_file;
        size_t row_header_index = 0;
        size_t data_header_index = 0;
        size_t[] column_header_indices;
        for (size_t i = 0; i < headers.length; i++) {
             if (headers[i] == row_header) { 
                row_header_index = i;
            }
            
            if (headers[i] == data_header) { 
                data_header_index = i;
            }

            if (columns_headers.canFind(headers[i])) {
                column_header_indices ~= i;
            }
        }
        
        writeln("row_header_index: ", row_header, " ", row_header_index, " data_header_index: ", data_header, " ", data_header_index);
        int row_no = 0;
        foreach (string[] row; data) {
            import std.string: join;
            import std.array: array;
            import std.datetime: Date, Duration, dur;
            import std.conv: to;

            auto msexcel_base_date = Date(1899, 12, 30);

            string column_values_merged = column_header_indices.map!(a => row[a]).array().join(", ");
            writeln(row_no, " ", row[data_header_index]);
            int excel_date = row[data_header_index].to!float.to!int;
            Duration excel_duration = dur!"days"(excel_date);
            lookup_first_column[row[row_header_index]] = 1;
            lookup_headers[column_values_merged] = 1;
            lookup_table[row[row_header_index]][column_values_merged] ~= (msexcel_base_date + excel_duration).toISOExtString();
            row_no++;
        }
       writeln("......................................................");

        string[] students = lookup_first_column.keys.sort!("b > a").array();
        string[] classes = lookup_headers.keys.sort!("b > a").array();
        string[][] result;
        string[] first_row = ["Cursanti"];
        first_row ~= classes;
        result ~= first_row;
        foreach (student; students) {
            string[] row = [student];

            foreach (c; classes) {
                if (c in lookup_table[student] ) {
                    row ~=  lookup_table[student][c].join("\n");
                } else {
                    row ~= "";
                }
            }
            result ~= row;
        }

        return result;
    }
}

string[][] read_xlsx_file(string file_path) {
    import std.algorithm: map;
    import std.array: array;
    import xlsxreader;

    SheetNameId[] sheets = file_path.sheetNames();
    auto sheet = file_path.readSheet(sheets[0].name);
    auto table = sheet.table;
    string[] headers = table[0].map!(a => a.convertToString).array();
    string[][] data;
    for (size_t i = 1; i < table.length; i++) {
        data ~= table[i].map!(a => a.convertToString).array();
    }
    return data;
}

void export_to_csv(string file_path, string[][] data) {
    import std.array: join, array;
    import std.algorithm: map;
    import std.file: write;

    string separator = `|`;
    string[] result;
    foreach (row; data) {
        result ~= row.map!(a => `"` ~ a ~ `"`).array.join(separator);
    }

    write(file_path, result.join("\n"));
}


// void export_to_xlsx(string file_path, string[][] data) {
//     import libxlsxd;
//     writeln("exporting to xlsx: ", file_path, "; row count: ", data.length, "; row lenght: ", data[0].length);
//     auto workbook  = new Workbook(file_path);
//     auto worksheet = workbook.addWorksheet("Pivoted data");
//     auto format = workbook.addFormat();
//     format.setTextWrap();
//     auto headers = data[0];
//     ushort first_column = 0;
//     ushort last_column = cast(ushort) headers.length;
//     worksheet.setColumn(first_column, last_column, 30, format);

//     for (uint y = 1; y < data.length; y++ ) {
//         for (ushort x = 0; x < data[0].length; x++) {
//             writeln(y, " ", x);
//             worksheet.writeString(y, x, data[y][x]);
//         }
//     }
//     lxw_error result = workbook.close();
//     writeln("exported to xlsx: ", file_path, "; row count: ", data.length, "; row lenght: ", data[0].length, "; result: ", lxw_strerror(result));
// }

string[] get_xlsx_headers(string file_path) {
    import std.algorithm: map;
    import std.array: array;
    import xlsxreader;

    SheetNameId[] sheets = file_path.sheetNames();
    auto sheet = file_path.readSheet(sheets[0].name);
    string[] headers = sheet.table[0].map!(a => a.convertToString).array();

    return headers;
}

class VboxFormContainer: Box {
    CustomCombo combo_row;
    string[] columns_in_file = []; // columns to pick the row from
    CustomCombo combo_data;
    string[] data_columns = []; // columns to pick the data from 
    ColumnsTreeView combo_columns;
    string[] columns_to_merge = []; // columns to pick the new pivot table columns from
    TreeIter columns_to_merge_iter;
    ListStore combo_columns_store;
    Box hbox_columns;

    string file_path;
    bool data_is_date = true; // if the data columns contains a date, i.e. 2021-10-21

    this(string path, string[] column_names) {
        import gtk.ScrolledWindow;
        super(Orientation.VERTICAL, 5);

        this.file_path = path;
        this.columns_in_file = column_names;
        
        //// Row, Columns, Data

        // Row
        auto hbox_row = new Box(Orientation.HORIZONTAL, 5);
        auto label_row = new Label("Row");
        hbox_row.packStart(label_row, false, false, 5);
        this.combo_row = new CustomCombo();
        this.set_pivot_row();
        this.combo_row.on_change = () { set_data_columns(); set_data_row(); };
        

        hbox_row.packStart(combo_row, false, false, 5);
        this.packStart(hbox_row, false, false, 5);

        // Data
        auto hbox_data = new Box(Orientation.HORIZONTAL, 5);
        auto label_data = new Label("Data");
        hbox_data.packStart(label_data, false, false, 5);

        this.combo_data = new CustomCombo();
        this.combo_data.on_change = () { set_combo_columns_list(); set_columns_tree_view(); };

        hbox_data.packStart(combo_data, false, false, 5);
        this.packStart(hbox_data, false, false, 5);


        // Columns
        hbox_columns = new Box(Orientation.HORIZONTAL, 5);
        auto label_columns = new Label("Columns");
        hbox_columns.packStart(label_columns, false, false, 5);
        this.packStart(hbox_columns, false, false, 5);

	auto scrolled_columns_window = new ScrolledWindow();
        scrolled_columns_window.setMinContentHeight(280);
        hbox_columns.packStart(scrolled_columns_window, true, true, 5);
        this.combo_columns = new ColumnsTreeView(this.columns_to_merge);
//        hbox_columns.packStart(this.combo_columns, false, false, 5);
	scrolled_columns_window.add(this.combo_columns);
    }

    void set_pivot_row() {
        foreach (size_t i; 0..this.columns_in_file.length) {
            this.combo_row.appendText(this.columns_in_file[i]);
        }
    }
    
    void set_data_columns() {
        import std.algorithm: filter;
        import std.array: array;

        string selected_row = this.columns_in_file[this.combo_row.getActive()];
        writeln("selected row: ", selected_row);

        this.data_columns = this.columns_in_file.filter!(a => a != selected_row).array();
    }

    void set_data_row() {
        import gtk.ListStore;

        this.combo_data.removeAll();
        this.combo_data.setActiveId(null);
        foreach (size_t i; 0..this.data_columns.length) {
            this.combo_data.appendText(this.data_columns[i]);
        }
        // this.export_button.setSensitive(false);
    }

    void set_combo_columns_list() {
        import std.algorithm: filter;
        import std.array: array;

        int data_row_id = this.combo_data.getActive();

        if (data_row_id == -1) {
            this.columns_to_merge = [];
        } else {
            string selected_data = this.data_columns[data_row_id];
            this.columns_to_merge = this.data_columns.filter!(a => a != selected_data).array;
        }
    }

    void set_columns_tree_view() {
        ColumnsListStore model = cast(ColumnsListStore) this.combo_columns.getModel();
        model.clear();
        model.column_names = this.columns_to_merge;
        model.set_values();
        // this.export_button.setSensitive(true);
    }
}

class CustomCombo : ComboBoxText {
    void delegate() on_change;

    this() {
        this.addOnChanged(&changed);
    }
 
    void changed(ComboBoxText cbt) {
        writeln("changed selected");
        if (this.on_change) {
            writeln("++ changed selected");
            this.on_change();
        }
    }   
}

class ColumnsTreeView: TreeView {
    CustomTreeViewColumn tree_view_column;
    ColumnsListStore list_store;

    this(string[] columns_to_show) {
        super();
	
	this.setHeadersVisible(false);

        list_store = new ColumnsListStore(columns_to_show);
        this.setModel(list_store);

        tree_view_column = new CustomTreeViewColumn;
        this.appendColumn(tree_view_column);

        auto combo_columns_selection = this.getSelection();
        combo_columns_selection.setMode(SelectionMode.MULTIPLE);
    }

    int[] get_active() {
        import gtk.TreeSelection;
        import gtk.TreeModel;
        import gtk.TreePath;
        import std.algorithm: map;
        import std.array: array;

        auto model = this.getModel();
        auto selection = this.getSelection();
        int[] rows = selection.getSelectedRows(model).map!(a => a.getIndices()[0]).array();

        return rows;
    }
}

class ColumnsListStore: ListStore {
    string[] column_names;
    TreeIter tree_iter;

    this(string[] columns_to_show) {
        super([GType.STRING]);
        // if (columns_to_show.length > 0) {
            this.column_names = columns_to_show;
        // }


        this.set_values();
    }

    void set_values() {
        foreach (size_t i; 0..this.column_names.length) {
            string column_name = this.column_names[i];
            writeln(column_name);
            this.tree_iter = this.createIter();
            this.setValue(this.tree_iter, 0, column_name);
        }
    }
}

class CustomTreeViewColumn: TreeViewColumn {
    import gtk.CellRendererText;

    CellRendererText cell_renderer_text;
    string title = "Available columns";
    string attribute_type = "text";
    int column_number = 0;

    this() {
        this.cell_renderer_text = new CellRendererText();
        super(this.title, this.cell_renderer_text, this.attribute_type, this.column_number);
    }
}
