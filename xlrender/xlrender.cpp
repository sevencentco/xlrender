#include <iostream>
#include <fstream>
#include <json/json.h>
#include <xlrender/xlrender.hpp>
#include <xlnt/cell/index_types.hpp>

using namespace std;
namespace xlrender {
xlrender::xlrender(const std::string & tplpath, const std::string & defpath)
{
	ifstream ifs(defpath);
	//Json::Reader reader;
	reader.parse(ifs, mapping_properties);
	read_book.load(tplpath);

	//coppy properties
	xlnt::worksheet read_sheet = read_book.active_sheet();
	xlnt::worksheet write_sheet = write_book.active_sheet();

	if(read_sheet.has_page_setup()){
		write_sheet.page_setup( read_sheet.page_setup());
	}
	if(read_sheet.has_page_margins()){
		write_sheet.page_margins( read_sheet.page_margins());
	}

	//write_book
}
void xlrender::put_area(const std::string& name,const std::string& datastr){

	//data
	Json::Value data;
	if(!(datastr.empty())){
		reader.parse(datastr, data);
	}

	xlnt::worksheet read_sheet = read_book.active_sheet();
	xlnt::worksheet write_sheet = write_book.active_sheet();
	unsigned int start_row = 0;
	unsigned int end_row = 0;

	Json::Value parameters;

	xlnt::row_t max_row = write_sheet.highest_row();
	xlnt::column_t ws_max_column = write_sheet.highest_column();

	bool last_row_has_data = false;

	//find area
	if(mapping_properties.isArray()){
		for (unsigned int i = 0; i < mapping_properties.size(); i++){
			//std::cout << "areaname name "<< mapping_properties[i]["area_name"].asString() << " "<< name << " "<< mapping_properties.size()<< " "<< i <<std::endl;
			if (mapping_properties[i]["area_name"].asString() == name){
				start_row = mapping_properties[i]["start_row"].asUInt();
				end_row = mapping_properties[i]["end_row"].asUInt();
				parameters = mapping_properties[i]["parameters"];
				//std::cout << "start_row "<< start_row << " "<<end_row <<std::endl;
				break;
			}
		}
	}

	if((start_row < 1)|| (end_row < 1) || (end_row < start_row)){
		return;
	}

	for(xlnt::column_t wfirstcolumn = 1; wfirstcolumn < ws_max_column; wfirstcolumn++) {
		//xlnt::cell wfcell = write_sheet.cell(wfirstcolumn, max_row);
		if(write_sheet.cell(wfirstcolumn, max_row).has_value() || write_sheet.cell(wfirstcolumn, max_row).has_format()){
			last_row_has_data = true;
			break;
		}
	}
	if(last_row_has_data){
		//std::cout << "Has not data" << std::endl;
		max_row = max_row + 1;
	}
	//std::cout << "max_row " << name<< ": " <<  max_row << " "<<last_row_has_data << " "<< write_sheet.highest_row()<< std::endl;

	//process
	xlnt::column_t last_column = read_sheet.highest_column();

	std::vector<xlnt::range_reference> merged_ranges = read_sheet.merged_ranges();

	for(xlnt::row_t row = start_row; row <= end_row; row++) {
		//set row height
		write_sheet.add_row_properties(max_row,read_sheet.row_properties(row));

		for(xlnt::column_t column = 1; column < last_column; column++) {
			//set column properties
			write_sheet.add_column_properties(column, read_sheet.column_properties(column));
			//std::cout << "Read sheet column: " << column.column_string() << " " <<read_sheet.column_width(column)  << std::endl;
			//std::cout << "Write sheet column: "<< column.column_string() << " " << write_sheet.column_width(column)  << std::endl;
			//xlnt::cell readcell = read_sheet.cell(column, row);
			xlnt::cell_type cell_datatype =  read_sheet.cell(column, row).data_type();

			bool ignore = false;

			if(read_sheet.cell(column, row).is_merged()){
				for(unsigned int i = 0; i < merged_ranges.size(); i++){
					xlnt::range_reference mrange_ref = merged_ranges.at(i);
					xlnt::range mrange = read_sheet.range(mrange_ref);
					if(mrange.contains(read_sheet.cell(column, row).reference())){
						//std::cout << mrange_ref.to_string() << std::endl;
						if(mrange_ref.top_left() == read_sheet.cell(column, row).reference()){
							xlnt::range_reference wmrange = xlnt::range_reference(mrange_ref.top_left().column(),
									max_row,
									mrange_ref.bottom_right().column(),
									mrange_ref.bottom_right().row() - mrange_ref.top_left().row() + max_row);
							write_sheet.merge_cells(wmrange);
						}else{
							ignore = true;
						}
						break;
					}
				}
			}

			if(!ignore){
				bool isHasParam = false;
				if(!(data.isNull()) && parameters.isArray()){
					for(unsigned int i = 0; i < parameters.size(); i++){
						//std::cout << readcell.reference().to_string() << std::endl;
						if (parameters[i]["position"] == read_sheet.cell(column, row).reference().to_string()){
							Json::Value dataVal = data[parameters[i]["name"].asString()];
							if(dataVal.isNull()){

							}
							else if(dataVal.isString()){
								//std::cout << dataVal.asString() << std::endl;
								write_sheet.cell(column, max_row).value(dataVal.asString());
							}else{
								write_sheet.cell(column, max_row).value(dataVal.asString(), true);
							}
							isHasParam = true;
							break;
						}
					}
				}
				if ((!isHasParam) && (read_sheet.cell(column, row).has_value())){
					if(cell_datatype == xlnt::cell_type::inline_string){
						write_sheet.cell(column, max_row).value(read_sheet.cell(column, row).value<xlnt::rich_text>());
					}else{
						write_sheet.cell(column, max_row).value(read_sheet.cell(column, row).value<std::string>(), true);
					}
				}
			}
			//set cell format
			if(read_sheet.cell(column, row).has_format()){
				write_sheet.cell(column, max_row).font(read_sheet.cell(column, row).font());
				write_sheet.cell(column, max_row).fill(read_sheet.cell(column, row).fill());
				write_sheet.cell(column, max_row).border(read_sheet.cell(column, row).border());
				write_sheet.cell(column, max_row).number_format(read_sheet.cell(column, row).number_format());


				//std::cout << readcell.reference().to_string() << std::endl;
				try {
					write_sheet.cell(column, max_row).alignment(read_sheet.cell(column, row).alignment());
				}catch(...) {
				   // code to handle any exception
				}
				try {
					write_sheet.cell(column, max_row).protection(read_sheet.cell(column, row).protection());
				}catch(...) {
				   // code to handle any exception
				}
			}
		}
		max_row = max_row + 1;
	}
	//std::cout << "highest_row " << write_sheet.highest_row()<< std::endl;
}

void xlrender::save(const std::string& path){
	write_book.save(path);
}

}
