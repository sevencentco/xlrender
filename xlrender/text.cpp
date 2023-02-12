#include <iostream>
#include <fstream>
#include <xlnt/xlnt.hpp>
#include <xlnt/cell/index_types.hpp>
#include <json/json.h>
#include <xlrender/xlrender.hpp>

using namespace std;

void jsontest(){
	ifstream ifs("def.json");
	Json::Reader reader;
	Json::Value obj;
	reader.parse(ifs, obj); // reader can also read strings
	//cout << "Book: " << obj["book"].asString() << endl;
	//cout << "Year: " << obj["year"].asUInt() << endl;
	//const Json::Value& characters = obj["characters"]; // array of characters
	/*for (int i = 0; i < characters.size(); i++){
		cout << "    name: " << characters[i]["name"].asString();
		cout << " chapter: " << characters[i]["chapter"].asUInt();
		cout << endl;
	}*/

	//for key - val
	/*Json::Value val;
	reader.parse("{\"one\":1,\"two\":2,\"three\":3}",  val);
	for (Json::Value::const_iterator it=val.begin(); it!=val.end(); ++it)
	    cout << it.key().asString() << ':' << it->asInt() << '\n';*/

	for (int i = 0; i < obj.size(); i++){
		cout << "    name: " << obj[i]["area_name"].asString();
	}
}

/*void initDeffile(std::string& path){
	cout << "    path: " << path << endl;
}*/


int main(){

	//std::string data("{\"chuquan\":\"ABC CDE\", \"tendonvi\":\"123456\"}");

	Json::Reader reader;

	std::string tplpath("./thongketheodonvi.xlsx");
	std::string defpath("thongketheodonvi.json");

	std::string datapath("data.json");
	ifstream dataifs(datapath);


	//std::ifstream ifs("myfile.txt");
	std::string content( (std::istreambuf_iterator<char>(dataifs) ),
	                       (std::istreambuf_iterator<char>()    ) );




	std::string outpath("output.xlsx");
	xlrender::xlrender xl(tplpath, defpath);

	//std::string data("{\"ma\": \"So: so1\",\"chuquan\": null,\"kybaocao\": 1,\"tendonvi\": \"So Yt HN\",\"ngaybaocao\": \"....., ngay 9 thang 3 nam 2017\",\"tungaydenngay\": \"(Tu ngay 1 tang 1 nnam 2017 den ngay 1 thang 1 nam 2017)\"}");
	//Json::Value dataJson;
	//reader.parse(dataifs, dataJson);

	//dataJson.asString();

	std::string name("bangnguoiheader");
	std::string name2("bangnguoickchitiet");
	//xl.put_area(name,"");
	xl.put_area(name2, content);
	xl.save(outpath);
	//xl.close();
}


