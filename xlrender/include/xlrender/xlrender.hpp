#pragma once

#include <pybind11/pybind11.h>
#include <xlnt/xlnt.hpp>
#include <json/json.h>

namespace xlrender {
class XLNT_API xlrender
{
public:
	xlrender(const std::string& tplpath, const std::string& defpath);
	void put_area(const std::string& name,const std::string& datastr);
	void save(const std::string& path);
	//void close(void);

private:
	xlnt::workbook read_book;
	xlnt::workbook write_book;
	Json::Value mapping_properties;
	Json::Reader reader;
};
}


namespace py = pybind11;

PYBIND11_PLUGIN(xlrender) {
    py::module m("xlrender", "pybind11 xlrender plugin");
    py::class_<xlrender::xlrender>(m, "xlrender")
    		.def(py::init<const std::string &, const std::string &>())
			.def("put_area", &xlrender::xlrender::put_area, "A function which adds two numbers", py::arg("name") = "", py::arg("datastr") = "")
			.def("save", &xlrender::xlrender::save);
    return m.ptr();
}
