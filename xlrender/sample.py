import xlrender

from sanic import Sanic
from sanic.response import json
import ujson

app = Sanic()

@app.route("/")
async def test(request):
    filename = "test.xlsx"
    deffile = "sample.json"
    template = "sample.xlsx"
    render = xlrender.xlrender(template, deffile)
    data = {
        "data": []
    }

    headerobj = {
        "chuquan": "HANOI",
        "tendonvi": "ABC",
        #"ngaybaocao": u"..., ngày " + "" + " tháng " + "" + " năm " + str(baocao.nambaocao)
        "ngaybaocao": '....., ngày %d tháng %d năm %d' % (1,2,2020),
        "tungaydenngay": ""
    }
    
    render.put_area("header", ujson.dumps( headerobj))
    render.save(filename)
    return json({"hello": "world"})

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)