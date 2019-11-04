set json = createobject("protium.json")

set dict = json.parse(	"{ width: '200', frame: false, height: 130, bodyStyle:'background-color: #ffffcc;',buttonAlign:'right', " & _
				"items: [{ xtype: 'form',  url: '/content.asp'},{ xtype: 'form2',  url: '/content2.asp'}] }")

wscript.echo dict("width")
wscript.echo dict("bodyStyle")
set dict2 = dict("items")
wscript.echo dict2.item(1).item("url")
wscript.echo dict2.item(2).item("url")
wscript.echo json.stringtojson("dog~cat|cow~" & 1)
set cows = createobject("scripting.dictionary")
cows.add "cow","moo"
cows.add "legs", 4
names = array( "菊花","范晓萱","伯大尼" )
cows.add "names",names
cows.add "dict", dict
wscript.echo json.toString(cows)

